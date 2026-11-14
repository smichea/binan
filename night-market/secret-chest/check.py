#!/usr/bin/env python3
"""Export the chest STLs and check that the mechanism works on paper.

    python3 check.py            # export stl/*.stl and run every check
    python3 check.py --no-stl   # checks only

Needs: openscad on PATH, and  pip install trimesh manifold3d numpy

The checks move the real meshes through their motions and look for collisions:
  * each printable part is a closed solid that sits flat on the bed
  * closed chest: no two parts overlap
  * lid: swings freely, then stops against the body and stays open
  * drawer: slides out freely once the ball is out of the way
  * lock: with the ball in its seat the drawer is held; tilting frees it;
    pushing the drawer back in moves the ball aside and it re-locks by itself

Friction and print tolerances are not modelled - print one and try it.
"""
import math
import re
import subprocess
import sys
import tempfile
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

import manifold3d as m3
import numpy as np
import trimesh

HERE = Path(__file__).resolve().parent
SCAD = HERE / "chest.scad"
PARTS = ["body", "lid", "plinth", "drawer", "lid_strap", "body_strap", "lock_plate"]
TWICE = ("lid_strap", "body_strap")
PLATES = ["plate_a", "plate_b"]
BED = 220
PLA_G_PER_CM3 = 1.24
TOUCH = 0.05  # mm3 of overlap we treat as "just touching"

failures = []


def report(ok, text):
    print(("  ok   " if ok else "  FAIL ") + text)
    if not ok:
        failures.append(text)


def openscad(out, **defs):
    cmd = ["openscad", "-o", str(out)]
    for k, v in defs.items():
        cmd += ["-D", f'{k}="{v}"' if isinstance(v, str) else f"{k}={str(v).lower()}"]
    subprocess.run(cmd + [str(SCAD)], check=True, capture_output=True)


def solid(mesh):
    return m3.Manifold(m3.Mesh(vert_properties=np.asarray(mesh.vertices, dtype=np.float32),
                               tri_verts=np.asarray(mesh.faces, dtype=np.uint32)))


def overlap(a, b):
    """Volume shared by two solids."""
    return (a ^ b).volume()


def main():
    export_stl = "--no-stl" not in sys.argv
    tmp = Path(tempfile.mkdtemp(prefix="chest-"))

    # ---- export -----------------------------------------------------------
    info_file = tmp / "info.echo"
    openscad(info_file, part="info")
    info = dict(re.findall(r"INFO (\w+)=([^\"]+)", info_file.read_text()))
    num = {k: float(v) for k, v in info.items() if re.fullmatch(r"-?[\d.]+", v)}

    jobs = [(tmp / f"{p}.stl", dict(part=p, in_place=True)) for p in PARTS]
    if export_stl:
        (HERE / "stl").mkdir(exist_ok=True)
        jobs += [(HERE / "stl" / f"{p}.stl", dict(part=p)) for p in PARTS + PLATES]
    with ThreadPoolExecutor(8) as pool:
        list(pool.map(lambda j: openscad(j[0], **j[1]), jobs))
    if export_stl:
        for p in PARTS + PLATES:  # OpenSCAD writes text STL; binary is about 5x smaller
            path = HERE / "stl" / f"{p}.stl"
            trimesh.load_mesh(path, process=False).export(path)

    mesh = {p: trimesh.load_mesh(tmp / f"{p}.stl", process=True) for p in PARTS}
    asm = {p: solid(m) for p, m in mesh.items()}

    print(f"\nOutside size      {info['outer']} mm")
    print(f"Main compartment  {info['body_inner']} mm (plus the lid)")
    print(f"Secret drawer     {info['drawer_inner']} mm")

    # ---- 1. printable parts ----------------------------------------------
    print("\nParts")
    total = 0
    for p in PARTS:
        m = mesh[p]
        grams = m.volume / 1000 * PLA_G_PER_CM3
        n = 2 if p in TWICE else 1
        total += grams * n
        ok = m.is_watertight and m.body_count == 1 and asm[p].status() == m3.Error.NoError
        line = f"{p:11s} closed solid, {grams:4.1f} g" + (f" (x{n})" if n > 1 else "")
        if export_stl:
            pm = trimesh.load_mesh(HERE / "stl" / f"{p}.stl", process=True)
            e = pm.extents
            ok = ok and pm.is_watertight and abs(pm.bounds[0][2]) < 0.02
            line += f", prints flat on the bed, {e[0]:.0f} x {e[1]:.0f} x {e[2]:.0f} mm"
        report(ok, line)
    print(f"       about {total:.0f} g of PLA if printed solid; less with infill")
    if export_stl:
        for p in PLATES:
            e = trimesh.load_mesh(HERE / "stl" / f"{p}.stl").extents
            report(e[0] < BED - 10 and e[1] < BED - 10, f"{p} fits the bed: {e[0]:.0f} x {e[1]:.0f} mm")

    # ---- 2. closed chest ---------------------------------------------------
    print("\nClosed chest")
    full = dict(asm)
    for p in TWICE:
        full[p + " (left)"] = asm[p].mirror((1, 0, 0))
    names = list(full)
    worst = 0
    for i, a in enumerate(names):
        for b in names[i + 1:]:
            v = overlap(full[a], full[b])
            worst = max(worst, v)
            if v > TOUCH:
                report(False, f"{a} overlaps {b} by {v:.2f} mm3")
    report(worst <= TOUCH, f"no two parts overlap (largest shared volume {worst:.3f} mm3)")

    # ---- 3. lid ------------------------------------------------------------
    print("\nLid")
    ya, za = num["ya"], num["za"]

    def about_hinge(deg):
        return trimesh.transformations.rotation_matrix(-math.radians(deg), [1, 0, 0], [0, ya, za])

    def lid_at(deg):
        return asm["lid"].transform(about_hinge(deg)[:3, :])

    stop = next((d for d in np.arange(1, 181, 1.0) if overlap(lid_at(d), asm["body"]) > TOUCH), None)
    report(stop is not None and 95 <= stop <= 125,
           f"opens freely and stops against the body at about {stop:.0f} degrees")
    if stop:
        open_lid = mesh["lid"].copy()
        open_lid.apply_transform(about_hinge(stop - 1))
        lean = open_lid.center_mass[1] - ya
        report(lean > 2, f"stays open by itself (its weight sits {lean:.0f} mm behind the hinge)")
        rest = [mesh[p] for p in ("body", "plinth", "drawer")]
        mass = sum(m.volume for m in rest) + open_lid.volume
        cy = (sum(m.volume * m.center_mass[1] for m in rest) + open_lid.volume * open_lid.center_mass[1]) / mass
        margin = mesh["plinth"].bounds[1][1] - cy
        report(margin > 8, f"chest does not tip backwards with the lid open ({margin:.0f} mm of margin)")

    # ---- 4. drawer and lock -----------------------------------------------
    print("\nSecret drawer")
    r, tf = num["r"], num["tf"]
    slope = math.radians(num["slope_deg"])
    plinth, body, drawer = asm["plinth"], asm["body"], asm["drawer"]
    xs, yc = num["x_seat"], num["y_trc"]
    x_end = num["x_tr0"] + r + 0.1
    sphere = m3.Manifold.sphere(r - 0.01, 64)

    def drawer_at(t):
        return drawer.translate((0, -t, 0))

    def ball_z(x):
        # the ball rides on the flat floor or on the slope, whichever is higher
        return max(tf + r, tf + (num["x_slope0"] - x) * math.tan(slope) + r / math.cos(slope))

    def free(x, y, d):
        b = sphere.translate((x, y, ball_z(x)))
        return overlap(b, plinth) <= TOUCH and overlap(b, d) <= TOUCH

    bite = max(overlap(drawer_at(t), plinth) + overlap(drawer_at(t), body) for t in np.arange(0, 46, 0.25))
    report(bite < 1.0, f"slides 45 mm with nothing in the way (the click bumps bite {bite:.2f} mm3 at most)")

    seated = sphere.translate((xs, yc, tf + r - num["seat_sink"]))
    report(overlap(seated, drawer) <= TOUCH and overlap(seated, plinth) < 0.3,
           "locked ball sits in its dimple, clear of the drawer")

    step = 0.05

    def push_out(x, y, d):
        """Ball is being squeezed: move it to the nearest free spot, if any."""
        best = None
        for dx in np.arange(-0.8, 0.81, step):
            for dy in np.arange(-0.8, 0.81, step):
                if free(x + dx, y + dy, d):
                    dist = math.hypot(dx, dy)
                    if best is None or dist < best[0]:
                        best = (dist, x + dx, y + dy)
        return best and best[1:]

    def roll(x, y, d, target):
        """Let gravity roll the ball along the track towards target as far as it can."""
        sign = 1 if target > x else -1
        while abs(target - x) > 1e-6:
            nx = x + sign * min(step, abs(target - x))
            for dy in (0, step, -step, 2 * step, -2 * step):
                if free(nx, y + dy, d):
                    x, y = nx, y + dy
                    break
            else:
                break
        return x, y

    def pull(x, y):
        """Chest upright: pull the drawer and return how far it gets."""
        for t in np.arange(0, 45.01, 0.2):
            d = drawer_at(t)
            if not free(x, y, d):
                spot = push_out(x, y, d)
                if spot is None:
                    return t
                x, y = spot
            x, y = roll(x, y, d, xs)
        return None

    held = pull(xs, yc)
    report(held is not None and held <= 2.0,
           f"chest upright: drawer is held after {held:.1f} mm of play" if held is not None
           else "chest upright: the drawer pulls out - not locked")

    x, y = roll(xs, yc, drawer, x_end)
    report(abs(x - x_end) < 0.11, f"chest tilted left: ball rolls {xs - x:.1f} mm to the end of its track")
    report(all(free(x, y, drawer_at(t)) for t in np.arange(0, 45.01, 0.25)),
           "chest tilted left: drawer pulls out freely")

    print("\nClosing the drawer with the chest upright")
    x, y = xs, yc
    jam = None
    leftmost = x
    relock_at = None
    for t in np.arange(30, -0.001, -0.1):
        t = max(float(t), 0.0)
        d = drawer_at(t)
        if not free(x, y, d):
            spot = push_out(x, y, d)
            if spot is None:
                jam = t
                break
            x, y = spot
        x, y = roll(x, y, d, xs)
        leftmost = min(leftmost, x)
        if relock_at is None and leftmost < xs - 1 and abs(x - xs) < 0.11:
            relock_at = t
    if jam is not None:
        report(False, f"ball jams with the drawer {jam:.1f} mm from closed")
    else:
        report(leftmost < num["xh0"] - r + 0.3,
               f"hook pushes the ball {xs - leftmost:.1f} mm aside without jamming")
        report(relock_at is not None,
               f"ball drops back into its seat once the drawer is within {relock_at:.1f} mm of home"
               if relock_at is not None else "ball does not return to its seat")
        again = pull(x, y)
        report(again is not None and again <= 2.0, "drawer is locked again")

    print()
    if failures:
        print(f"{len(failures)} check(s) FAILED")
        sys.exit(1)
    print("All checks passed.")


if __name__ == "__main__":
    main()
