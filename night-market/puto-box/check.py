#!/usr/bin/env python3
"""Export the puto box STLs and check them.

    python3 check.py

Needs: openscad on PATH, and  pip install trimesh manifold3d numpy networkx

Checks: every part is a closed solid that sits flat on the bed and needs no
supports; nothing overlaps when the box is assembled; the cake cannot be
pulled off the leaf but unscrews freely; the lid lifts off; the cheese goes in
from below and cannot pull out through the top; the plates fit a 220 mm bed.
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
SCAD = HERE / "puto.scad"
PARTS = ["base", "body", "lid", "cheese"]
PLATES = ["plate_white", "plate_green", "plate_yellow"]
BED = 220
MAX_HANG = 1.0   # mm a layer may hang beyond a 50 degree overhang (small bridges)
TOUCH = 0.05
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


def deepest_overhang(part, layer=0.2):
    """Slice the part the way a printer does and return how far (mm) any layer
    hangs out beyond what a 50 degree overhang from the layer below allows."""
    lo, hi = part.bounding_box()[2], part.bounding_box()[5]
    worst = 0.0
    below = part.slice(lo + layer/2)
    for z in np.arange(lo + 1.5*layer, hi, layer):
        here = part.slice(float(z))
        hanging = here - below.offset(layer*1.2, m3.JoinType.Round)
        d = 0.0
        while d < 5 and hanging.offset(-(d + 0.1)/2, m3.JoinType.Round).area() > 1e-3:
            d += 0.1
        worst = max(worst, d)
        below = here
    return worst


def main():
    tmp = Path(tempfile.mkdtemp(prefix="puto-"))
    stl = HERE / "stl"
    stl.mkdir(exist_ok=True)
    info_file = tmp / "info.echo"
    openscad(info_file, part="info")
    info = dict(re.findall(r"INFO (\w+)=([^\"]+)", info_file.read_text()))
    lead, thread_h = float(info["lead"]), float(info["thread_h"])

    jobs = [(stl / f"{p}.stl", dict(part=p)) for p in PARTS + PLATES]
    jobs += [(tmp / f"{p}.stl", dict(part=p, in_place=True)) for p in PARTS]
    with ThreadPoolExecutor(8) as pool:
        list(pool.map(lambda j: openscad(j[0], **j[1]), jobs))
    for path in stl.glob("*.stl"):  # OpenSCAD writes text STL; binary is much smaller
        trimesh.load_mesh(path, process=False).export(path)

    print(f"\nOutside size        {info['outer']} mm")
    print(f"Main compartment    {info['main_inner']} mm")
    print(f"Hidden compartment  {info['hidden_inner']} mm")

    print("\nParts")
    for p in PARTS:
        m = trimesh.load_mesh(stl / f"{p}.stl", process=True)
        e = m.extents
        ok = m.is_watertight and m.body_count == 1 and abs(m.bounds[0][2]) < 0.02
        report(ok, f"{p:9s} closed solid, flat on the bed, {e[0]:.0f} x {e[1]:.0f} x {e[2]:.0f} mm, "
                   f"{m.volume/1000*1.24:.1f} g solid")
        hang = deepest_overhang(solid(m))
        report(hang <= MAX_HANG, f"{p:9s} needs no supports (no layer hangs out more than {hang:.1f} mm)")

    asm = {p: solid(trimesh.load_mesh(tmp / f"{p}.stl", process=True)) for p in PARTS}

    print("\nAssembled")
    worst = 0
    for i, a in enumerate(PARTS):
        for b in PARTS[i + 1:]:
            v = (asm[a] ^ asm[b]).volume()
            worst = max(worst, v)
            if v > TOUCH:
                report(False, f"{a} overlaps {b} by {v:.2f} mm3")
    report(worst <= TOUCH, f"no two parts overlap (largest shared volume {worst:.3f} mm3)")

    print("\nScrew joint between cake and leaf")
    base, body = asm["base"], asm["body"]

    def unscrewed(deg, sense):
        """Cake turned by deg (sense +1 = anticlockwise seen from above) and raised to match."""
        return body.rotate((0, 0, sense * deg)).translate((0, 0, deg / 360 * lead))

    pulled = max((base ^ body.translate((0, 0, dz))).volume() for dz in (1.0, 2.0, 3.0))
    report(pulled > 1.0, "cake cannot be pulled straight off the leaf")
    turn_needed = thread_h / lead * 360
    free = {}
    for sense in (1, -1):
        free[sense] = max((base ^ unscrewed(d, sense)).volume() for d in np.arange(2, turn_needed + 60, 2))
    opens = [s for s in free if free[s] <= TOUCH]
    report(len(opens) == 1, "cake unscrews freely one way and not the other "
           f"(overlap: anticlockwise {free[1]:.2f}, clockwise {free[-1]:.2f} mm3)")
    if opens:
        print(f"       to open: turn the cake {'anticlockwise' if opens[0] == 1 else 'clockwise'} "
              f"about {turn_needed:.0f} degrees (seen from above)")

    print("\nLid and cheese")
    lid, top = asm["lid"], asm["cheese"]
    lifted = max((body ^ lid.translate((0, 0, dz))).volume() for dz in np.arange(0.5, 8, 0.5))
    report(lifted <= TOUCH, "lid lifts straight off the cake")
    held = min((lid ^ top.translate((0, 0, dz))).volume() for dz in (0.5, 1.0, 2.0))
    report(held > 1.0, "cheese cannot pull out through the top of the lid")
    drop = max((lid ^ top.translate((0, 0, -dz))).volume() for dz in (0.5, 1.5, 3.0))
    report(drop <= TOUCH, "cheese goes into the lid from below")

    print("\nPlates")
    for p in PLATES:
        e = trimesh.load_mesh(stl / f"{p}.stl").extents
        report(e[0] < BED - 10 and e[1] < BED - 10, f"{p} fits the bed: {e[0]:.0f} x {e[1]:.0f} mm")

    print()
    if failures:
        print(f"{len(failures)} check(s) FAILED")
        sys.exit(1)
    print("All checks passed.")


if __name__ == "__main__":
    main()
