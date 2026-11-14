#!/usr/bin/env python3
"""Export the keychain STLs and check them.

    python3 check.py

Needs: openscad on PATH, and  pip install trimesh manifold3d numpy networkx

Checks: every part is a closed solid that sits flat on the bed, strap and sole
do not overlap when clicked together, the hooks really catch under the sole,
nothing needs supports, and the batch plates fit a 220 mm bed.
Friction and print tolerances are not modelled - print the fit test first.
"""
import subprocess
import sys
import tempfile
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

import manifold3d as m3
import numpy as np
import trimesh

HERE = Path(__file__).resolve().parent
SCAD = HERE / "tsinelas.scad"
BED = 220
MAX_BRIDGE = 6.0   # mm; anything wider would want supports
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


def widest_overhang(mesh):
    """Narrow dimension of the widest surface that faces down without anything under it."""
    low = mesh.triangles[:, :, 2].min(axis=1)
    faces = np.where((mesh.face_normals[:, 2] < -0.72) & (low > 0.05))[0]
    if len(faces) == 0:
        return 0.0
    patches = mesh.submesh([faces], append=True, repair=False).split(only_watertight=False)
    return max(min(p.extents[0], p.extents[1]) for p in patches)


def main():
    tmp = Path(tempfile.mkdtemp(prefix="tsinelas-"))
    stl = HERE / "stl"
    stl.mkdir(exist_ok=True)
    sides = {"right": 1, "left": -1}
    jobs = []
    for name, f in sides.items():
        for part in ("sole", "strap"):
            jobs.append((stl / f"{part}_{name}.stl", dict(part=part, foot=f)))
            jobs.append((tmp / f"{part}_{name}.stl", dict(part=part, foot=f, in_place=True)))
    for plate in ("plate_soles", "plate_straps", "fit_test"):
        jobs.append((stl / f"{plate}.stl", dict(part=plate)))
    with ThreadPoolExecutor(8) as pool:
        list(pool.map(lambda j: openscad(j[0], **j[1]), jobs))
    for path in stl.glob("*.stl"):  # OpenSCAD writes text STL; binary is much smaller
        trimesh.load_mesh(path, process=False).export(path)

    print("\nParts")
    for name in sides:
        for part in ("sole", "strap"):
            m = trimesh.load_mesh(stl / f"{part}_{name}.stl", process=True)
            e = m.extents
            ok = m.is_watertight and m.body_count == 1 and abs(m.bounds[0][2]) < 0.02
            report(ok, f"{part}_{name:5s} closed solid, flat on the bed, "
                       f"{e[0]:.0f} x {e[1]:.0f} x {e[2]:.0f} mm, {m.volume/1000*1.24:.1f} g solid")
            span = widest_overhang(m)
            report(span <= MAX_BRIDGE, f"{part}_{name:5s} needs no supports (widest bridge {span:.1f} mm)")

    print("\nClicked together")
    for name in sides:
        sole = solid(trimesh.load_mesh(tmp / f"sole_{name}.stl", process=True))
        strap = solid(trimesh.load_mesh(tmp / f"strap_{name}.stl", process=True))
        v = (sole ^ strap).volume()
        report(v < 0.05, f"{name}: strap and sole do not overlap ({v:.3f} mm3)")
        lifted = [(sole ^ strap.translate((0, 0, dz))).volume() for dz in (0.1, 0.8, 1.5)]
        report(lifted[0] < 0.05 and min(lifted[1:]) > 0.2,
               f"{name}: hooks catch under the sole when the strap is pulled up")

    print("\nPlates")
    for plate in ("plate_soles", "plate_straps"):
        e = trimesh.load_mesh(stl / f"{plate}.stl").extents
        report(e[0] < BED - 10 and e[1] < BED - 10, f"{plate} fits the bed: {e[0]:.0f} x {e[1]:.0f} mm")

    print()
    if failures:
        print(f"{len(failures)} check(s) FAILED")
        sys.exit(1)
    print("All checks passed.")


if __name__ == "__main__":
    main()
