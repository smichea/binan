# Secret-drawer treasure chest ("baul")

A palm-sized chest for the night-market stand. The gold base looks like a plain plinth, but it is a hidden drawer held shut by a loose ball. Tilt the chest to its left and the base slides out of the front.

![Closed chest](img/closed.png)
![Lid open and secret drawer out](img/open.png)

| | |
|---|---|
| Outside | 75 × 53 × 54 mm |
| Main compartment | 68 × 46 × 22 mm, plus the space in the lid |
| Secret drawer | 65 × 33 × 9 mm (coins, rings, SIM and SD cards, a folded note) |

**Status: designed and checked by simulation, not yet printed.** Print the plinth and drawer first and try the lock before printing the rest.

## Opening the secret drawer

1. Hold the chest facing you and tilt it to **its left** (left side down, about 45°, or lay it on its left side).
2. Keep it tilted and slide the base forward. There is a ribbed finger grip under the front, in the notch of the base.
3. To close, set the chest upright and push the drawer in until it stops. It locks by itself.

Tilt first, then pull: if you pull before tilting, the hook pinches the ball and it cannot roll.

![How the ball lock works](img/lock.png)

## Parts

Everything prints in one colour per part, flat on the bed, without supports. Print `stl/plate_a.stl` in the chest colour and `stl/plate_b.stl` in the trim colour, or use the single-part files to fill a bed with several sets.

| Part | Colour | Qty | Print time | Filament |
|---|---|---|---|---|
| `body` | A (chest) | 1 | 2 h 35 | 23 g |
| `lid` | A (chest) | 1 | 1 h 35 | 15 g |
| `plinth` | B (trim) | 1 | 1 h 30 | 15 g |
| `drawer` | B (trim) | 1 | 50 min | 9 g |
| `lid_strap` | B (trim) | 2 | 5 min each | 1 g each |
| `body_strap` | B (trim) | 2 | 2 min each | under 1 g |
| `lock_plate` | B (trim) | 1 | 3 min | under 1 g |
| **One chest** | | | **about 6 h 45** | **about 64 g** |

Times and weights come from slicing with the settings below at Ender-style speeds (45 to 60 mm/s, 500 mm/s² acceleration). A draft profile (0.28 mm layers, 2 walls, 10 % infill) brings one chest down to about 4 h 25 and 59 g.

Hardware per chest:

- 1 ball, 6 mm: steel ball, or an airsoft BB. For a 1/4" (6.35 mm) bicycle bearing ball set `ball_d = 6.35` and re-export.
- 2 hinge pins: 18 mm lengths of 1.75 mm filament.
- 2 screws M3 × 10, pan head. Or superglue instead.
- Superglue for the straps and the lock plate.

## Print settings

- PLA, 0.4 mm nozzle, 0.2 mm layers, 3 walls, 15 % infill.
- No supports and no brim. Keep the orientation of the STL files.
- The lid has a 24 mm bridge across its inside ceiling. The slicer may warn about a long bridge; leave supports off and run the part fan at 100 %.
- A squashed first layer makes the drawer tight. The parts have a small chamfer at the bottom for this, but if the drawer still drags, raise the nozzle slightly or turn on elephant-foot compensation.

## Assembly

1. Slide the drawer into the plinth. It should move freely and click lightly near the end. Take it out again.
2. Drop the ball into the slot at the back of the plinth.
3. Try the lock now, before closing it up: push the drawer in, check it will not pull out, then tilt left and slide it out.
4. Put the body on the two pegs. Fix it with the two M3 screws from underneath. If gluing instead, put small dots on the pegs and the side walls only, and keep glue away from the ball slot.
5. Hold the lid in place and push a filament pin through each hinge from the outside. The pin grips in the body knuckles and turns freely in the lid. Trim it flush.
6. Glue the two arched straps into the lid grooves, the two flat straps into the body grooves, and the lock plate into its recess.

The lid opens to about 108° and stays open. The keyhole plate is decoration only, which also makes it a decoy.

## Adjusting the design

All dimensions are parameters at the top of `chest.scad` ([OpenSCAD](https://openscad.org)).

| If | Change |
|---|---|
| Drawer too tight or too loose | `cl` (0.3 mm per side) |
| Click too strong, or not wanted | `detent` (0.5; 0 removes the bumps) |
| Different ball | `ball_d` |
| Hinge pin too tight or too loose | `pin_tight`, `pin_loose` |
| Lid should stop earlier or later | `heel` |
| Different size | `W`, `D`, `Hb`; keep `D` at 44 or more, the lock needs the depth |

After any change run `python3 check.py`. It re-exports every STL and re-runs the checks (needs `openscad`, and `pip install trimesh manifold3d numpy`).

## What the checks cover

`check.py` moves the real meshes through their motions and looks for collisions. The current design passes all of these:

- every part is a closed solid, sits flat on the bed, and both plates fit a 220 × 220 mm bed
- with the chest closed, no two parts overlap
- the lid swings freely, stops against the body at about 108°, stays open, and the chest does not tip backwards
- the drawer slides out 45 mm with nothing in the way
- chest upright: the ball holds the drawer after 1.4 mm of play
- chest tilted left: the ball rolls clear and the drawer comes out
- closing: the hook pushes the ball aside without jamming and the drawer is locked again

Not covered: friction, real print tolerances, and how reliably the ball rolls on a printed surface. Only a test print answers those.

## Safety

The ball is enclosed once the body is screwed on, so there are no loose small parts in normal use. Still, sell it as not suitable for children under 3.
