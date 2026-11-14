// Tsinelas (flip-flop) keychain for the Biñan night-market stand.
//
// Two parts, each printed in one colour without supports:
//   sole   - prints flat, name of the town pressed into the top
//   strap  - prints upside down, snaps into the sole with two hooks
// Customers pick a sole colour and a strap colour and you click them together.
//
// Axes: X = across the foot, Y = heel (0) to toes, Z = up (0 = underside).
// The model is a right foot; foot = -1 mirrors it into a left foot.
//
// Export one part:   openscad -D 'part="sole"' -D foot=-1 -o sole_left.stl tsinelas.scad

/* [What to show] */
part = "pair";     // pair | colours | assembly | exploded | sole | strap | plate_soles | plate_straps | fit_test
foot = 1;          // 1 = right foot, -1 = left foot
in_place = false;  // true: export a part where it sits in the assembly (used by check.py)
pairs = 6;         // pairs per plate

/* [Sole] */
label = "BIÑAN";   // text pressed into the sole; "" for none
label_size = 4.2;
label_pos = [0.45, 18];  // centre of the text; it has to sit between the two strap slots
ts = 5;            // thickness
heel_r = 8.5;
ball_r = 11.5;
toe_r = 9.5;
ball = [-0.5, 34]; // centre of the widest part
toe = [-1.8, 42.5];
ring_d = 4.5;      // hole for the key ring
ring_y = 5;
tread = true;      // grooves underneath

/* [Strap] */
sw = 3.6;            // strap width
st = 2.2;            // strap thickness
post = [-1.5, 41];   // toe post
anchor = [5.6, 22];  // where the strap meets the sole, each side
crown = 5.5;         // height of the strap above the sole
post_h = 2.0;
leg_run = 2.4;       // how far the strap leans in on its way down
pin_d = 3.0;         // toe pin
pin_fit = 0.2;       // toe hole = pin_d + pin_fit

/* [Snap hooks] */
hook = 0.7;          // how far each hook sticks out
ledge = 2.4;         // depth of the pocket underneath that the hooks catch in
slot_gap = 0.1;      // gap between strap leg and the outer wall of its slot

/* [Hidden] */
eps = 0.01;
ef = 0.3;            // bottom chamfer against "elephant foot"
ch = 0.6;            // top edge chamfer
$fn = 64;

// ---------------------------------------------------------------- derived
slot_len = st + hook + 0.2;  // room for the leg to bend in while the hook passes
slot_w = sw + 0.4;
anchors = [[anchor[0], anchor[1]], [-anchor[0], anchor[1]]];
function heading(a) = atan2(a[1] - post[1], a[0] - post[0]);
function reach(a) = norm([a[0] - post[0], a[1] - post[1]]);
strap_top = ts + crown + st/2;

sole_colour = "Gold";
strap_colour = "RoyalBlue";

// ---------------------------------------------------------------- helpers
// Work in the plane of one strap arm: X runs from the toe post to the anchor.
module arm_frame(a) {
  translate([post[0], post[1], 0]) rotate([0, 0, heading(a)]) children();
}

// 2D shape drawn in (distance along the arm, height), given the strap's width
module arm_extrude(w) {
  rotate([90, 0, 0]) linear_extrude(w, center = true) children();
}

// ---------------------------------------------------------------- sole
module outline2d() {
  hull() {
    translate([0, heel_r]) circle(heel_r);
    translate(ball) circle(ball_r);
    translate(toe) circle(toe_r);
  }
}

module sole_right() {
  difference() {
    hull() {
      linear_extrude(eps) offset(delta = -ef) outline2d();
      translate([0, 0, ef]) linear_extrude(ts-ef-ch) outline2d();
      translate([0, 0, ts-eps]) linear_extrude(eps) offset(delta = -ch) outline2d();
    }
    // key ring
    translate([0, ring_y, -1]) cylinder(d = ring_d, h = ts+2);
    // toe pin
    translate([post[0], post[1], -1]) cylinder(d = pin_d+pin_fit, h = ts+2);
    // slots for the two strap legs, with a pocket underneath for the hooks
    for (a = anchors) arm_frame(a) {
      d1 = reach(a) + st/2 + slot_gap;   // outer wall of the slot
      translate([d1-slot_len, -slot_w/2, -1]) cube([slot_len, slot_w, ts+2]);
      translate([d1-slot_len, -slot_w/2, -1]) cube([slot_len+hook+0.3, slot_w, ledge+1]);
    }
    // tread: V-shaped grooves, so their 45 degree walls print without bridging
    // They stop short of the edge, which keeps the first layer in one piece.
    if (tread) intersection() {
      translate([0, 0, -1]) linear_extrude(3) offset(delta = -1.8) outline2d();
      for (y = concat([for (y = [9 : 3 : 15]) y], [for (y = [28 : 3 : 49]) y]))
        translate([-20, y, 0]) rotate([90, 0, 90]) linear_extrude(40)
          polygon([[-0.55, -eps], [0.55, -eps], [0, 0.55]]);
    }
  }
}

module sole(f = foot) {
  difference() {
    mirror([f < 0 ? 1 : 0, 0, 0]) sole_right();
    if (label != "")
      translate([label_pos[0], label_pos[1], ts-0.6]) rotate([0, 0, 90]) linear_extrude(1)
        text(label, size = label_size, font = "Liberation Sans:style=Bold",
             halign = "center", valign = "center");
  }
}

// ---------------------------------------------------------------- strap
// One arm of the strap as a flat profile: up from the toe post at 45 degrees,
// level across the foot, down to the sole, then a leg with a hook that goes
// through the slot and catches under the ledge.
module arm2d(D) {
  // starts a little in front of the post so the arm fully covers the post's top
  pts = [[-0.7, ts+post_h-0.7], [crown-post_h, ts+crown], [D-leg_run, ts+crown], [D, ts+st/2+0.2]];
  for (i = [0:2]) hull() {
    translate(pts[i]) circle(d = st, $fn = 32);
    translate(pts[i+1]) circle(d = st, $fn = 32);
  }
  xo = D + st/2;
  polygon([[D-st/2, ts+st/2+0.2], [xo, ts+st/2+0.2],
           [xo, ledge-0.15], [xo+hook, ledge-0.15-hook],   // 45 degree catch
           [xo-0.2, 0.25], [D-st/2+0.3, 0.25], [D-st/2, 0.55]]);
}

module strap_right() {
  for (a = anchors) arm_frame(a) arm_extrude(sw) arm2d(reach(a));
  // toe post and its pin
  translate([post[0], post[1], 0]) {
    translate([0, 0, ts+0.05]) cylinder(d = sw, h = post_h);
    translate([0, 0, 1.6]) cylinder(d = pin_d, h = ts-1.5);
    translate([0, 0, 1.0]) cylinder(d1 = pin_d-1.0, d2 = pin_d, h = 0.6+eps);
  }
}

module strap(f = foot) {
  mirror([f < 0 ? 1 : 0, 0, 0]) strap_right();
}

// ---------------------------------------------------------------- views
module assembly(f = foot, c1 = sole_colour, c2 = strap_colour) {
  color(c1) sole(f);
  color(c2) strap(f);
}

module pair() {
  translate([-14, 0, 0]) assembly(-1);
  translate([14, 0, 0]) assembly(1);
}

// a few colour combinations side by side
module colours() {
  combos = [["WhiteSmoke", "Crimson"], ["Gold", "RoyalBlue"], ["DimGray", "Gold"],
            ["DeepSkyBlue", "WhiteSmoke"], ["LimeGreen", "DimGray"]];
  for (i = [0 : len(combos)-1])
    translate([(i-2)*27, 0, 0]) assembly(i % 2 == 0 ? 1 : -1, combos[i][0], combos[i][1]);
}

module exploded() {
  color(sole_colour) sole(foot);
  color(strap_colour) translate([0, 0, 14]) strap(foot);
}

// ---------------------------------------------------------------- print layouts
module print_sole(f = foot) { sole(f); }
module print_strap(f = foot) { translate([0, 0, strap_top]) rotate([0, 180, 0]) strap(-f); }

module plate_soles() {
  for (i = [0 : pairs-1]) {
    translate([i*(2*ball_r+4), 0, 0]) print_sole(1);
    translate([i*(2*ball_r+4), 58, 0]) print_sole(-1);
  }
}

module plate_straps() {
  for (i = [0 : pairs-1]) {
    translate([i*20, 0, 0]) print_strap(1);
    translate([i*20, 32, 0]) print_strap(-1);
  }
}

// A small coupon to tune the snap on your printer before a big batch:
// the middle of a sole with its three holes.  Print it with one strap.
module fit_test() {
  intersection() {
    sole(1);
    translate([-9.5, 17, -1]) cube([19, 28.5, ts+2]);
  }
}

if (part == "pair") pair();
else if (part == "colours") colours();
else if (part == "assembly") assembly(foot);
else if (part == "exploded") exploded();
else if (part == "plate_soles") plate_soles();
else if (part == "plate_straps") plate_straps();
else if (part == "fit_test") fit_test();
else if (part == "sole") { if (in_place) sole(foot); else print_sole(foot); }
else if (part == "strap") { if (in_place) strap(foot); else print_strap(foot); }
