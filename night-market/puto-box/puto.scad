// Puto Biñan trinket box for the night-market stand.
//
// A small round box that looks like a puto: a white domed cake with one
// triangular slice of cheese on top, sitting in a pleated cup on a banana leaf.
// The dome is the lid.  The secret: the cake unscrews from the leaf with a
// third of a turn, and the leaf's cup hides a second compartment underneath.
//
// Four parts, each printed in one colour without supports:
//   green   base    banana leaf with a pleated, threaded cup
//   white   body    the cake, with a threaded foot
//   white   lid     the domed top of the cake
//   yellow  cheese  the triangle; it pushes into the lid from below
//
// Z = up, 0 = table.  Export one part:  openscad -D 'part="body"' -o body.stl puto.scad

/* [What to show] */
part = "assembly";  // assembly | open | exploded | base | body | lid | cheese | plate_white | plate_green | plate_yellow
in_place = false;   // true: export a part where it sits in the assembly (used by check.py)
sets = 4;           // boxes per plate

/* [Cake] */
Rb = 23.2;          // radius of the cake
wall = 2.0;
floor_t = 1.2;
body_h = 13;        // height of cake showing between leaf cup and lid

/* [Screw joint between cake and leaf] */
rm = 21.6;          // mean radius of the thread on the cake's foot
ta = 0.6;           // half the thread depth
starts = 3;
pitch = 4;
thread_h = 4;       // a third of a turn at these numbers
tclear = 0.4;       // radial clearance between the two threads
seat = 0.8;         // cone the cake sits on
hidden_h = 4;       // depth of the hidden compartment

/* [Leaf] */
leaf = 68;          // size of the square of banana leaf
leaf_t = 1.2;
collar_t = 2.0;
flutes = 24;        // pleats around the cup
label = "PUTO BIÑAN";
label_size = 4;

/* [Lid] */
lid_over = 1.2;     // how far the lid overhangs the cake (this is what you grip)
dome_h = 9;         // height of the dome
band = 0.8;         // straight edge below the dome
skirt = 1.6;        // how far the lid reaches down over the rim of the cake
rim_clear = 0.25;   // clearance each side of the rim

/* [Cheese] */
cheese_pts = [[-3, 0], [15, 8.5], [15, -8.5]];  // corners of the triangle; [0, 0] is the middle of the lid
cheese_r = 1.2;     // corner rounding
cheese_up = 0.8;    // how far the cheese stands above the cake
flare = 1.2;        // wider foot that stops the cheese pulling out upwards
fit = 0.15;         // clearance between cheese and lid

/* [Hidden] */
eps = 0.01;
ef = 0.3;
$fn = 96;

// ---------------------------------------------------------------- derived
lead = starts*pitch;
base_h = floor_t + hidden_h + thread_h + seat;
body_z = floor_t + hidden_h;                 // where the cake's foot ends up
shoulder = Rb - (rm + ta);                   // height of the 45 degree shoulder
body_total = thread_h + shoulder + body_h;
rim_z = body_z + body_total;                 // rim of the cake
lid_z = rim_z - rim_clear;                   // the rim sits a little way up the lid's groove
Rl = Rb + lid_over;
body_turn = 360*hidden_h/lead;               // lines the two threads up

green = "ForestGreen";
white = "Ivory";
yellow = "Gold";

// ---------------------------------------------------------------- helpers
// Round cross-section with one bump per thread start; twisting it makes a
// thread whose flanks are gentle enough to print without supports.
module lobes2d(r) {
  polygon([for (i = [0:179]) let (t = i*2) (r + ta*cos(starts*t))*[cos(t), sin(t)]]);
}

module thread(r, h) {
  linear_extrude(h, twist = -360*h/lead, slices = ceil(h/0.2), convexity = 4) lobes2d(r);
}

module groove(len, w = 0.8, d = 0.4) {  // V groove along X, cut downwards from z = 0
  rotate([90, 0, 90]) linear_extrude(len) polygon([[-w/2, eps], [w/2, eps], [0, -d]]);
}

// ---------------------------------------------------------------- base (leaf)
module leaf2d() { offset(r = 9) square(leaf-18, center = true); }

module veins() {
  // midrib along X, side veins fanning away from it
  intersection() {
    translate([0, 0, -1]) linear_extrude(2) difference() {
      offset(delta = -1.5) leaf2d();
      translate([-leaf/2, -leaf/2]) square([leaf, 10.5]);   // keep the label clear
      difference() {                                        // and stay out of the cup's wall
        circle(r = rm+tclear+ta+collar_t+0.3);
        circle(r = rm+tclear-ta-0.3);
      }
    }
    union() {
      translate([-leaf, 0, 0]) groove(2*leaf, 1.2, 0.4);
      for (s = [-1, 1]) for (i = [-12:12])
        translate([i*4.2, 0, 0]) rotate([0, 0, s*72]) groove(leaf, 0.7, 0.3);
    }
  }
}

module base() {
  rc = rm + tclear + ta + collar_t;
  difference() {
    union() {
      hull() {
        linear_extrude(eps) offset(delta = -ef) leaf2d();
        translate([0, 0, ef]) linear_extrude(leaf_t-ef) leaf2d();
      }
      cylinder(r = rc, h = base_h);
    }
    // threaded cup; the part below the cake's foot is the hidden compartment
    translate([0, 0, floor_t]) thread(rm + tclear, base_h);
    // cone the cake's shoulder sits in
    translate([0, 0, base_h-seat-2*ta]) cylinder(r1 = rm-ta, r2 = rm+ta+seat+eps, h = seat+2*ta+eps);
    // pleats round the cup, like a paper case (and a grip for twisting)
    for (i = [0 : flutes-1]) rotate([0, 0, i*360/flutes])
      translate([rc, 0, leaf_t+0.6]) linear_extrude(base_h) polygon([[eps, -1.1], [eps, 1.1], [-0.7, 0]]);
    // leaf veins and name
    translate([0, 0, leaf_t]) veins();
    if (label != "")
      translate([0, -leaf/2+5.6, leaf_t-0.4]) linear_extrude(1)
        text(label, size = label_size, font = "Liberation Sans:style=Bold",
             halign = "center", valign = "center");
  }
}

// ---------------------------------------------------------------- body (cake)
module body() {
  difference() {
    union() {
      // threaded foot, tapered at the bottom so it starts easily
      intersection() {
        thread(rm, thread_h+eps);
        union() {
          cylinder(r1 = rm-ta-0.2, r2 = rm+ta+0.2, h = 2*ta+0.4);
          translate([0, 0, 2*ta+0.4-eps]) cylinder(r = rm+ta+1, h = thread_h);
        }
      }
      // 45 degree shoulder; it starts at the bottom of the thread grooves so no ledge hangs over them
      translate([0, 0, thread_h-2*ta]) cylinder(r1 = rm-ta, r2 = Rb+eps, h = shoulder+2*ta+eps);
      translate([0, 0, thread_h+shoulder-eps]) cylinder(r = Rb, h = body_h+eps);
    }
    translate([0, 0, floor_t]) cylinder(r = rm-ta-1.6, h = body_total);
    translate([0, 0, thread_h+shoulder]) cylinder(r = Rb-wall, h = body_total);
  }
}

// ---------------------------------------------------------------- lid and cheese
// Both are drawn with z = 0 at the top of the straight part of the lid's groove.

// Half outline of the lid, to be spun round: a straight edge, then a dome.
module dome2d(lift = 0) {
  polygon(concat(
    [[0, -skirt], [Rl-ef, -skirt], [Rl, -skirt+ef]],
    [for (t = [0 : 3 : 90]) [Rl*cos(t), band+lift+dome_h*sin(t)]]));
}

module cheese2d() {
  hull() for (p = cheese_pts) translate(p) circle(r = cheese_r, $fn = 24);
}

// The cheese as a column through the lid, with a flared foot at the bottom.
module cheese_column(grow = 0) {
  hull() {
    translate([0, 0, -skirt]) linear_extrude(eps) offset(r = flare+grow) cheese2d();
    translate([0, 0, -skirt+flare]) linear_extrude(eps) offset(r = grow+eps) cheese2d();
  }
  translate([0, 0, -skirt+flare-eps]) linear_extrude(skirt+band+dome_h+5) offset(r = grow+eps) cheese2d();
}

module lid() {
  difference() {
    rotate_extrude() dome2d();
    // groove that drops over the rim of the cake; its pointed roof prints without bridging
    rotate_extrude() polygon([[Rb-wall-rim_clear, -skirt-1], [Rb+rim_clear, -skirt-1], [Rb+rim_clear, 0],
                              [Rb-wall/2, wall/2+rim_clear], [Rb-wall-rim_clear, 0]]);
    // window for the cheese
    cheese_column(fit);
    translate([0, 0, -skirt-1]) linear_extrude(1+eps) offset(r = flare+fit) cheese2d();
  }
}

module cheese() {
  intersection() {
    cheese_column();
    rotate_extrude() dome2d(cheese_up);   // top follows the dome, a little above it
    translate([0, 0, -skirt-1]) cylinder(r = Rb-wall-rim_clear-0.5, h = skirt+band+dome_h+5);
  }
}

// ---------------------------------------------------------------- views
// turn = degrees the cake has been unscrewed; lift = how far the lid is raised
module assembly(turn = 0, lift = 0) {
  color(green) base();
  translate([0, 0, turn/360*lead]) rotate([0, 0, body_turn + turn]) {
    translate([0, 0, body_z]) color(white) body();
    translate([0, 0, lid_z + lift]) {
      color(white) lid();
      color(yellow) cheese();
    }
  }
}

module exploded() {
  color(green) base();
  translate([0, 0, body_z+16]) rotate([0, 0, body_turn]) color(white) body();
  translate([0, 0, lid_z+34]) color(white) lid();
  translate([0, 0, lid_z+52]) color(yellow) cheese();
}

// ---------------------------------------------------------------- print layouts
module print_base() { base(); }
module print_body() { body(); }
module print_lid() { translate([0, 0, skirt]) lid(); }
module print_cheese() { translate([-6, 0, skirt]) cheese(); }

module plate_white() {
  for (k = [0 : 2*sets-1])
    translate([(k % 3)*(2*Rl+4), floor(k/3)*(2*Rl+4), 0])
      if (k < sets) print_body(); else print_lid();
}
module plate_green() {
  for (k = [0 : sets-1]) translate([(k % 2)*(leaf+4), floor(k/2)*(leaf+4), 0]) print_base();
}
module plate_yellow() {
  for (k = [0 : sets-1]) translate([(k % 2)*28, floor(k/2)*28, 0]) print_cheese();
}

if (part == "assembly") assembly();
else if (part == "open") assembly(turn = 150, lift = 16);
else if (part == "exploded") exploded();
else if (part == "plate_white") plate_white();
else if (part == "plate_green") plate_green();
else if (part == "plate_yellow") plate_yellow();
else if (in_place) {
  if (part == "base") base();
  else if (part == "body") translate([0, 0, body_z]) rotate([0, 0, body_turn]) body();
  else if (part == "lid") translate([0, 0, lid_z+0.03]) rotate([0, 0, body_turn]) lid();
  else if (part == "cheese") translate([0, 0, lid_z+0.03]) rotate([0, 0, body_turn]) cheese();
}
else if (part == "base") print_base();
else if (part == "body") print_body();
else if (part == "lid") print_lid();
else if (part == "cheese") print_cheese();

// numbers that check.py needs
if (part == "info") {
  echo(str("INFO lead=", lead)); echo(str("INFO thread_h=", thread_h));
  echo(str("INFO outer=", leaf, " x ", leaf, " x ", lid_z+band+dome_h+cheese_up));
  echo(str("INFO main_inner=", 2*(Rb-wall), " dia x ", body_total-floor_t-skirt+rim_clear));
  echo(str("INFO hidden_inner=", 2*(rm+tclear-ta), " dia x ", hidden_h));
}
