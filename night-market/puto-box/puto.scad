// Puto Biñan trinket box for the night-market stand.
//
// A small round box that looks like a puto on a banana leaf.  The cake is the
// box and the top lifts off by its butter pat.  The secret: the cake unscrews
// from the leaf with a third of a turn, and the leaf cup hides a second
// compartment underneath.
//
// Four parts, each printed in one colour without supports:
//   green   base      banana leaf with a threaded cup
//   white   body      the cake, with a threaded foot
//   white   lid       top of the cake (prints upside down)
//   yellow  toppings  butter pat, cheese strips and egg yolks on one plate
//                     that presses into the lid from below
//
// Z = up, 0 = table.  Export one part:  openscad -D 'part="body"' -o body.stl puto.scad

/* [What to show] */
part = "assembly";  // assembly | open | exploded | section | base | body | lid | toppings | plate_white | plate_green | plate_yellow
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
label = "PUTO BIÑAN";
label_size = 4;

/* [Lid] */
lid_over = 0.8;     // how far the lid overhangs the cake
lid_t = 2.4;
bevel = 1.2;
skin = 1.2;         // thickness of lid above the toppings plate
lip_h = 3;
lip_t = 1.6;
lip_clear = 0.25;

/* [Toppings] */
eggs = true;        // salted-egg slices between the cheese strips
plate_r = 18.8;
plate_t = 1.0;
fit = 0.15;         // clearance around each topping where it passes through the lid
top_r = 13;         // distance of the toppings from the centre
cheese = [3.2, 11];
cheese_up = 1.0;
yolk_d = 5.5;
yolk_up = 0.6;
knob = 10;
knob_up = 5;

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
lid_z = body_z + body_total;                 // rim of the cake
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
    // pleats on the outside of the cup, like a folded leaf (and a grip for twisting)
    for (i = [0:17]) rotate([0, 0, i*20])
      translate([rc, 0, leaf_t+0.6]) linear_extrude(base_h) polygon([[eps, -0.7], [eps, 0.7], [-0.5, 0]]);
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

// ---------------------------------------------------------------- toppings
// Footprints of everything that pokes through the lid.
module cheese2d() {
  for (a = [60, 180, 300]) rotate(a) hull() {
    translate([top_r-(cheese[1]-cheese[0])/2, 0]) circle(d = cheese[0], $fn = 24);
    translate([top_r+(cheese[1]-cheese[0])/2, 0]) circle(d = cheese[0], $fn = 24);
  }
}
module yolks2d() { if (eggs) for (a = [0, 120, 240]) rotate(a) translate([top_r, 0]) circle(d = yolk_d, $fn = 32); }
module knob2d() { offset(r = 2, $fn = 24) square(knob-4, center = true); }

// drawn with z = 0 at the rim of the cake
module toppings() {
  z0 = lid_t - skin;
  translate([0, 0, z0-plate_t]) cylinder(r = plate_r, h = plate_t);
  translate([0, 0, z0-eps]) {
    linear_extrude(skin+cheese_up) cheese2d();
    linear_extrude(skin+yolk_up) yolks2d();
    hull() {
      linear_extrude(skin+knob_up-1) knob2d();
      linear_extrude(skin+knob_up) offset(delta = -1) knob2d();
    }
  }
}

// ---------------------------------------------------------------- lid
module lid() {
  rl = Rb - wall - lip_clear;
  difference() {
    union() {
      hull() {
        cylinder(r = Rl, h = lid_t-bevel);
        cylinder(r = Rl-bevel, h = lid_t);
      }
      translate([0, 0, -lip_h]) difference() {
        cylinder(r = rl, h = lip_h+eps);
        translate([0, 0, -1]) cylinder(r = rl-lip_t, h = lip_h+2);
      }
    }
    // pocket for the toppings plate, open from below
    translate([0, 0, -lip_h-1]) cylinder(r = plate_r+0.2, h = lip_h+1+lid_t-skin);
    // windows for the toppings
    translate([0, 0, -1]) linear_extrude(lid_t+2) offset(delta = fit) {
      cheese2d(); yolks2d(); knob2d();
    }
    // the cake is scored into six slices
    for (a = [30 : 60 : 330]) rotate(a) translate([knob/2+2.5, 0, lid_t]) groove(Rl);
    // outline of each egg slice around its yolk
    if (eggs) for (a = [0, 120, 240]) rotate(a) translate([top_r, 0, lid_t])
      rotate_extrude($fn = 48) translate([yolk_d/2+2.2, 0]) polygon([[-0.4, eps], [0.4, eps], [0, -0.4]]);
  }
}

// ---------------------------------------------------------------- views
// turn = degrees the cake has been unscrewed; lift = how far the lid is raised
module assembly(turn = 0, lift = 0) {
  color(green) base();
  translate([0, 0, body_z + turn/360*lead]) rotate([0, 0, body_turn + turn]) {
    color(white) body();
    translate([0, 0, body_total + lift]) {
      color(white) lid();
      color(yellow) toppings();
    }
  }
}

module exploded() {
  color(green) base();
  translate([0, 0, body_z+16]) rotate([0, 0, body_turn]) color(white) body();
  translate([0, 0, lid_z+28]) color(white) lid();
  translate([0, 0, lid_z+20]) color(yellow) toppings();
}

module section() {
  difference() {
    assembly();
    translate([-100, -200, -1]) cube([200, 200, 100]);
  }
}

// ---------------------------------------------------------------- print layouts
module print_base() { base(); }
module print_body() { body(); }
module print_lid() { translate([0, 0, lid_t]) rotate([180, 0, 0]) lid(); }
module print_toppings() { translate([0, 0, plate_t-(lid_t-skin)]) toppings(); }

module grid(n, dx, dy) {
  cols = ceil(n/2);
  for (i = [0 : n-1]) translate([(i % cols)*dx, floor(i/cols)*dy, 0]) children();
}

module plate_white() {
  grid(sets, 2*Rl+4, 2*Rl+4) print_body();
  translate([0, 2*(2*Rl+4), 0]) grid(sets, 2*Rl+4, 2*Rl+4) print_lid();
}
module plate_green() { grid(sets, leaf+4, leaf+4) print_base(); }
module plate_yellow() { grid(sets, 2*plate_r+4, 2*plate_r+4) print_toppings(); }

if (part == "assembly") assembly();
else if (part == "open") assembly(turn = 150, lift = 14);
else if (part == "exploded") exploded();
else if (part == "section") section();
else if (part == "plate_white") plate_white();
else if (part == "plate_green") plate_green();
else if (part == "plate_yellow") plate_yellow();
else if (in_place) {
  if (part == "base") base();
  else if (part == "body") translate([0, 0, body_z]) rotate([0, 0, body_turn]) body();
  else if (part == "lid") translate([0, 0, lid_z]) rotate([0, 0, body_turn]) lid();
  else if (part == "toppings") translate([0, 0, lid_z]) rotate([0, 0, body_turn]) toppings();
}
else if (part == "base") print_base();
else if (part == "body") print_body();
else if (part == "lid") print_lid();
else if (part == "toppings") print_toppings();

// numbers that check.py needs
if (part == "info") {
  echo(str("INFO lead=", lead)); echo(str("INFO thread_h=", thread_h));
  echo(str("INFO body_z=", body_z)); echo(str("INFO lid_z=", lid_z));
  echo(str("INFO outer=", leaf, " x ", leaf, " x ", lid_z+lid_t+knob_up));
  echo(str("INFO main_inner=", 2*(Rb-wall), " dia x ", body_total-floor_t-lip_h+lip_h));
  echo(str("INFO hidden_inner=", 2*(rm+tclear-ta), " dia x ", hidden_h));
}
