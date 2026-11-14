// Secret-drawer treasure chest ("baul") for the Biñan night-market stand.
//
// A small chest whose base is really a hidden drawer.  The drawer is held by a
// loose steel ball: tilt the chest to its left and the ball rolls out of the
// way, so the base slides out of the front.  Push it back in and it re-locks.
//
// Every part prints in one colour, without supports, on a 220 x 220 bed.
//   colour A: body, lid
//   colour B: plinth, drawer, lid straps, body straps, lock plate
//   hardware: 1 ball (6 mm), 2 hinge pins (1.75 mm filament offcuts, 18 mm),
//             2 screws M3 x 10 (or glue)
//
// Axes: X = width (negative = chest's left, seen from the front),
//       Y = depth (0 = front face of the body), Z = up (0 = table).
//
// Export one part:   openscad -D 'part="body"' -o body.stl chest.scad

/* [What to show] */
part = "assembly";   // assembly | exploded | lock_view | body | lid | plinth | drawer | lid_strap | body_strap | lock_plate | plate_a | plate_b
in_place = false;    // true: export a part where it sits in the assembly (used by check.py)
open_angle = 0;      // lid angle in the assembly views
drawer_out = 0;      // drawer travel in the assembly views
ball_clear = false;  // true: draw the ball at the unlocked end of its track

/* [Chest size] */
W = 72;         // body width
D = 50;         // body depth
Hb = 24;        // body height
wall = 2.0;
floor_t = 2.0;  // body floor
lid_v = 6;      // vertical part of the lid
lid_s = 12;     // run and rise of the 45 degree shoulders

/* [Plinth and drawer] */
po = 1.6;        // how far the plinth sticks out around the body
Hp = 12;         // plinth height
tf = 1.2;        // plinth floor
wp = 3.2;        // plinth side walls
tp = 2.4;        // drawer front panel
cl = 0.3;        // sliding clearance each side of the drawer
dwall = 1.6;     // drawer walls
dfloor = 1.2;    // drawer floor
top_gap = 0.4;   // gap between drawer and body floor
detent = 0.5;    // height of the two click bumps (0 = none)

/* [Ball lock] */
ball_d = 6;       // set to your ball: 6 (steel ball / airsoft BB) or 6.35 (1/4" bearing ball)
ball_gap = 0.35;  // extra track width
th = 3.6;         // thickness of the drawer's hook
stem_w = 5;
head_e = 1.2;     // how far the hook reaches past the ball centre
head_tip = 2.4;
head_off = 0.5;   // gap between ball track and hook when the drawer is home
wedge_deg = 40;   // lead-in that pushes the ball aside when the drawer closes
slope_deg = 15;   // track slope that rolls the ball back to "locked"
seat_d = 2.6;     // dimple the locked ball rests in
div_t = 1.6;
back_skin = 1.2;

/* [Hinge] */
hinge_r = 3;
hinge_up = 2;      // hinge axis above the rim
hinge_gap = 0.4;
pin_d = 1.75;
pin_tight = 0.2;   // body knuckles grip the pin
pin_loose = 0.55;  // lid knuckles turn on it
lidk_w = 7;
bodyk_w = 5;
k_gap = 0.35;
heel = 5;          // lid stop: longer heel = lid stops earlier (5 -> about 107 degrees)

/* [Trim] */
strap_x = 20;     // strap / hinge positions, +- from centre
strap_w = 7;
strap_t = 1.3;
groove = 0.5;
plate_w = 12;
plate_h = 13;
plate_t = 1.5;

/* [Fixing body to plinth] */
screw_d = 3;      // M3 x 10; holes are simply left empty if you glue instead
peg_d = 4;

/* [Hidden] */
eps = 0.01;
ef = 0.3;         // bottom chamfer against "elephant foot"
$fn = 48;

// ---------------------------------------------------------------- derived
PW = W + 2*po;
PD = D + 2*po;
yF = -po;
yB = D + po;
zr = Hp + Hb;                       // rim = parting line between body and lid
r = ball_d/2;

xs0 = -stem_w/2;
xs1 = stem_w/2;
x_seat = xs0 - 0.2 - r;             // ball centre when locked
xh0 = x_seat - head_e;              // tip of the hook
wedge = (xs0 - xh0)*tan(wedge_deg);
pocket = head_off + head_tip + wedge + 0.4;
y_pk1 = yB - back_skin;
y_tr1 = y_pk1 - pocket;             // track rear edge
track_w = ball_d + ball_gap;
y_tr0 = y_tr1 - track_w;            // track front edge
y_trc = (y_tr0 + y_tr1)/2;
y_div0 = y_tr0 - div_t;
y_dback = y_div0 - 0.4;             // drawer rear face
y_dfront = yF + tp;                 // back of the front panel
y_fix0 = y_dfront + 0.2;            // where the fixed plinth walls start
x_open0 = xh0 - 0.5;
x_open1 = xs1 + 0.5;
x_tr0 = xh0 - ball_d - 1.0;         // left end of the track
x_tr1 = xs0;                        // right end: the ball cannot get in front of the stem
x_slope0 = xh0 - 0.1;               // the track floor rises to the left of this
open_h = th + 0.5;
y_hf = y_tr1 + head_off;            // front face of the hook head
xd = PW/2 - wp - cl;                // drawer half width
zt = Hp - top_gap;                  // drawer top

ya = D + hinge_r + hinge_gap;       // hinge axis
za = zr + hinge_up;
kc = sqrt(hinge_r*hinge_r - hinge_up*hinge_up);

x_scr = W/2 - wall - 3.5;
y_scr = D - wall - 3.5;
x_peg = 18;
y_peg = D - wall - 4;
y_det = y_dback - 10;
z_det = Hp - 3.5;

seat_sink = r - sqrt(r*r - seat_d*seat_d/4);
function track_floor(x) = tf + (x < x_slope0 ? (x_slope0 - x)*tan(slope_deg) : 0);
function ball_z(x) = max(tf + r, tf + (x_slope0 - x)*tan(slope_deg) + r/cos(slope_deg));
x_clear = x_tr0 + r + 0.2;
ball_locked = [x_seat, y_trc, tf + r - seat_sink];
ball_free = [x_clear, y_trc, ball_z(x_clear)];

colA = "SaddleBrown";
colB = "Goldenrod";

// ---------------------------------------------------------------- helpers
// box between two corners, bottom edges chamfered by c
module cbox(p0, p1, c = ef) {
  hull() {
    translate([p0[0]+c, p0[1]+c, p0[2]]) cube([p1[0]-p0[0]-2*c, p1[1]-p0[1]-2*c, eps]);
    translate([p0[0], p0[1], p0[2]+c]) cube([p1[0]-p0[0], p1[1]-p0[1], p1[2]-p0[2]-c]);
  }
}

// 2D shape drawn in (y, z), extruded along x from x0 over width w
module yz_extrude(x0, w) {
  translate([x0, 0, 0]) rotate([90, 0, 90]) linear_extrude(w) children();
}

module mirror_x() {
  children();
  mirror([1, 0, 0]) children();
}

// ---------------------------------------------------------------- plinth
// Outer shape of plinth + drawer front together: a block whose top edges are
// chamfered in to the footprint of the body.
module plinth_env(s = 0) {
  hull() {
    translate([-PW/2+ef+s, yF+ef+s, 0]) cube([PW-2*ef-2*s, PD-2*ef-2*s, eps]);
    translate([-PW/2+s, yF+s, ef]) cube([PW-2*s, PD-2*s, Hp-po-ef]);
    translate([-W/2+s, s, Hp-eps]) cube([W-2*s, D-2*s, eps]);
  }
}

module track_cut() {
  translate([0, y_tr1, 0]) rotate([90, 0, 0]) linear_extrude(track_w)
    polygon([[x_tr1, tf], [x_tr1, Hp+1], [x_tr0, Hp+1],
             [x_tr0, track_floor(x_tr0)], [x_slope0, tf]]);
}

module plinth() {
  difference() {
    union() {
      difference() {
        plinth_env();
        // room for the drawer's front panel
        translate([-PW, yF-1, tf]) cube([2*PW, y_fix0-yF+1, Hp]);
        // drawer cavity
        translate([-(PW/2-wp), yF, tf]) cube([PW-2*wp, y_div0-yF, Hp]);
        // ball track, open at the top (the body closes it)
        track_cut();
        // low tunnel for the drawer's hook: through the divider, under the end
        // of the track and into the back wall.  Too low for the ball to follow.
        // (The middle piece is 1 mm taller so its roof bridges between two
        // finished walls instead of hanging from one side.)
        translate([x_open0, y_div0-eps, tf]) cube([x_open1-x_open0, div_t+2*eps, open_h]);
        translate([x_open0, y_tr1-eps, tf]) cube([x_open1-x_open0, y_pk1-y_tr1+eps, open_h]);
        translate([x_tr1-eps, y_tr0-eps, tf]) cube([x_open1-x_tr1+eps, track_w+2*eps, open_h+1]);
        // dimple that centres the locked ball
        translate([x_seat, y_trc, 0.4]) cylinder(d = seat_d, h = tf);
        // finger notch under the drawer front, plus matching dummy notches
        translate([-11, yF-1, -1]) cube([22, 13, tf+1+eps]);
        translate([-11, yB-1.5, -1]) cube([22, 3, tf+1]);
        mirror_x() translate([PW/2-1.5, D/2-9, -1]) cube([3, 18, tf+1]);
        // dummy seams on sides and back, so the real drawer seams look like trim
        difference() {
          union() {
            translate([-PW, y_fix0, tf]) cube([2*PW, PD, 0.25]);
            translate([-PW, D-y_fix0, tf]) cube([2*PW, 0.25, Hp]);
          }
          plinth_env(0.3);
        }
        // screws up into the body
        mirror_x() translate([x_scr, y_scr, 0]) {
          translate([0, 0, -1]) cylinder(d = screw_d+0.4, h = Hp+2);
          translate([0, 0, -1]) cylinder(d = 2*screw_d+0.2, h = Hp-4.4+1);
          translate([0, 0, Hp-4.4-eps]) cylinder(d1 = 2*screw_d+0.2, d2 = screw_d+0.4, h = 1.4);
        }
      }
      // pegs that locate the body
      mirror_x() translate([x_peg, y_peg, Hp-eps]) cylinder(d = peg_d, h = 0.9);
      // click bumps for the drawer
      if (detent > 0)
        mirror_x() translate([PW/2-wp+1.2-detent, y_det, z_det]) sphere(r = 1.2, $fn = 24);
    }
  }
}

// ---------------------------------------------------------------- drawer
module hook2d() {
  polygon([[xs0, y_dback-0.5], [xs1, y_dback-0.5],
           [xs1, y_hf+head_tip+wedge], [xs0, y_hf+head_tip+wedge],
           [xh0, y_hf+head_tip], [xh0, y_hf], [xs0, y_hf]]);
}

module drawer() {
  difference() {
    union() {
      // front panel: looks like the front of the plinth
      intersection() {
        plinth_env();
        cbox([-PW/2, yF, tf], [PW/2, y_dfront, Hp-0.25]);
      }
      // tray
      cbox([-xd, yF+0.5, tf], [xd, y_dback, zt]);
      // hook
      translate([0, 0, tf]) linear_extrude(th) hook2d();
    }
    // inside of the tray
    translate([-xd+dwall, y_dfront, tf+dfloor])
      cube([2*(xd-dwall), y_dback-dwall-y_dfront, Hp]);
    // lead-in on the rear corners
    mirror_x() translate([xd, y_dback, tf-1]) rotate([0, 0, 45]) cube([1.6, 1.6, 2*Hp], center = true);
    // grooves the click bumps drop into
    if (detent > 0)
      mirror_x() translate([xd+1.2-(detent-cl+0.1), y_det, tf+dfloor+1]) cylinder(r = 1.2, h = Hp, $fn = 24);
    // finger grip underneath, reached through the notch in the plinth
    for (i = [0:4]) translate([-9, y_dfront+1.5+i*2.1, tf-eps]) cube([18, 1.0, 0.5]);
  }
}

// ---------------------------------------------------------------- hinge
module body_knuckle2d() {
  union() {
    circle(r = hinge_r);
    polygon([[-(hinge_r+hinge_gap)-0.6, -hinge_up], [kc, -hinge_up],
             [-(hinge_r+hinge_gap)-0.6, -hinge_up-(kc+hinge_r+hinge_gap+0.6)]]);
  }
}

module lid_knuckle2d() {
  hull() {
    intersection() {
      circle(r = hinge_r);
      translate([-hinge_r, -hinge_up]) square([2*hinge_r, hinge_r+hinge_up]);
    }
    translate([-(hinge_r+hinge_gap)-0.6, -hinge_up]) square([0.6+hinge_r+hinge_gap+heel, 1.2]);
    translate([-(hinge_r+hinge_gap)-0.6, -hinge_up]) square([0.6, lid_v]);
  }
}

module body_knuckles() {
  for (sx = [-1, 1]) for (s = [-1, 1])
    translate([0, ya, za])
      yz_extrude(sx*strap_x + s*(lidk_w/2+k_gap+bodyk_w/2) - bodyk_w/2, bodyk_w) body_knuckle2d();
}

module lid_knuckles() {
  for (sx = [-1, 1])
    translate([0, ya, za]) yz_extrude(sx*strap_x - lidk_w/2, lidk_w) lid_knuckle2d();
}

module pin_holes(d) {
  mirror_x() translate([strap_x-lidk_w/2-k_gap-bodyk_w-1, ya, za])
    rotate([0, 90, 0]) cylinder(d = d, h = lidk_w+2*k_gap+2*bodyk_w+2, $fn = 24);
}

// ---------------------------------------------------------------- body
module body() {
  difference() {
    union() {
      difference() {
        cbox([-W/2, 0, Hp], [W/2, D, zr]);
        translate([-W/2+wall, wall, Hp+floor_t]) cube([W-2*wall, D-2*wall, Hb]);
      }
      // corner blocks that take the screws
      mirror_x() hull() {
        translate([x_scr-3.5, y_scr-3.5, Hp+floor_t-eps]) cube([7+eps, 7+eps, 5.5]);
        translate([x_scr-2, y_scr-2, Hp+floor_t-eps]) cube([5.5+eps, 5.5+eps, 7]);
      }
      body_knuckles();
    }
    pin_holes(pin_d + pin_tight);
    // strap grooves and lock-plate recess on the front
    mirror_x() translate([strap_x-strap_w/2-0.2, -1, Hp-1]) cube([strap_w+0.4, 1+groove, Hb+2]);
    translate([-plate_w/2-0.2, -1, zr-1.5-plate_h-0.2]) cube([plate_w+0.4, 1+groove, plate_h+0.4]);
    // peg sockets and screw pilot holes underneath
    mirror_x() translate([x_peg, y_peg, Hp-eps]) cylinder(d = peg_d+0.5, h = 1.2);
    mirror_x() translate([x_scr, y_scr, Hp-eps]) cylinder(d = screw_d-0.5, h = 7.6, $fn = 24);
  }
}

// ---------------------------------------------------------------- lid
module lid_profile(ext = 0) {
  polygon([[0, -ext], [0, lid_v], [lid_s, lid_v+lid_s],
           [D-lid_s, lid_v+lid_s], [D, lid_v], [D, -ext]]);
}

// where straps run: front face, both shoulders and the top (not the back face)
module strap_clip() {
  translate([-5, 0]) square([5+D/2, 80]);
  translate([D/2-eps, lid_v]) square([D/2+5, 80]);
}

module lid() {
  difference() {
    union() {
      translate([0, 0, zr]) yz_extrude(-W/2, W) lid_profile();
      lid_knuckles();
    }
    translate([0, 0, zr]) yz_extrude(-W/2+wall, W-2*wall) offset(delta = -wall) lid_profile(10);
    pin_holes(pin_d + pin_loose);
    mirror_x() translate([0, 0, zr]) yz_extrude(strap_x-strap_w/2-0.2, strap_w+0.4)
      intersection() {
        difference() { offset(delta = 1) lid_profile(10); offset(delta = -groove) lid_profile(10); }
        strap_clip();
      }
  }
}

// ---------------------------------------------------------------- trim
module lid_strap2d() {
  intersection() {
    difference() {
      offset(delta = strap_t-groove+0.1) lid_profile(10);
      offset(delta = -groove+0.1) lid_profile(10);
    }
    strap_clip();
  }
}

module lid_strap(sx = 1) {
  translate([0, 0, zr]) yz_extrude(sx*strap_x-strap_w/2, strap_w) lid_strap2d();
}

module body_strap(sx = 1) {
  translate([sx*strap_x-strap_w/2, groove-0.1-strap_t, Hp+0.15]) cube([strap_w, strap_t, Hb-0.3]);
}

module lock_plate_flat() {
  difference() {
    union() {
      hull() for (x = [-plate_w/2+1.5, plate_w/2-1.5]) for (y = [1.5, plate_h-1.5])
        translate([x, y, 0]) cylinder(r = 1.5, h = plate_t, $fn = 24);
      for (x = [-plate_w/2+2, plate_w/2-2]) for (y = [2, plate_h-2])
        translate([x, y, plate_t-0.2]) sphere(r = 0.8, $fn = 16);
    }
    translate([0, plate_h*0.6, plate_t-0.6]) cylinder(d = 3.2, h = 1, $fn = 24);
    translate([0, 0, plate_t-0.6]) linear_extrude(1)
      polygon([[-0.7, plate_h*0.6], [0.7, plate_h*0.6], [1.3, plate_h*0.22], [-1.3, plate_h*0.22]]);
  }
}

module lock_plate() {
  translate([0, groove-0.1, zr-1.5-plate_h]) rotate([90, 0, 0]) lock_plate_flat();
}

module ball(clear = false) {
  translate(clear ? ball_free : ball_locked) sphere(d = ball_d);
}

module hinge_pins() {
  mirror_x() translate([strap_x-lidk_w/2-k_gap-bodyk_w, ya, za])
    rotate([0, 90, 0]) cylinder(d = pin_d, h = lidk_w+2*k_gap+2*bodyk_w, $fn = 16);
}

// ---------------------------------------------------------------- views
module lid_group(angle) {
  translate([0, ya, za]) rotate([-angle, 0, 0]) translate([0, -ya, -za]) {
    color(colA) lid();
    color(colB) { lid_strap(1); lid_strap(-1); }
  }
}

module assembly(angle = 0, out = 0, clear = false) {
  color(colB) plinth();
  color(colB) translate([0, -out, 0]) drawer();
  color(colA) body();
  color(colB) { body_strap(1); body_strap(-1); lock_plate(); }
  color("Silver") { ball(clear); hinge_pins(); }
  lid_group(angle);
}

module exploded() {
  color(colB) plinth();
  color(colB) translate([0, -42, 0]) drawer();
  color("Silver") translate([0, 0, 14]) ball();
  translate([0, 0, 26]) {
    color(colA) body();
    color(colB) translate([0, -10, 0]) { body_strap(1); body_strap(-1); lock_plate(); }
    color("Silver") translate([0, 12, 0]) hinge_pins();
  }
  translate([0, 0, 50]) {
    color(colA) lid();
    color(colB) translate([0, 0, 12]) { lid_strap(1); lid_strap(-1); }
  }
}

// plinth and drawer from above with the body removed, to show the lock
module lock_view() {
  color(colB) plinth();
  color("Peru") translate([0, -drawer_out, 0]) drawer();
  color("Silver") ball(ball_clear);
}

// ---------------------------------------------------------------- print layouts
module print_body() { translate([0, 0, -Hp]) body(); }
module print_lid() { translate([0, 0, -zr]) lid(); }
module print_plinth() { plinth(); }
module print_drawer() { translate([0, 0, -tf]) drawer(); }
module print_lid_strap() { linear_extrude(strap_w) lid_strap2d(); }
module print_body_strap() { cube([strap_w, Hb-0.3, strap_t]); }
module print_lock_plate() { lock_plate_flat(); }

module plate_a() {
  translate([-W/2-5, 0, 0]) print_body();
  translate([W/2+5, 0, 0]) print_lid();
}

module plate_b() {
  translate([0, 0, 0]) print_plinth();
  translate([0, -62, 0]) print_drawer();
  translate([PW/2+8, -40, 0]) print_lid_strap();
  translate([PW/2+8, 15, 0]) print_lid_strap();
  translate([-PW/2-16, 0, 0]) print_body_strap();
  translate([-PW/2-26, 0, 0]) print_body_strap();
  translate([-PW/2-20, -35, 0]) print_lock_plate();
}

if (part == "assembly") assembly(open_angle, drawer_out, ball_clear);
else if (part == "exploded") exploded();
else if (part == "lock_view") lock_view();
else if (part == "plate_a") plate_a();
else if (part == "plate_b") plate_b();
else if (in_place) {
  if (part == "body") body();
  else if (part == "lid") lid();
  else if (part == "plinth") plinth();
  else if (part == "drawer") drawer();
  else if (part == "lid_strap") lid_strap(1);
  else if (part == "body_strap") body_strap(1);
  else if (part == "lock_plate") lock_plate();
}
else if (part == "body") print_body();
else if (part == "lid") print_lid();
else if (part == "plinth") print_plinth();
else if (part == "drawer") print_drawer();
else if (part == "lid_strap") print_lid_strap();
else if (part == "body_strap") print_body_strap();
else if (part == "lock_plate") print_lock_plate();

// numbers that check.py needs
if (part == "info") {
  echo(str("INFO ya=", ya)); echo(str("INFO za=", za)); echo(str("INFO r=", r));
  echo(str("INFO tf=", tf)); echo(str("INFO th=", th)); echo(str("INFO Hp=", Hp));
  echo(str("INFO x_seat=", x_seat)); echo(str("INFO y_trc=", y_trc));
  echo(str("INFO x_tr0=", x_tr0)); echo(str("INFO x_tr1=", x_tr1));
  echo(str("INFO y_tr0=", y_tr0)); echo(str("INFO y_tr1=", y_tr1));
  echo(str("INFO x_slope0=", x_slope0)); echo(str("INFO slope_deg=", slope_deg));
  echo(str("INFO xh0=", xh0)); echo(str("INFO seat_sink=", seat_sink));
  echo(str("INFO drawer_inner=", 2*(xd-dwall), " x ", y_dback-dwall-y_dfront, " x ", zt-tf-dfloor));
  echo(str("INFO body_inner=", W-2*wall, " x ", D-2*wall, " x ", Hb-floor_t));
  echo(str("INFO outer=", PW, " x ", PD, " x ", zr+lid_v+lid_s));
}
