"""Draw the maps of the Biñan book as SVG files (book/layout/maps/).

Boundaries: Philippine Standard Geographic Code shapefiles (PSA / OCHA),
as TopoJSON from github.com/faeldon/philippines-json-maps (MIT licence),
copied in book/maps/data/.

Usage:  python3 book/maps/make_maps.py
"""

import math
import sys
from pathlib import Path

from shapely.geometry import LineString, MultiPolygon, Point, Polygon, box
from shapely.ops import unary_union

sys.path.insert(0, str(Path(__file__).resolve().parent))
import topo  # noqa: E402

DATA = Path(__file__).resolve().parent / "data"
OUT = Path(__file__).resolve().parent.parent / "layout" / "maps"

INK = "#2a2420"
MUTED = "#6f655c"
ACCENT = "#7a2e22"
GOLD = "#b08a3e"
PAPER = "#fbf8f1"
LAND = "#f3ecdd"
LAND_ALT = "#ebe1cc"
WATER = "#d9e4e4"
WATER_INK = "#5d7c80"
RULE = "#c9b996"
FONT = "Source Serif 4, Georgia, serif"


# --------------------------------------------------------------------------
# Data
# --------------------------------------------------------------------------

def municipalities():
    """Every city and municipality around Laguna de Bay, by name."""
    files = [
        "municities-provdist-403400000.topo.0.1.json",   # Laguna
        "municities-provdist-402100000.topo.0.1.json",   # Cavite
        "municities-provdist-405800000.topo.0.1.json",   # Rizal
        "municities-provdist-401000000.topo.0.1.json",   # Batangas
        "municities-provdist-1303900000.topo.0.1.json",  # Manila
        "municities-provdist-1307400000.topo.0.1.json",  # NCR 2nd district
        "municities-provdist-1307500000.topo.0.1.json",  # NCR 3rd district
        "municities-provdist-1307600000.topo.0.1.json",  # NCR 4th district
    ]
    result = {}
    for name in files:
        province = name.split("-")[2].split(".")[0]
        for props, geom in topo.load(DATA / name):
            label = props["adm3_en"].replace("City of ", "")
            result[label] = {"geom": geom.buffer(0), "province": province}
    return result


def barangays():
    return {p["adm4_en"]: g.buffer(0) for p, g in topo.load(
        DATA / "bgysubmuns-municity-403403000.topo.0.1.json")}


def luzon():
    for props, geom in topo.load(DATA / "islandgroups.topo.0.001.json"):
        if str(props.get("island_group", "")).lower() == "luzon":
            return geom.buffer(0)
    raise SystemExit("Luzon outline not found")


# --------------------------------------------------------------------------
# SVG helpers
# --------------------------------------------------------------------------

class Map:
    """An SVG canvas over a lon/lat box (equirectangular, cos(lat) scaled)."""

    def __init__(self, west, south, east, north, width=116.0, height=None):
        self.west, self.south, self.east, self.north = west, south, east, north
        self.kx = math.cos(math.radians((south + north) / 2))
        if height is not None:  # fit a tall map to a given height
            width = height * (east - west) * self.kx / (north - south)
        self.width = width
        self.scale = width / ((east - west) * self.kx)  # mm per degree lat
        self.height = (north - south) * self.scale
        self.items = []
        self.clip = box(west, south, east, north)

    def xy(self, lon, lat):
        return ((lon - self.west) * self.kx * self.scale,
                (self.north - lat) * self.scale)

    def km(self, km):
        return km / 111.32 * self.scale  # mm for a distance in km

    def path(self, geom):
        geom = geom.intersection(self.clip)
        polys = []
        if isinstance(geom, Polygon):
            polys = [geom]
        elif isinstance(geom, MultiPolygon):
            polys = list(geom.geoms)
        elif hasattr(geom, "geoms"):
            polys = [g for g in geom.geoms if isinstance(g, Polygon)]
        d = []
        for poly in polys:
            for ring in [poly.exterior, *poly.interiors]:
                pts = [self.xy(x, y) for x, y in ring.coords]
                d.append("M" + " L".join(f"{x:.2f},{y:.2f}" for x, y in pts) + " Z")
        return " ".join(d)

    def shape(self, geom, fill, stroke="none", width=0.25, extra=""):
        d = self.path(geom)
        if d:
            self.items.append(
                f'<path d="{d}" fill="{fill}" stroke="{stroke}" '
                f'stroke-width="{width}" stroke-linejoin="round" {extra}/>')

    def text(self, lon, lat, label, size=3.0, color=INK, anchor="middle",
             weight=400, italic=False, dx=0.0, dy=0.0, halo=True, spacing=0):
        x, y = self.xy(lon, lat)
        style = "italic" if italic else "normal"
        lines = label.split("\n")
        tspans = "".join(
            f'<tspan x="{x + dx:.2f}" dy="{0 if i == 0 else size * 1.1:.2f}">{line}</tspan>'
            for i, line in enumerate(lines))
        common = (f'font-family="{FONT}" font-size="{size}" text-anchor="{anchor}" '
                  f'font-weight="{weight}" font-style="{style}" letter-spacing="{spacing}"')
        if halo:
            self.items.append(
                f'<text x="{x + dx:.2f}" y="{y + dy:.2f}" {common} fill="none" '
                f'stroke="{PAPER}" stroke-width="{size * 0.28:.2f}" '
                f'stroke-linejoin="round" opacity="0.9">{tspans}</text>')
        self.items.append(
            f'<text x="{x + dx:.2f}" y="{y + dy:.2f}" {common} fill="{color}">{tspans}</text>')

    def dot(self, lon, lat, r=0.8, fill=INK, stroke=PAPER):
        x, y = self.xy(lon, lat)
        self.items.append(
            f'<circle cx="{x:.2f}" cy="{y:.2f}" r="{r}" fill="{fill}" '
            f'stroke="{stroke}" stroke-width="0.35"/>')

    def arrow(self, start, end, color=ACCENT, width=1.1, shorten=1.6):
        (x1, y1), (x2, y2) = self.xy(*start), self.xy(*end)
        length = math.hypot(x2 - x1, y2 - y1)
        ux, uy = (x2 - x1) / length, (y2 - y1) / length
        x2, y2 = x2 - ux * shorten, y2 - uy * shorten
        # slightly curved shaft
        mx, my = (x1 + x2) / 2 - uy * length * 0.08, (y1 + y2) / 2 + ux * length * 0.08
        head = 2.4
        hx, hy = x2 - ux * head, y2 - uy * head
        self.items.append(
            f'<path d="M{x1:.2f},{y1:.2f} Q{mx:.2f},{my:.2f} {hx:.2f},{hy:.2f}" '
            f'fill="none" stroke="{color}" stroke-width="{width}" stroke-linecap="round"/>')
        px, py = -uy * head * 0.55, ux * head * 0.55
        self.items.append(
            f'<path d="M{x2:.2f},{y2:.2f} L{hx + px:.2f},{hy + py:.2f} '
            f'L{hx - px:.2f},{hy - py:.2f} Z" fill="{color}"/>')

    def leader(self, start, end, color=MUTED):
        (x1, y1), (x2, y2) = self.xy(*start), self.xy(*end)
        self.items.append(
            f'<line x1="{x1:.2f}" y1="{y1:.2f}" x2="{x2:.2f}" y2="{y2:.2f}" '
            f'stroke="{color}" stroke-width="0.25"/>')

    def line(self, coords, color=INK, width=0.6, dash=None, opacity=1.0, cap="round"):
        """A polyline through (lon, lat) points."""
        pts = [self.xy(lon, lat) for lon, lat in coords]
        if len(pts) < 2:
            return
        d = "M" + " L".join(f"{x:.2f},{y:.2f}" for x, y in pts)
        dash_attr = f' stroke-dasharray="{dash}"' if dash else ""
        self.items.append(
            f'<path d="{d}" fill="none" stroke="{color}" stroke-width="{width}" '
            f'stroke-linecap="{cap}" stroke-linejoin="round" opacity="{opacity}"{dash_attr}/>')

    def circle(self, lon, lat, r_mm, fill="none", stroke=INK, width=0.3, opacity=1.0, dash=None):
        x, y = self.xy(lon, lat)
        dash_attr = f' stroke-dasharray="{dash}"' if dash else ""
        self.items.append(
            f'<circle cx="{x:.2f}" cy="{y:.2f}" r="{r_mm:.2f}" fill="{fill}" fill-opacity="{opacity}" '
            f'stroke="{stroke}" stroke-width="{width}"{dash_attr}/>')

    def raw(self, svg):
        self.items.append(svg)

    def scale_bar(self, km, x=None, y=None):
        x = 4 if x is None else x
        y = self.height - 5 if y is None else y
        w = self.km(km)
        half = w / 2
        self.raw(
            f'<g font-family="{FONT}" font-size="2.6" fill="{INK}">'
            f'<rect x="{x}" y="{y}" width="{half:.2f}" height="1" fill="{INK}"/>'
            f'<rect x="{x + half:.2f}" y="{y}" width="{half:.2f}" height="1" fill="{PAPER}" '
            f'stroke="{INK}" stroke-width="0.25"/>'
            f'<text x="{x}" y="{y - 1}">0</text>'
            f'<text x="{x + w:.2f}" y="{y - 1}" text-anchor="middle">{km} km</text></g>')

    def north_arrow(self, x=None, y=7):
        x = self.width - 6 if x is None else x
        self.raw(
            f'<g fill="{INK}" font-family="{FONT}">'
            f'<path d="M{x},{y - 4} L{x + 1.6},{y + 1.5} L{x},{y + 0.6} '
            f'L{x - 1.6},{y + 1.5} Z"/>'
            f'<text x="{x}" y="{y + 5}" font-size="2.8" text-anchor="middle">N</text></g>')

    def save(self, name):
        OUT.mkdir(parents=True, exist_ok=True)
        body = "\n".join(self.items)
        svg = (f'<svg xmlns="http://www.w3.org/2000/svg" width="{self.width:.1f}mm" '
               f'height="{self.height:.1f}mm" viewBox="0 0 {self.width:.2f} {self.height:.2f}">\n'
               f'<rect width="100%" height="100%" fill="{WATER}"/>\n{body}\n'
               f'<rect width="100%" height="100%" fill="none" stroke="{RULE}" stroke-width="0.5"/>\n'
               f'</svg>\n')
        (OUT / f"{name}.svg").write_text(svg, encoding="utf-8")
        print(f"wrote book/layout/maps/{name}.svg")


def centre(geom):
    p = geom.representative_point()
    return p.x, p.y


def dot_pattern(pid, color, spacing=1.3, radius=0.28):
    """Dots: a fill that stays distinct from hatching when printed in black and white."""
    return (f'<defs><pattern id="{pid}" patternUnits="userSpaceOnUse" width="{spacing}" '
            f'height="{spacing}"><circle cx="{spacing / 2}" cy="{spacing / 2}" r="{radius}" '
            f'fill="{color}"/></pattern></defs>')


def hatch_pattern(pid, color, spacing=1.6, angle=45, width=0.3):
    return (f'<defs><pattern id="{pid}" patternUnits="userSpaceOnUse" width="{spacing}" '
            f'height="{spacing}" patternTransform="rotate({angle})">'
            f'<line x1="0" y1="0" x2="0" y2="{spacing}" stroke="{color}" '
            f'stroke-width="{width}"/></pattern></defs>')


# --------------------------------------------------------------------------
# The maps
# --------------------------------------------------------------------------

def map_location(munis, luzon_shape):
    """Biñan on the western shore of Laguna de Bay, south of Manila."""
    m = Map(120.86, 14.08, 121.56, 14.70)
    for name, info in munis.items():
        fill = LAND_ALT if info["province"] == "403400000" else LAND
        m.shape(info["geom"], fill, stroke=RULE, width=0.18)
    laguna = unary_union([i["geom"] for i in munis.values() if i["province"] == "403400000"])
    m.shape(laguna, "none", stroke=MUTED, width=0.45, extra='stroke-dasharray="1.6,0.8"')
    m.shape(munis["Biñan"]["geom"], ACCENT, stroke=PAPER, width=0.3)

    m.text(121.165, 14.455, "Laguna de Bay", size=4.4, italic=True, color=WATER_INK, halo=False, spacing=0.15)
    m.text(120.905, 14.555, "Manila\nBay", size=3.4, italic=True, color=WATER_INK, halo=False)

    places = [  # name, label, anchor, dx, dy
        ("Manila", "Manila", "middle", -2.0, 4.4),
        ("Muntinlupa", "Muntinlupa", "start", 1.8, 1.0),
        ("San Pedro", "San Pedro", "end", -1.8, -0.6),
        ("Carmona", "Carmona", "end", -1.8, 2.2),
        ("Santa Rosa", "Santa Rosa", "start", 1.8, 1.0),
        ("Cabuyao", "Cabuyao (old Tabuco)", "start", 1.8, 1.0),
        ("Calamba", "Calamba", "start", 1.8, 1.0),
        ("Los Baños", "Los Baños", "start", 1.8, 1.0),
        ("Santa Cruz", "Santa Cruz", "start", 1.8, 1.0),
    ]
    for key, label, anchor, dx, dy in places:
        lon, lat = centre(munis[key]["geom"])
        m.dot(lon, lat, r=0.65)
        m.text(lon, lat, label, size=3.0, anchor=anchor, dx=dx, dy=dy,
               weight=600 if key == "Manila" else 400)
    poblacion = centre(barangays()["Poblacion"])
    m.dot(*poblacion, r=0.9, fill=PAPER, stroke=INK)
    m.text(*poblacion, "BIÑAN", size=3.9, weight=700, color=ACCENT, anchor="start",
           dx=2.2, dy=-1.2, spacing=0.3)
    m.text(121.43, 14.17, "L A G U N A", size=3.2, color=MUTED, spacing=0.6)
    m.text(120.93, 14.24, "C A V I T E", size=3.2, color=MUTED, spacing=0.6)
    m.text(121.20, 14.60, "R I Z A L", size=3.2, color=MUTED, spacing=0.6)
    m.text(121.04, 14.66, "METRO MANILA", size=2.8, color=MUTED, spacing=0.4)

    # Inset: Luzon with the area of the main map.
    inset_w, pad = 30.0, 3.0
    lw, ls, le, ln = 119.4, 12.2, 124.6, 18.8
    kx = math.cos(math.radians(15.5))
    s = inset_w / ((le - lw) * kx)
    inset_h = (ln - ls) * s
    ox, oy = m.width - inset_w - pad, pad

    def ixy(lon, lat):
        return ox + (lon - lw) * kx * s, oy + (ln - lat) * s

    d = []
    shape = luzon_shape.intersection(box(lw, ls, le, ln))
    for poly in getattr(shape, "geoms", [shape]):
        if not isinstance(poly, Polygon) or poly.area < 0.004:
            continue
        d.append("M" + " L".join(f"{x:.2f},{y:.2f}" for x, y in
                                 (ixy(a, b) for a, b in poly.exterior.coords)) + " Z")
    bx1, by1 = ixy(m.west, m.north)
    bx2, by2 = ixy(m.east, m.south)
    m.raw(f'<rect x="{ox}" y="{oy}" width="{inset_w}" height="{inset_h:.2f}" fill="{WATER}" '
          f'stroke="{RULE}" stroke-width="0.4"/>'
          f'<path d="{" ".join(d)}" fill="{LAND}" stroke="{MUTED}" stroke-width="0.15"/>'
          f'<rect x="{bx1:.2f}" y="{by1:.2f}" width="{bx2 - bx1:.2f}" height="{by2 - by1:.2f}" '
          f'fill="none" stroke="{ACCENT}" stroke-width="0.6"/>'
          f'<text x="{ox + 2}" y="{oy + 4}" font-family="{FONT}" font-size="2.8" '
          f'fill="{MUTED}" font-style="italic">Luzon</text>')
    m.scale_bar(10, x=m.width - 30)
    m.north_arrow(x=6, y=m.height - 16)
    m.save("location")


def map_territory(munis):
    """How the territory of Biñan was carved out and changed, 1688–1907."""
    m = Map(120.975, 14.185, 121.215, 14.465)
    m.raw(hatch_pattern("h-tabuco", GOLD, spacing=1.5, width=0.35))
    m.raw(hatch_pattern("h-sanpedro", ACCENT, spacing=1.5, angle=-45, width=0.3))
    m.raw(hatch_pattern("h-muntinlupa", MUTED, spacing=1.8, angle=-45, width=0.25))
    for name, info in munis.items():
        m.shape(info["geom"], LAND, stroke=RULE, width=0.2)
    m.shape(munis["Cabuyao"]["geom"], "url(#h-tabuco)", stroke=GOLD, width=0.45)
    m.shape(munis["Santa Rosa"]["geom"], "#e4cfa8", stroke=ACCENT, width=0.35,
            extra='stroke-dasharray="1.2,0.7"')
    m.shape(munis["San Pedro"]["geom"], "url(#h-sanpedro)", stroke=ACCENT, width=0.35)
    m.shape(munis["Muntinlupa"]["geom"], "url(#h-muntinlupa)", stroke=MUTED, width=0.3)
    m.shape(munis["Biñan"]["geom"], ACCENT, stroke=PAPER, width=0.35)

    m.text(*centre(munis["Biñan"]["geom"]), "BIÑAN", size=3.6, weight=700, color=PAPER,
           halo=False, spacing=0.3)
    notes = [  # title, note, label position
        ("CABUYAO (old Tabuco)", "Biñan separated from Tabuco, 1688", (121.135, 14.205)),
        ("SANTA ROSA", "barrio of Biñan,\nseparated 1791–92", (121.105, 14.268)),
        ("SAN PEDRO", "merged with Biñan\n1903–1907", (121.034, 14.352)),
        ("MUNTINLUPA", "attached to Biñan\n1903–1905", (121.030, 14.420)),
    ]
    for title, note, (lon, lat) in notes:
        m.text(lon, lat, title, size=3.0, weight=600)
        m.text(lon, lat, note, size=2.6, italic=True, color=MUTED, dy=3.4)
    m.text(121.035, 14.268, "Carmona", size=2.8, color=MUTED)
    m.text(121.18, 14.42, "Laguna de Bay", size=4.0, italic=True, color=WATER_INK, halo=False)
    m.scale_bar(5)
    m.north_arrow()
    m.save("territory")


def map_liberation(munis, bgys):
    """The guerrilla attack on Biñan, 3 February 1945."""
    m = Map(120.995, 14.235, 121.175, 14.415)
    for name, info in munis.items():
        m.shape(info["geom"], LAND, stroke=RULE, width=0.2)
    m.shape(munis["Biñan"]["geom"], LAND_ALT, stroke=ACCENT, width=0.5)
    poblacion = bgys["Poblacion"]
    m.shape(poblacion, ACCENT, stroke=PAPER, width=0.25)
    target = centre(poblacion)
    for key, label, origin in [
        ("Carmona", "from CARMONA\n(west)", None),
        ("San Pedro", "from SAN PEDRO\n(north)", None),
        ("Santa Rosa", "from SANTA ROSA\n(south)", None),
    ]:
        lon, lat = centre(munis[key]["geom"])
        m.arrow((lon, lat), target)
        m.text(lon, lat, label, size=2.9, weight=600, dy=-3.5)
    m.text(target[0], target[1], "Poblacion\n(Japanese garrison)", size=2.7, anchor="start",
           dx=2.6, dy=1.2, italic=True)
    m.text(121.145, 14.395, "Laguna de Bay", size=3.8, italic=True, color=WATER_INK, halo=False)
    m.raw(f'<g font-family="{FONT}"><rect x="3" y="3" width="52" height="11.5" fill="{PAPER}" '
          f'stroke="{RULE}" stroke-width="0.3"/>'
          f'<text x="5" y="8" font-size="3.4" font-weight="600" fill="{ACCENT}">3 February 1945</text>'
          f'<text x="5" y="12" font-size="2.6" fill="{INK}" font-style="italic">'
          f'Hunters ROTC guerrillas free Biñan in one day</text></g>')
    m.scale_bar(3)
    m.north_arrow()
    m.save("liberation")


def map_barangays(bgys, munis):
    """The 24 barangays of Biñan today."""
    city = munis["Biñan"]["geom"]
    w, s, e, n = city.bounds
    padx, pady = (e - w) * 0.06, (n - s) * 0.05
    m = Map(w - padx, s - pady, e + padx + 0.032, n + pady, height=150)
    m.raw(hatch_pattern("h-industry", MUTED, spacing=1.4, width=0.3))
    for name, info in munis.items():
        m.shape(info["geom"], LAND, stroke=RULE, width=0.2)
    for name, geom in bgys.items():
        m.shape(geom, LAND_ALT, stroke=PAPER, width=0.45)
    for name in ("Biñan", "Loma", "Timbao", "Mampalasan", "Ganado"):
        m.shape(bgys[name], "url(#h-industry)", stroke="none")
    m.shape(bgys["Poblacion"], ACCENT, stroke=PAPER, width=0.45)
    m.shape(city, "none", stroke=INK, width=0.5)

    # Small barangays of the old town centre are labelled outside, with a leader line.
    callouts = {
        "Casile": (121.106, 14.353), "San Jose": (121.106, 14.346),
        "Poblacion": (121.106, 14.339), "Santo Domingo": (121.106, 14.332),
        "Santo Niño": (121.106, 14.325),
    }
    for name, geom in bgys.items():
        lon, lat = centre(geom)
        if name in callouts:
            tlon, tlat = callouts[name]
            m.leader((lon, lat), (tlon, tlat))
            anchor = "start" if tlon > lon else "end"
            pad = 0.8 if anchor == "start" else -0.8
            m.text(tlon, tlat, name, size=2.5, anchor=anchor, dx=pad, dy=0.8,
                   weight=700 if name == "Poblacion" else 400,
                   color=ACCENT if name == "Poblacion" else INK)
            continue
        size = 2.4 if geom.area < 1e-4 else 2.8
        m.text(lon, lat, name.replace(" ", "\u00a0"), size=size, dy=size * 0.35)

    for key, label in [("Santa Rosa", "SANTA ROSA"), ("San Pedro", "SAN PEDRO"),
                       ("Carmona", "CARMONA"), ("Cabuyao", "CABUYAO")]:
        geom = munis[key]["geom"].intersection(m.clip)
        if not geom.is_empty:
            m.text(*centre(geom), label, size=2.6, color=MUTED, spacing=0.3)

    # Legend
    lx, ly = m.width - 59, m.height - 21
    m.raw(f'<g font-family="{FONT}" font-size="2.6" fill="{INK}">'
          f'<rect x="{lx}" y="{ly}" width="56" height="18" fill="{PAPER}" stroke="{RULE}" stroke-width="0.3"/>'
          f'<rect x="{lx + 2}" y="{ly + 2.5}" width="4" height="2.6" fill="{ACCENT}"/>'
          f'<text x="{lx + 8}" y="{ly + 4.8}">Poblacion – heritage town centre</text>'
          f'<rect x="{lx + 2}" y="{ly + 7.5}" width="4" height="2.6" fill="url(#h-industry)" stroke="{MUTED}" stroke-width="0.2"/>'
          f'<text x="{lx + 8}" y="{ly + 9.8}">Barangays with industrial parks</text>'
          f'<line x1="{lx + 2}" y1="{ly + 14}" x2="{lx + 6}" y2="{ly + 14}" stroke="{INK}" stroke-width="0.5"/>'
          f'<text x="{lx + 8}" y="{ly + 14.8}">City limits</text></g>')
    m.scale_bar(2, x=4, y=m.height - 5)
    m.north_arrow(x=7, y=8)
    m.save("barangays")


# --------------------------------------------------------------------------
# Geography chapter
# --------------------------------------------------------------------------

import json  # noqa: E402
from shapely.geometry import shape as geo_shape  # noqa: E402

# Active volcanoes (PHIVOLCS list), approximate summit coordinates (lon, lat).
VOLCANOES = {
    "Pinatubo": (120.35, 15.13), "Taal": (120.99, 14.00), "Banahaw": (121.49, 14.07),
    "Mayon": (123.69, 13.26), "Bulusan": (124.05, 12.77), "Kanlaon": (123.13, 10.41),
    "Hibok-Hibok": (124.67, 9.20), "Smith": (121.92, 19.52), "Didicas": (122.20, 19.08),
    "Biliran": (124.54, 11.52), "Musuan": (125.07, 7.88), "Matutum": (125.07, 6.36),
}


def lines_of(path, keep=lambda props: True):
    data = json.load(open(DATA / path, encoding="utf-8"))
    for feature in data["features"]:
        if keep(feature["properties"]):
            yield feature["properties"], geo_shape(feature["geometry"])


def line_path(m, geom):
    parts = getattr(geom, "geoms", [geom])
    d = []
    for part in parts:
        pts = [m.xy(x, y) for x, y in part.coords]
        d.append("M" + " L".join(f"{x:.2f},{y:.2f}" for x, y in pts))
    return " ".join(d)


def volcano(m, lon, lat, size=1.6, fill=ACCENT):
    x, y = m.xy(lon, lat)
    m.raw(f'<path d="M{x:.2f},{y - size:.2f} L{x + size * 0.9:.2f},{y + size * 0.7:.2f} '
          f'L{x - size * 0.9:.2f},{y + size * 0.7:.2f} Z" fill="{fill}" stroke="{PAPER}" stroke-width="0.25"/>')


def map_tectonics():
    """The Philippines between the Philippine Sea Plate and the Sunda Plate."""
    m = Map(116.2, 4.6, 128.4, 21.0, height=150)
    for props, geom in topo.load(DATA / "islandgroups.topo.0.001.json"):
        m.shape(geom.buffer(0), LAND, stroke=MUTED, width=0.2)
    for props, geom in lines_of("gem_active_faults_philippines.geojson"):
        is_wvf = props.get("name") == "West Valley Fault"
        m.raw(f'<path d="{line_path(m, geom)}" fill="none" stroke="{ACCENT if is_wvf else MUTED}" '
              f'stroke-width="{1.6 if is_wvf else 0.3}" opacity="{1 if is_wvf else 0.8}"/>')
        if is_wvf:  # a double line: still distinct when printed in black and white
            m.raw(f'<path d="{line_path(m, geom)}" fill="none" stroke="{PAPER}" stroke-width="0.5"/>')
    for props, geom in lines_of("pb2002_boundaries_philippines.geojson",
                                lambda p: p.get("Type") == "subduction"):
        sub = True
        m.raw(f'<path d="{line_path(m, geom)}" fill="none" stroke="{INK}" '
              f'stroke-width="{1.0 if sub else 0.6}" stroke-dasharray="{"none" if sub else "1.5,1"}"/>')
    for name, (lon, lat) in VOLCANOES.items():
        volcano(m, lon, lat)
    m.text(120.99, 14.00, "Taal", size=2.5, anchor="end", dx=-1.8, dy=3.4)
    m.text(120.35, 15.13, "Pinatubo", size=2.5, anchor="end", dx=-1.8, dy=1)
    m.text(123.69, 13.26, "Mayon", size=2.5, anchor="start", dx=1.8, dy=1)
    m.text(123.13, 10.41, "Kanlaon", size=2.5, anchor="end", dx=-1.8, dy=1)
    # Biñan
    m.dot(121.08, 14.34, r=1.0, fill=PAPER, stroke=ACCENT)
    m.text(121.08, 14.34, "BIÑAN", size=3.0, weight=700, color=ACCENT, anchor="start", dx=2, dy=-1)
    # Plates and structures
    m.text(126.6, 17.3, "PHILIPPINE\nSEA PLATE", size=3.6, weight=600, color=MUTED, spacing=0.3, halo=False)
    m.text(117.6, 11.0, "SUNDA\n(EURASIAN)\nPLATE", size=3.6, weight=600, color=MUTED, spacing=0.3, halo=False)
    m.text(118.6, 16.4, "Manila\nTrench", size=2.7, italic=True, color=INK, halo=False)
    m.text(127.55, 9.4, "Philippine\nTrench", size=2.7, italic=True, color=INK, halo=False)
    m.text(125.0, 11.0, "Philippine\nFault", size=2.8, italic=True, color=INK, anchor="start", halo=True)
    m.text(117.4, 7.0, "South China Sea", size=3.0, italic=True, color=WATER_INK, halo=False)
    m.text(126.9, 20.2, "Pacific Ocean", size=3.0, italic=True, color=WATER_INK, halo=False)
    m.arrow((127.9, 15.2), (125.9, 16.0), color=INK, width=0.9, shorten=0)
    m.arrow((116.7, 13.6), (118.5, 14.1), color=INK, width=0.9, shorten=0)
    m.text(127.5, 14.3, "~8–9 cm/yr", size=2.4, italic=True, halo=False)
    # Legend
    lx, ly = 3, m.height - 25
    m.raw(f'<g font-family="{FONT}" font-size="2.5" fill="{INK}">'
          f'<rect x="{lx}" y="{ly}" width="44" height="22" fill="{PAPER}" stroke="{RULE}" stroke-width="0.3"/>'
          f'<line x1="{lx + 2}" y1="{ly + 4}" x2="{lx + 8}" y2="{ly + 4}" stroke="{INK}" stroke-width="1"/>'
          f'<text x="{lx + 10}" y="{ly + 4.8}">Subduction trench</text>'
          f'<line x1="{lx + 2}" y1="{ly + 8.5}" x2="{lx + 8}" y2="{ly + 8.5}" stroke="{MUTED}" stroke-width="0.4"/>'
          f'<text x="{lx + 10}" y="{ly + 9.3}">Active fault</text>'
          f'<line x1="{lx + 2}" y1="{ly + 13}" x2="{lx + 8}" y2="{ly + 13}" stroke="{ACCENT}" stroke-width="1.6"/>'
          f'<line x1="{lx + 2}" y1="{ly + 13}" x2="{lx + 8}" y2="{ly + 13}" stroke="{PAPER}" stroke-width="0.5"/>'
          f'<text x="{lx + 10}" y="{ly + 13.8}">West Valley Fault</text>'
          f'<path d="M{lx + 5},{ly + 16.2} L{lx + 6.4},{ly + 18.6} L{lx + 3.6},{ly + 18.6} Z" fill="{ACCENT}"/>'
          f'<text x="{lx + 10}" y="{ly + 18.6}">Active volcano</text></g>')
    m.scale_bar(200, x=m.width - 34, y=m.height - 5)
    m.north_arrow(x=7, y=8)
    m.save("tectonics")


def map_westbay(munis):
    """Biñan, the West Valley Fault, the Laguna caldera and the old sea link."""
    m = Map(120.86, 13.96, 121.56, 14.70)
    for name, info in munis.items():
        m.shape(info["geom"], LAND, stroke=RULE, width=0.18)
    m.shape(munis["Biñan"]["geom"], ACCENT, stroke=PAPER, width=0.3)
    # Laguna caldera, approximate outline (central lobe between Talim and Jala-Jala)
    cx, cy = m.xy(121.27, 14.40)
    rx, ry = m.km(11), m.km(6)
    m.raw(f'<ellipse cx="{cx:.2f}" cy="{cy:.2f}" rx="{rx:.2f}" ry="{ry:.2f}" transform="rotate(-35 {cx:.2f} {cy:.2f})" '
          f'fill="none" stroke="{GOLD}" stroke-width="0.8" stroke-dasharray="2,1.2"/>')
    m.text(121.36, 14.47, "Laguna caldera\n(approximate)", size=2.7, italic=True, color="#8a6a26", anchor="start", halo=True)
    for props, geom in lines_of("gem_active_faults_philippines.geojson",
                                lambda p: p.get("name") in ("West Valley Fault", "East Valley Fault")):
        m.raw(f'<path d="{line_path(m, geom)}" fill="none" stroke="{ACCENT}" stroke-width="1.0"/>')
    m.text(121.005, 14.56, "West Valley\nFault", size=2.7, weight=600, color=ACCENT, anchor="end", halo=True)
    # Former sea link, 6,600-4,700 years ago (schematic)
    m.arrow((120.93, 14.50), (121.04, 14.43), color=WATER_INK, width=0.8, shorten=0)
    m.text(120.905, 14.44, "sea link\n6,600–4,700\nyears ago", size=2.5, italic=True, color=WATER_INK, anchor="middle", halo=True)
    for name, (lon, lat) in {"Taal": (120.99, 14.00), "Makiling": (121.20, 14.13), "Banahaw": (121.49, 14.07)}.items():
        volcano(m, lon, lat, size=2.0)
        right = lon > 121.4
        m.text(lon, lat, name, size=2.8, anchor="end" if right else "start",
               dx=-2.2 if right else 2.2, dy=1, weight=600)
    poblacion = centre(barangays()["Poblacion"])
    m.dot(*poblacion, r=0.9, fill=PAPER, stroke=INK)
    m.text(*poblacion, "BIÑAN", size=3.4, weight=700, color=ACCENT, anchor="start", dx=2, dy=-1.4, spacing=0.3)
    m.text(121.17, 14.55, "Laguna de Bay", size=4.0, italic=True, color=WATER_INK, halo=False)
    m.text(120.90, 14.60, "Manila\nBay", size=3.2, italic=True, color=WATER_INK, halo=False)
    m.text(120.985, 14.65, "Manila", size=2.8, weight=600, anchor="start", dx=1)
    m.scale_bar(10, x=m.width - 30)
    m.north_arrow(x=7, y=m.height - 16)
    m.save("westbay")


def chart_climate():
    """Monthly rainfall (PAGASA normals) and annual temperature (NASA POWER) for Biñan."""
    rain = [16.9, 11.1, 9.4, 18.5, 139.1, 264.5, 422.4, 457.5, 341.8, 224.3, 110.5, 62.7]
    power = json.load(open(Path(__file__).resolve().parent.parent / "research" /
                           "nasa_power_binan_1981_2024.json"))["properties"]["parameter"]["T2M"]
    years = list(range(1981, 2025))
    temps = [power[f"{y}13"] for y in years]
    W, H = 116.0, 62.0
    items = [f'<rect width="{W}" height="{H}" fill="{PAPER}"/>']
    font = f'font-family="{FONT}"'
    # Panel 1: rainfall
    x0, y0, pw, ph = 8, 10, 46, 40
    items.append(f'<text x="{x0}" y="5.5" {font} font-size="3" font-weight="600" fill="{ACCENT}">Rainfall by month (mm)</text>')
    for k, v in enumerate(rain):
        bw = pw / 12
        h = v / 480 * ph
        items.append(f'<rect x="{x0 + k * bw + 0.4:.2f}" y="{y0 + ph - h:.2f}" width="{bw - 0.8:.2f}" height="{h:.2f}" '
                     f'fill="{WATER_INK if 4 <= k <= 9 else RULE}"/>')
        items.append(f'<text x="{x0 + k * bw + bw / 2:.2f}" y="{y0 + ph + 3.2:.2f}" {font} font-size="2.3" '
                     f'text-anchor="middle" fill="{MUTED}">{"JFMAMJJASOND"[k]}</text>')
    for v in (0, 200, 400):
        yy = y0 + ph - v / 480 * ph
        items.append(f'<line x1="{x0}" x2="{x0 + pw}" y1="{yy:.2f}" y2="{yy:.2f}" stroke="{RULE}" stroke-width="0.15"/>')
        items.append(f'<text x="{x0 - 1}" y="{yy + 0.8:.2f}" {font} font-size="2.2" text-anchor="end" fill="{MUTED}">{v}</text>')
    items.append(f'<text x="{x0}" y="{H - 2}" {font} font-size="2.2" font-style="italic" fill="{MUTED}">PAGASA normals, Sangley Point</text>')
    # Panel 2: temperature
    x1, pw2 = 66, 46
    lo, hi = 25.2, 26.8
    items.append(f'<text x="{x1}" y="5.5" {font} font-size="3" font-weight="600" fill="{ACCENT}">Mean yearly temperature (°C)</text>')
    for v in (25.5, 26.0, 26.5):
        yy = y0 + ph - (v - lo) / (hi - lo) * ph
        items.append(f'<line x1="{x1}" x2="{x1 + pw2}" y1="{yy:.2f}" y2="{yy:.2f}" stroke="{RULE}" stroke-width="0.15"/>')
        items.append(f'<text x="{x1 - 1}" y="{yy + 0.8:.2f}" {font} font-size="2.2" text-anchor="end" fill="{MUTED}">{v}</text>')
    pts = [(x1 + (y - 1981) / 43 * pw2, y0 + ph - (t - lo) / (hi - lo) * ph) for y, t in zip(years, temps)]
    items.append('<polyline points="' + " ".join(f"{x:.2f},{y:.2f}" for x, y in pts) +
                 f'" fill="none" stroke="{ACCENT}" stroke-width="0.5"/>')
    n = len(years); mx = sum(years) / n; my = sum(temps) / n
    b = sum((x - mx) * (t - my) for x, t in zip(years, temps)) / sum((x - mx) ** 2 for x in years)
    t0, t1 = my + b * (1981 - mx), my + b * (2024 - mx)
    items.append(f'<line x1="{x1}" x2="{x1 + pw2}" y1="{y0 + ph - (t0 - lo) / (hi - lo) * ph:.2f}" '
                 f'y2="{y0 + ph - (t1 - lo) / (hi - lo) * ph:.2f}" stroke="{INK}" stroke-width="0.35" stroke-dasharray="1.2,0.8"/>')
    for y in (1981, 2000, 2024):
        xx = x1 + (y - 1981) / 43 * pw2
        items.append(f'<text x="{xx:.2f}" y="{y0 + ph + 3.2:.2f}" {font} font-size="2.3" text-anchor="middle" fill="{MUTED}">{y}</text>')
    items.append(f'<text x="{x1 + pw2}" y="{y0 + 3}" {font} font-size="2.4" text-anchor="end" fill="{INK}" '
                 f'font-style="italic">trend ≈ +{b * 10:.2f} °C per decade</text>')
    items.append(f'<text x="{x1}" y="{H - 2}" {font} font-size="2.2" font-style="italic" fill="{MUTED}">NASA POWER (MERRA-2), 14.34°N 121.08°E</text>')
    OUT.mkdir(parents=True, exist_ok=True)
    (OUT / "climate.svg").write_text(
        f'<svg xmlns="http://www.w3.org/2000/svg" width="{W}mm" height="{H}mm" viewBox="0 0 {W} {H}">'
        + "".join(items) + "</svg>\n", encoding="utf-8")
    print("wrote book/layout/maps/climate.svg")


def chart_population():
    """Census population of Biñan, 1903–2024 (PSA)."""
    data = [(1903, 9563), (1918, 10692), (1939, 16238), (1948, 20794), (1960, 33309),
            (1970, 58290), (1975, 67444), (1980, 83684), (1990, 134553), (1995, 160206),
            (2000, 201186), (2007, 262735), (2010, 283396), (2015, 333028), (2020, 407437),
            (2024, 584479)]
    W, H = 116.0, 64.0
    font = f'font-family="{FONT}"'
    x0, y0, pw, ph = 14, 9, 96, 44
    top = 600000
    X = lambda y: x0 + (y - 1900) / 125 * pw
    Y = lambda v: y0 + ph - v / top * ph
    items = [f'<rect width="{W}" height="{H}" fill="{PAPER}"/>',
             f'<text x="{x0}" y="5" {font} font-size="3" font-weight="600" fill="{ACCENT}">Population of Biñan at each census, 1903–2024</text>']
    for v in range(0, top + 1, 100000):
        items.append(f'<line x1="{x0}" x2="{x0 + pw}" y1="{Y(v):.2f}" y2="{Y(v):.2f}" stroke="{RULE}" stroke-width="0.15"/>')
        items.append(f'<text x="{x0 - 1}" y="{Y(v) + 0.8:.2f}" {font} font-size="2.2" text-anchor="end" fill="{MUTED}">{v:,}</text>')
    for y in (1900, 1925, 1950, 1975, 2000, 2025):
        items.append(f'<text x="{X(y):.2f}" y="{y0 + ph + 3.4:.2f}" {font} font-size="2.3" text-anchor="middle" fill="{MUTED}">{y}</text>')
    for a, b, label in [(1941.9, 1945.1, "war")]:
        if label:
            items.append(f'<rect x="{X(a):.2f}" y="{y0}" width="{X(b) - X(a):.2f}" height="{ph}" fill="{RULE}" opacity="0.5"/>')
            items.append(f'<text x="{X(a) - 0.6:.2f}" y="{y0 + 4}" {font} font-size="2.2" text-anchor="end" font-style="italic" fill="{MUTED}">{label}</text>')
    pts = " ".join(f"{X(y):.2f},{Y(v):.2f}" for y, v in data)
    items.append(f'<polyline points="{pts}" fill="none" stroke="{ACCENT}" stroke-width="0.6"/>')
    for y, v in data:
        items.append(f'<circle cx="{X(y):.2f}" cy="{Y(v):.2f}" r="0.6" fill="{ACCENT}"/>')
    notes = [(1948, 20794, "20,794 (1948)", "end", -1.5, -2.5),
             (1990, 134553, "Laguna Technopark opens, 1990", "end", -1.5, -2.5),
             (2010, 283396, "City, 2010", "start", 1.5, 3.5),
             (2024, 584479, "584,479 (2024): first in Laguna", "end", -1.5, 0.8)]
    for y, v, t, anchor, dx, dy in notes:
        items.append(f'<text x="{X(y) + dx:.2f}" y="{Y(v) + dy:.2f}" {font} font-size="2.4" text-anchor="{anchor}" fill="{INK}">{t}</text>')
    items.append(f'<text x="{x0}" y="{H - 2}" {font} font-size="2.2" font-style="italic" fill="{MUTED}">Philippine censuses (PSA).</text>')
    OUT.mkdir(parents=True, exist_ok=True)
    (OUT / "population.svg").write_text(
        f'<svg xmlns="http://www.w3.org/2000/svg" width="{W}mm" height="{H}mm" viewBox="0 0 {W} {H}">'
        + "".join(items) + "</svg>\n", encoding="utf-8")
    print("wrote book/layout/maps/population.svg")


# --------------------------------------------------------------------------
# Part II: transport, economy, land
# --------------------------------------------------------------------------

def osm_lines(road):
    """Road and railway lines from OpenStreetMap (ODbL), saved in data/osm_transport.geojson."""
    data = json.load(open(DATA / "osm_transport.geojson"))
    for feature in data["features"]:
        if feature["properties"]["road"] != road:
            continue
        geom = feature["geometry"]
        if geom["type"] == "LineString":
            yield geom["coordinates"]
        else:
            yield from geom["coordinates"]


def lines_geom(road):
    from shapely.geometry import MultiLineString
    return MultiLineString([LineString(c) for c in osm_lines(road) if len(c) > 1])


# Shore of Biñan (from the city boundary, simplified) and a schematic course for
# the planned Laguna Lakeshore Road Network, a little offshore.
LLRN = [(121.070, 14.384), (121.0790, 14.3690), (121.0840, 14.3625), (121.0930, 14.3580),
        (121.0955, 14.3535), (121.0975, 14.3495), (121.0980, 14.3433), (121.1010, 14.3402),
        (121.1050, 14.3360), (121.1110, 14.3270), (121.1180, 14.3150), (121.1260, 14.3020)]


def map_transport(munis, bgys):
    """Roads, expressways, railway and the lake around Biñan."""
    m = Map(121.000, 14.245, 121.135, 14.388)
    for name, info in munis.items():
        m.shape(info["geom"], LAND, stroke=RULE, width=0.2)
    m.shape(munis["Biñan"]["geom"], LAND_ALT, stroke=ACCENT, width=0.55)
    m.shape(bgys["Poblacion"], ACCENT, stroke=PAPER, width=0.25)

    road_red = "#b4533a"
    for coords in osm_lines("N65"):
        m.line(coords, PAPER, 1.5)
        m.line(coords, road_red, 0.8)
    for coords in osm_lines("N1"):
        m.line(coords, PAPER, 1.7)
        m.line(coords, road_red, 1.0)
    for road in ("SLEX", "CALAX"):
        for coords in osm_lines(road):
            m.line(coords, INK, 1.9)
            m.line(coords, GOLD, 0.9)
    for coords in osm_lines("NSCR"):
        m.line(coords, INK, 1.0)
        m.line(coords, PAPER, 0.5, dash="1.3,1.3", cap="butt")
    m.line(LLRN, ACCENT, 0.8, dash="2,1.2")

    slex, calax, n65 = lines_geom("SLEX"), lines_geom("CALAX"), lines_geom("N65")

    def on_slex(lon, lat):
        from shapely.ops import nearest_points
        point = nearest_points(slex, Point(lon, lat))[0]
        return point.x, point.y

    def crossing(a, b, near):
        hit = a.intersection(b)
        points = [hit] if hit.geom_type == "Point" else list(getattr(hit, "geoms", []))
        points = [p for p in points if p.geom_type == "Point"] or [Point(*on_slex(*near))]
        best = min(points, key=lambda p: p.distance(Point(*near)))
        return best.x, best.y

    exits = [("31 Southwoods", on_slex(121.0503, 14.3314), "end", -2.0, 0.6),
             ("33 Carmona (Cavite)", crossing(slex, n65, (121.0600, 14.3200)), "end", -2.0, 0.6),
             ("36 Mamplasan", crossing(slex, calax, (121.0790, 14.2925)), "start", 2.0, 1.2)]
    for label, (lon, lat), anchor, dx, dy in exits:
        m.dot(lon, lat, r=1.0, fill=PAPER, stroke=INK)
        m.text(lon, lat, label, size=2.4, anchor=anchor, dx=dx, dy=dy, weight=600)

    # Stations (Wikidata coordinates)
    for label, lon, lat, anchor, dx, dy in [
            ("Biñan station\n(closed 2023; NSCR c. 2028)", 121.081181, 14.3314, "start", 1.8, 2.6),
            ("Golden City 1", 121.063686, 14.332878, "start", 1.8, -1.2),
            ("San Pedro", 121.055278, 14.361667, "end", -1.8, 0.8),
            ("Santa Rosa", 121.1098, 14.3061, "start", 1.8, 0.8)]:
        x, y = m.xy(lon, lat)
        m.raw(f'<rect x="{x - 1.1:.2f}" y="{y - 1.1:.2f}" width="2.2" height="2.2" fill="{PAPER}" '
              f'stroke="{INK}" stroke-width="0.4"/>')
        m.text(lon, lat, label, size=2.3, anchor=anchor, dx=dx, dy=dy, italic=True)

    m.text(121.0843, 14.3384, "POBLACION", size=2.3, weight=700, color=ACCENT, anchor="start",
           dx=2.0, dy=-1.6, spacing=0.2)
    for label, lon, lat in [("SAN PEDRO", 121.028, 14.378), ("SANTA ROSA", 121.092, 14.262),
                            ("CARMONA", 121.020, 14.318), ("BIÑAN", 121.074, 14.300)]:
        m.text(lon, lat, label, size=3.0 if label == "BIÑAN" else 2.6,
               color=ACCENT if label == "BIÑAN" else MUTED, weight=700 if label == "BIÑAN" else 400,
               spacing=0.3)
    m.text(121.118, 14.366, "Laguna de Bay", size=3.8, italic=True, color=WATER_INK, halo=False)
    m.text(121.1035, 14.352, "lakeshore road\n(planned)", size=2.2, italic=True, color=ACCENT,
           anchor="start", dx=1.5)
    m.text(121.0620, 14.2790, "CALAX", size=2.4, weight=600, color=INK, anchor="start", dx=1.5)
    m.text(121.064, 14.360, "SLEX", size=2.4, weight=600, color=INK)
    m.text(121.093, 14.322, "N1", size=2.4, weight=600, color=road_red, anchor="start", dx=1.2)
    m.text(121.0705, 14.3285, "N65", size=2.2, weight=600, color=road_red)

    lx, ly = m.width - 61, m.height - 33
    m.raw(f'<g font-family="{FONT}" font-size="2.5" fill="{INK}">'
          f'<rect x="{lx}" y="{ly}" width="57" height="29" fill="{PAPER}" stroke="{RULE}" stroke-width="0.3"/>'
          f'<line x1="{lx + 2}" y1="{ly + 4}" x2="{lx + 8}" y2="{ly + 4}" stroke="{INK}" stroke-width="1.9"/>'
          f'<line x1="{lx + 2}" y1="{ly + 4}" x2="{lx + 8}" y2="{ly + 4}" stroke="{GOLD}" stroke-width="0.9"/>'
          f'<text x="{lx + 10}" y="{ly + 4.9}">Expressway, with exit number</text>'
          f'<line x1="{lx + 2}" y1="{ly + 9}" x2="{lx + 8}" y2="{ly + 9}" stroke="{road_red}" stroke-width="1.0"/>'
          f'<text x="{lx + 10}" y="{ly + 9.9}">National road (N1, N65)</text>'
          f'<line x1="{lx + 2}" y1="{ly + 14}" x2="{lx + 8}" y2="{ly + 14}" stroke="{INK}" stroke-width="1.0"/>'
          f'<line x1="{lx + 2}" y1="{ly + 14}" x2="{lx + 8}" y2="{ly + 14}" stroke="{PAPER}" stroke-width="0.5" stroke-dasharray="1.3,1.3"/>'
          f'<text x="{lx + 10}" y="{ly + 14.9}">Railway (rebuilt as the NSCR)</text>'
          f'<rect x="{lx + 3.9}" y="{ly + 17.9}" width="2.2" height="2.2" fill="{PAPER}" stroke="{INK}" stroke-width="0.4"/>'
          f'<text x="{lx + 10}" y="{ly + 19.9}">Railway station</text>'
          f'<line x1="{lx + 2}" y1="{ly + 24}" x2="{lx + 8}" y2="{ly + 24}" stroke="{ACCENT}" stroke-width="0.8" stroke-dasharray="2,1.2"/>'
          f'<text x="{lx + 10}" y="{ly + 24.9}">Lakeshore road (planned, schematic)</text></g>')
    m.scale_bar(2, x=4, y=m.height - 5)
    m.north_arrow()
    m.save("transport")


ECONOMY = {
    "industrial": ["Biñan", "Loma", "Malamig", "Mampalasan", "Ganado"],
    "centre": ["Poblacion", "San Vicente", "Santo Domingo", "San Jose", "Casile"],
    "lake": ["Malaban", "De La Paz"],
    "farm": ["Bungahan", "Langkiwa", "Platero", "Santo Tomas", "Timbao"],
}


def map_economy(bgys, munis):
    """Where Biñan makes a living: industry, trade, the last farms and the lake."""
    city = munis["Biñan"]["geom"]
    w, s, e, n = city.bounds
    m = Map(w - 0.006, s - 0.006, e + 0.035, n + 0.006, height=150)
    # Each zone has its own colour and its own texture, so that the map can be
    # read in the black-and-white edition too.
    m.raw(hatch_pattern("h-farm", "#4f6a2e", spacing=1.2, width=0.35))
    m.raw(hatch_pattern("h-ind1", "#6b5536", spacing=1.3, angle=45, width=0.3))
    m.raw(hatch_pattern("h-ind2", "#6b5536", spacing=1.3, angle=-45, width=0.3))
    m.raw(dot_pattern("h-centre", "#7a3a2e", spacing=1.1, radius=0.3))
    m.raw(hatch_pattern("h-lake", "#3d6b70", spacing=1.0, angle=0, width=0.28))
    textures = {"industrial": ["url(#h-ind1)", "url(#h-ind2)"], "centre": ["url(#h-centre)"],
                "lake": ["url(#h-lake)"]}
    for name, info in munis.items():
        m.shape(info["geom"], LAND, stroke=RULE, width=0.2)
    fills = {"industrial": "#c9b48e", "centre": "#d9a79b", "lake": "#b9d0d1"}
    for name, geom in bgys.items():
        fill, texture = LAND_ALT, []
        for kind, names in ECONOMY.items():
            if name in names and kind in fills:
                fill, texture = fills[kind], textures[kind]
        m.shape(geom, fill, stroke=PAPER, width=0.4)
        for t in texture:
            m.shape(geom, t, stroke="none")
    for name in ECONOMY["farm"]:
        m.shape(bgys[name], "url(#h-farm)", stroke="none")
    m.shape(city, "none", stroke=INK, width=0.5)
    # Fish pens and cages offshore: their exact place is not mapped here (the LLDA
    # zoning map was not available), so a symbol marks the waters off Malaban.
    for dlon, dlat in [(0, 0), (0.0035, 0.0012), (0.0016, -0.0026), (0.0052, -0.0016), (0.0030, 0.0040)]:
        x, y = m.xy(121.1040 + dlon, 14.3560 + dlat)
        m.raw(f'<rect x="{x - 1.3:.2f}" y="{y - 1.3:.2f}" width="2.6" height="2.6" fill="none" '
              f'stroke="{WATER_INK}" stroke-width="0.35"/>')
    m.text(121.1065, 14.3500, "fish pens and cages\n(offshore)", size=2.3, italic=True,
           color=WATER_INK, halo=False)

    def marker(lon, lat, symbol, label, anchor="start", dx=1.8, dy=0.9, size=2.4, color=INK):
        x, y = m.xy(lon, lat)
        if symbol == "mall":
            m.raw(f'<rect x="{x - 1.2:.2f}" y="{y - 1.2:.2f}" width="2.4" height="2.4" fill="{ACCENT}" '
                  f'stroke="{PAPER}" stroke-width="0.4"/>')
        elif symbol == "market":
            m.raw(f'<circle cx="{x:.2f}" cy="{y:.2f}" r="1.4" fill="{ACCENT}" stroke="{PAPER}" stroke-width="0.4"/>')
        elif symbol == "pig":
            m.raw(f'<path d="M{x - 1.1:.2f},{y - 1.1:.2f} L{x + 1.1:.2f},{y + 1.1:.2f} M{x + 1.1:.2f},{y - 1.1:.2f} '
                  f'L{x - 1.1:.2f},{y + 1.1:.2f}" stroke="{INK}" stroke-width="0.6"/>')
        elif symbol == "dam":
            m.raw(f'<path d="M{x - 1.4:.2f},{y + 1:.2f} L{x:.2f},{y - 1.2:.2f} L{x + 1.4:.2f},{y + 1:.2f} Z" '
                  f'fill="#6f8a4a" stroke="{PAPER}" stroke-width="0.3"/>')
        elif symbol == "office":
            m.raw(f'<rect x="{x - 1:.2f}" y="{y - 1:.2f}" width="2" height="2" fill="{PAPER}" '
                  f'stroke="{ACCENT}" stroke-width="0.5"/>')
        m.text(lon, lat, label, size=size, anchor=anchor, dx=dx, dy=dy, color=color)

    marker(121.0847, 14.3380, "market", "Public market", dx=2.0)
    marker(121.0896, 14.3288, "mall", "Pavilion Mall (1999)")
    marker(121.0821, 14.3323, "mall", "Central Mall (2011)", anchor="end", dx=-2.0)
    marker(121.0505, 14.3315, "mall", "Southwoods Mall\nand offices (2017)", anchor="end", dx=-2.0)
    marker(121.0870, 14.3246, "office", "One Asia IT Park")
    marker(121.0794, 14.3088, "office", "Unilab pharma campus", anchor="end", dx=-2.0)
    marker(121.0610, 14.2972, "pig", "hog farm\n(closed 2025)", anchor="end", dx=-1.8, size=2.2)
    marker(121.0723, 14.3142, "pig", "hog farm\n(closed 2025)", anchor="start", dx=1.8, size=2.2)
    marker(121.0449, 14.2780, "dam", "Timbao dam\n(irrigation)", anchor="end", dx=-2.0, size=2.2)
    m.text(121.0565, 14.2700, "LAGUNA\nTECHNOPARK", size=2.6, weight=700, color=INK, spacing=0.2)
    m.text(121.0790, 14.2880, "LIIP", size=2.6, weight=700, color=INK, spacing=0.2)
    for key, label in [("Santa Rosa", "SANTA ROSA"), ("San Pedro", "SAN PEDRO"), ("Carmona", "CARMONA")]:
        geom = munis[key]["geom"].intersection(m.clip)
        if not geom.is_empty:
            m.text(*centre(geom), label, size=2.6, color=MUTED, spacing=0.3)
    m.text(121.118, 14.372, "Laguna de Bay", size=3.4, italic=True, color=WATER_INK, halo=False)

    lx, ly = m.width - 58, m.height - 37
    m.raw(f'<g font-family="{FONT}" font-size="2.5" fill="{INK}">'
          f'<rect x="{lx}" y="{ly}" width="55" height="34" fill="{PAPER}" stroke="{RULE}" stroke-width="0.3"/>'
          f'<rect x="{lx + 2}" y="{ly + 2.5}" width="4" height="2.6" fill="#c9b48e"/>'
          f'<rect x="{lx + 2}" y="{ly + 2.5}" width="4" height="2.6" fill="url(#h-ind1)"/>'
          f'<rect x="{lx + 2}" y="{ly + 2.5}" width="4" height="2.6" fill="url(#h-ind2)"/>'
          f'<text x="{lx + 8}" y="{ly + 4.8}">Industrial parks (Technopark, LIIP)</text>'
          f'<rect x="{lx + 2}" y="{ly + 7.5}" width="4" height="2.6" fill="#d9a79b"/>'
          f'<rect x="{lx + 2}" y="{ly + 7.5}" width="4" height="2.6" fill="url(#h-centre)"/>'
          f'<text x="{lx + 8}" y="{ly + 9.8}">Old town centre: market and shops</text>'
          f'<rect x="{lx + 2}" y="{ly + 12.5}" width="4" height="2.6" fill="#b9d0d1"/>'
          f'<rect x="{lx + 2}" y="{ly + 12.5}" width="4" height="2.6" fill="url(#h-lake)"/>'
          f'<text x="{lx + 8}" y="{ly + 14.8}">Lakeshore barangays: fishing</text>'
          f'<rect x="{lx + 2}" y="{ly + 17.5}" width="4" height="2.6" fill="url(#h-farm)" stroke="#4f6a2e" stroke-width="0.2"/>'
          f'<text x="{lx + 8}" y="{ly + 19.8}">Barangays with remaining farmland</text>'
          f'<rect x="{lx + 2}" y="{ly + 22.5}" width="4" height="2.6" fill="{LAND_ALT}" stroke="{RULE}" stroke-width="0.2"/>'
          f'<text x="{lx + 8}" y="{ly + 24.8}">Mostly residential</text>'
          f'<rect x="{lx + 2.8}" y="{ly + 27.6}" width="2.4" height="2.4" fill="{ACCENT}"/>'
          f'<text x="{lx + 8}" y="{ly + 29.8}">Mall · circle: public market</text></g>')
    m.scale_bar(1, x=4, y=m.height - 5)
    m.north_arrow(x=7, y=8)
    m.save("economy")


# Largest documented landholdings (book/research/foncier.md): number, label,
# hectares (in Biñan where known), approximate position, colour.
HOLDINGS = [
    (1, "Laguna Technopark (Ayala Land), c. 240 ha in Biñan", 240, 121.0560, 14.2730, ACCENT),
    (2, "Southwoods City (Megaworld), 561 ha with Carmona", 561, 121.0420, 14.3290, ACCENT),
    (3, "Broadfield (Ayala Land), c. 120 ha", 120, 121.0490, 14.2618, ACCENT),
    (4, "Laguna Int. Industrial Park (Samsung, Solid), c. 117 ha", 117, 121.0790, 14.2880, ACCENT),
    (5, "Brentville (Filinvest), c. 53 ha", 53, 121.0880, 14.2985, ACCENT),
    (6, "De La Salle University campus, 50 ha", 50, 121.0428, 14.2609, GOLD),
    (7, "Meadowcrest (Federal Land), 48 ha", 48, 121.0450, 14.2740, ACCENT),
    (8, "Gonzalo Puyat &amp; Sons, 43.7 ha (court case 1998–2012)", 43.7, 121.0590, 14.2900, MUTED),
    (9, "Loma lots, c. 37 ha (tenants upheld, 2000)", 37, 121.0691, 14.2846, MUTED),
    (10, "Sevina Park (ArthaLand), 8.1 ha", 8.1, 121.0455, 14.2575, ACCENT),
]


def map_land(bgys, munis):
    """Who owns the land: the largest documented holdings, to scale, at approximate places."""
    city = munis["Biñan"]["geom"]
    w, s, e, n = city.bounds
    m = Map(w - 0.012, s - 0.006, e + 0.012, n + 0.006, height=150)
    m.raw(hatch_pattern("h-state", WATER_INK, spacing=1.0, width=0.3))
    for name, info in munis.items():
        m.shape(info["geom"], LAND, stroke=RULE, width=0.2)
    for name, geom in bgys.items():
        m.shape(geom, LAND_ALT, stroke=PAPER, width=0.4)
    m.shape(city, "none", stroke=INK, width=0.5)
    # Public land: the lake bed (at or below 12.5 m on the lake gauge) along the shore,
    # drawn as a schematic strip, and the railway right of way.
    shore = city.boundary.intersection(unary_union([bgys["Malaban"], bgys["De La Paz"]]).buffer(0.0005))
    strip = shore.buffer(0.0018).intersection(city.buffer(0.0002))
    m.shape(strip, "url(#h-state)", stroke=WATER_INK, width=0.25)
    for coords in osm_lines("NSCR"):
        m.line(coords, WATER_INK, 1.0)
        m.line(coords, PAPER, 0.45, dash="1.3,1.3", cap="butt")
    x, y = m.xy(121.0610, 14.2985)
    m.raw(f'<rect x="{x - 1.1:.2f}" y="{y - 1.1:.2f}" width="2.2" height="2.2" fill="{WATER_INK}" '
          f'stroke="{PAPER}" stroke-width="0.3"/>')
    m.text(121.0610, 14.2985, "The Meadows (rail\nrelocation, 2026)", size=2.2, anchor="end", dx=-1.8,
           dy=0.8, italic=True, color=WATER_INK)
    for number, label, hectares, lon, lat, color in HOLDINGS:
        radius_km = math.sqrt(hectares / 100 / math.pi)
        r = m.km(radius_km)
        dash = "1.2,0.8" if number == 2 else None
        m.circle(lon, lat, r, fill=color, stroke=color, width=0.4, opacity=0.28, dash=dash)
        m.text(lon, lat, str(number), size=2.6, weight=700, color=INK, dy=0.9)
    m.text(121.0880, 14.3470, "lake bed\n(public land)", size=2.3, italic=True, color=WATER_INK,
           anchor="end", dx=-1.0)
    for key, label in [("Santa Rosa", "SANTA ROSA"), ("San Pedro", "SAN PEDRO"), ("Carmona", "CARMONA")]:
        geom = munis[key]["geom"].intersection(m.clip)
        if not geom.is_empty:
            m.text(*centre(geom), label, size=2.6, color=MUTED, spacing=0.3)
    m.scale_bar(1, x=m.width - 22, y=m.height - 5)
    m.north_arrow()
    m.save("land")


# --------------------------------------------------------------------------
# Prehistory chapter
# --------------------------------------------------------------------------

def world_land():
    """Natural Earth land (1:50m), via world-atlas (ISC licence)."""
    topo_data = json.load(open(DATA / "world-land-50m.topo.json"))
    obj = topo_data["objects"]["land"]
    tmp = dict(topo_data)
    if obj["type"] != "GeometryCollection":
        obj = {"type": "GeometryCollection", "geometries": [obj]}
    tmp["objects"] = {"land": obj}
    path = Path(OUT.parent / "_land.json")
    path.write_text(json.dumps(tmp))
    shapes = [g for _, g in topo.load(path)]
    path.unlink()
    return unary_union([g.buffer(0) for g in shapes])


def curved_arrow(m, points, color, width=0.9, dash=None):
    """A smooth path through lon/lat points, with an arrowhead at the end."""
    pts = [m.xy(*p) for p in points]
    d = f"M{pts[0][0]:.2f},{pts[0][1]:.2f}"
    for (x0, y0), (x1, y1) in zip(pts, pts[1:]):
        mx, my = (x0 + x1) / 2, (y0 + y1) / 2
        d += f" Q{x0:.2f},{y0:.2f} {mx:.2f},{my:.2f}"
    d += f" L{pts[-1][0]:.2f},{pts[-1][1]:.2f}"
    dash_attr = f' stroke-dasharray="{dash}"' if dash else ""
    m.raw(f'<path d="{d}" fill="none" stroke="{color}" stroke-width="{width}" '
          f'stroke-linecap="round"{dash_attr}/>')
    (xa, ya), (xb, yb) = pts[-2], pts[-1]
    length = math.hypot(xb - xa, yb - ya) or 1
    ux, uy = (xb - xa) / length, (yb - ya) / length
    head = 2.2
    hx, hy = xb - ux * head, yb - uy * head
    px, py = -uy * head * 0.55, ux * head * 0.55
    m.raw(f'<path d="M{xb:.2f},{yb:.2f} L{hx + px:.2f},{hy + py:.2f} L{hx - px:.2f},{hy - py:.2f} Z" fill="{color}"/>')


def map_migrations():
    """Early human journeys from Africa to Island Southeast Asia."""
    m = Map(22, -32, 160, 46)
    m.shape(world_land(), LAND, stroke=MUTED, width=0.15)
    archaic, sapiens, austro = GOLD, ACCENT, WATER_INK
    curved_arrow(m, [(36, 2), (40, 25), (44.5, 41.3)], archaic, 1.0)
    curved_arrow(m, [(44.5, 41.3), (75, 30), (100, 12), (110.8, -7.4)], archaic, 1.0)
    curved_arrow(m, [(110.8, -7.4), (121.0, -8.6)], archaic, 0.8, "1.4,0.8")
    curved_arrow(m, [(112, -5), (119.9, -4.4)], archaic, 0.8, "1.4,0.8")
    curved_arrow(m, [(112, 2), (117, 12), (121.5, 17.4)], archaic, 0.8, "1.4,0.8")
    curved_arrow(m, [(39, 9), (45, 13), (60, 22), (78, 12), (98, 8), (106, 0)], sapiens, 1.1)
    curved_arrow(m, [(106, 0), (118, -3), (133, -14)], sapiens, 1.1)
    curved_arrow(m, [(106, 0), (114, 6), (117.9, 9.3)], sapiens, 1.0)
    curved_arrow(m, [(121.0, 23.3), (122, 20.4), (121.7, 17.9)], austro, 1.0, "0.5,0.5")
    sites = [
        ("Dmanisi\n1.8 million years", 44.5, 41.3, "start", 2, 1, archaic),
        ("Java\n1.5–1.3 million", 110.8, -7.4, "end", -2.5, 4.5, archaic),
        ("Flores\n700,000", 121.0, -8.6, "start", 1.5, 4.8, archaic),
        ("Sulawesi\n≥118,000", 119.9, -4.4, "start", 2.5, -1.2, archaic),
        ("Luzon\n709,000", 121.5, 17.4, "start", 2.2, 0.5, archaic),
        ("Tabon\n~47,000", 117.9, 9.3, "end", -2.2, -1.6, sapiens),
        ("Sahul\n65,000–50,000", 133, -14, "start", 2, 3.5, sapiens),
        ("Taiwan → Luzon\n4,500–3,500 years", 121.0, 23.3, "start", 2.5, -2.5, austro),
    ]
    for label, lon, lat, anchor, dx, dy, color in sites:
        m.dot(lon, lat, r=0.7, fill=color)
        m.text(lon, lat, label, size=2.4, anchor=anchor, dx=dx, dy=dy, color=INK)
    m.text(31, -6, "AFRICA", size=3.0, color=MUTED, spacing=0.4, halo=False)
    m.text(80, 40, "ASIA", size=3.0, color=MUTED, spacing=0.4, halo=False)
    m.text(135, -26, "AUSTRALIA", size=2.8, color=MUTED, spacing=0.4, halo=False)
    x0, y0 = m.xy(124.0, 30)
    # Wallace Line, schematic
    wl = [(116.8, -9), (116.2, -6), (118.5, -1), (119.5, 3), (120.5, 6.5)]
    m.raw('<path d="M' + " L".join(f"{m.xy(*p)[0]:.2f},{m.xy(*p)[1]:.2f}" for p in wl) +
          f'" fill="none" stroke="{INK}" stroke-width="0.3" stroke-dasharray="0.8,0.6"/>')
    m.text(113.5, -13.5, "Wallace Line", size=2.2, italic=True, anchor="middle", color=INK)
    lx, ly = m.xy(24, 0)[0], m.height - 15
    m.raw(f'<g font-family="{FONT}" font-size="2.4" fill="{INK}">'
          f'<rect x="{lx}" y="{ly}" width="50" height="13" fill="{PAPER}" stroke="{RULE}" stroke-width="0.3"/>'
          f'<line x1="{lx + 2}" y1="{ly + 3.3}" x2="{lx + 8}" y2="{ly + 3.3}" stroke="{archaic}" stroke-width="1"/>'
          f'<text x="{lx + 10}" y="{ly + 4}">Early humans (Homo erectus and others)</text>'
          f'<line x1="{lx + 2}" y1="{ly + 7.3}" x2="{lx + 8}" y2="{ly + 7.3}" stroke="{sapiens}" stroke-width="1"/>'
          f'<text x="{lx + 10}" y="{ly + 8}">Homo sapiens</text>'
          f'<line x1="{lx + 2}" y1="{ly + 11.1}" x2="{lx + 8}" y2="{ly + 11.1}" stroke="{austro}" stroke-width="1" stroke-dasharray="0.5,0.5"/>'
          f'<text x="{lx + 10}" y="{ly + 11.8}">Austronesian seafarers</text></g>')
    m.save("migrations")


def map_prehistoric_sites(munis):
    """Prehistoric sites of the Philippines, and around Laguna de Bay."""
    m = Map(116.2, 4.6, 128.4, 21.4, height=150)
    for props, geom in topo.load(DATA / "islandgroups.topo.0.001.json"):
        m.shape(geom.buffer(0), LAND, stroke=MUTED, width=0.2)
    sites = [  # name, point, label position, anchor, note
        ("Batanes", (121.97, 20.45), (122.4, 20.45), "start", "route from Taiwan"),
        ("Nagsabaran", (121.66, 18.20), (122.9, 19.2), "start", "red-slipped pottery, ~4,000 years"),
        ("Callao Cave", (121.83, 17.71), (122.9, 18.2), "start", "Homo luzonensis"),
        ("Andarayan", (121.69, 17.62), (122.9, 17.2), "start", "rice grain, ~3,400 years"),
        ("Rizal, Kalinga", (121.55, 17.45), (122.9, 16.2), "start", "butchered rhinoceros, 709,000 years"),
        ("Ille Cave", (119.41, 11.20), (120.3, 10.9), "start", "cremation, ~9,000 years"),
        ("Tabon Caves", (117.98, 9.28), (118.5, 8.3), "start", "Homo sapiens ~47,000 years;\nManunggul Jar"),
        ("Butuan", (125.54, 8.95), (124.9, 9.9), "end", "balangay boats, from ~320 CE"),
    ]
    for name, (lon, lat), (tlon, tlat), anchor, note in sites:
        if (tlon, tlat) != (lon, lat):
            m.leader((lon, lat), (tlon, tlat))
        m.dot(lon, lat, r=0.85, fill=ACCENT)
        dx = 0.8 if anchor == "start" else -0.8
        m.text(tlon, tlat, name, size=2.7, weight=600, anchor=anchor, dx=dx, dy=0.2)
        m.text(tlon, tlat, note, size=2.2, italic=True, color=MUTED, anchor=anchor, dx=dx, dy=3.0)
    # Laguna de Bay area box
    bx1, by1 = m.xy(120.9, 14.75)
    bx2, by2 = m.xy(121.6, 14.05)
    m.raw(f'<rect x="{bx1:.2f}" y="{by1:.2f}" width="{bx2 - bx1:.2f}" height="{by2 - by1:.2f}" fill="none" '
          f'stroke="{ACCENT}" stroke-width="0.6"/>')
    m.text(121.65, 14.4, "Laguna de Bay\n(see detail)", size=2.4, anchor="start", italic=True, color=ACCENT)
    # Inset: the lake
    iw = 37.0
    sub = Map(120.92, 14.12, 121.62, 14.66, width=iw)
    for name, info in munis.items():
        sub.shape(info["geom"], LAND, stroke=RULE, width=0.15)
    sub.shape(munis["Biñan"]["geom"], LAND_ALT, stroke=ACCENT, width=0.4)
    lake_sites = [
        ("Biñan", 121.08, 14.34, "end", 700),
        ("Santa Ana (Manila)", 121.01, 14.58, "start", 400),
        ("Angono–Binangonan petroglyphs", 121.18, 14.545, "start", 400),
        ("Baras, Tanay, Pililla\nstone tools", 121.28, 14.51, "start", 400),
        ("Bay (Puypuy)\nburial jars", 121.28, 14.18, "end", 400),
        ("Pila", 121.36, 14.23, "start", 400),
        ("Lumban\n(copper plate)", 121.46, 14.30, "end", 400),
    ]
    for name, lon, lat, anchor, weight in lake_sites:
        sub.dot(lon, lat, r=0.7, fill=ACCENT if name != "Biñan" else PAPER, stroke=INK if name == "Biñan" else PAPER)
        sub.text(lon, lat, name, size=2.2, anchor=anchor, dx=1.6 if anchor == "start" else -1.6,
                 dy=-1.2 if "petroglyphs" in name else (3.2 if "Baras" in name else 0.8),
                 weight=weight, color=ACCENT if name == "Biñan" else INK)
    sub.text(121.22, 14.40, "Laguna\nde Bay", size=2.6, italic=True, color=WATER_INK, halo=False)
    ox, oy = 3, m.xy(0, 15.9)[1]
    m.raw(f'<defs><clipPath id="inset-clip"><rect width="{iw}" height="{sub.height:.2f}"/></clipPath></defs>'
          f'<g transform="translate({ox:.2f},{oy:.2f})" clip-path="url(#inset-clip)">'
          f'<rect width="{iw}" height="{sub.height:.2f}" fill="{WATER}"/>'
          + "".join(sub.items) + f'<rect width="{iw}" height="{sub.height:.2f}" fill="none" stroke="{ACCENT}" '
          f'stroke-width="0.6"/></g>')
    m.scale_bar(200, x=4, y=m.height - 5)
    m.north_arrow(x=7, y=8)
    m.save("prehistoric-sites")


def map_polities(munis):
    """Laguna de Bay and its communities around 1570."""
    m = Map(120.90, 14.05, 121.60, 14.68)
    for name, info in munis.items():
        m.shape(info["geom"], LAND, stroke=RULE, width=0.15)
    places = [  # name, lon, lat, anchor, dx, dy, note
        ("Tondo", 120.968, 14.612, "start", 1.6, -1.6, "Lakandula"),
        ("Maynila", 120.975, 14.590, "start", 1.6, 2.4, "Rajah Sulayman"),
        ("Cainta", 121.115, 14.578, "start", 1.6, -0.6, "fortified; taken by Salcedo, Aug. 1571"),
        ("Taytay", 121.132, 14.560, "start", 1.6, 2.6, ""),
        ("Tunasan", 121.045, 14.372, "end", -1.6, 0.8, ""),
        ("Tabuco", 121.125, 14.272, "start", 1.6, 0.8, "encomienda: 1,000 tributes (1591)"),
        ("Bay", 121.284, 14.183, "start", 1.6, 2.6, "gave its name to the lake"),
        ("Pila", 121.364, 14.234, "start", 1.6, 0.8, "graves of the 12th–14th c."),
        ("Lumban", 121.460, 14.297, "end", -1.6, -1.6, "copper plate of 900 found here"),
        ("Nagcarlán", 121.417, 14.137, "start", 1.6, 0.8, "Plasencia writes, 1589"),
    ]
    for name, lon, lat, anchor, dx, dy, note in places:
        m.dot(lon, lat, r=0.75, fill=INK)
        m.text(lon, lat, name, size=2.9, weight=600, anchor=anchor, dx=dx, dy=dy)
        if note:
            m.text(lon, lat, note, size=2.2, italic=True, color=MUTED, anchor=anchor, dx=dx, dy=dy + 2.8)
    poblacion = centre(barangays()["Poblacion"])
    m.dot(*poblacion, r=0.95, fill=PAPER, stroke=ACCENT)
    m.text(*poblacion, "site of Biñan", size=3.0, weight=700, color=ACCENT, anchor="start", dx=1.8, dy=-0.8)
    m.text(*poblacion, "no record before 1571", size=2.2, italic=True, color=ACCENT, anchor="start", dx=1.8, dy=2.0)
    m.text(121.25, 14.40, "Laguna de Bay", size=4.2, italic=True, color=WATER_INK, halo=False, spacing=0.2)
    m.text(120.93, 14.52, "Manila\nBay", size=3.0, italic=True, color=WATER_INK, halo=False)
    m.text(121.03, 14.63, "Pasig River", size=2.4, italic=True, color=WATER_INK, anchor="start")
    m.scale_bar(10, x=m.width - 30)
    m.north_arrow(x=7, y=m.height - 16)
    m.save("polities")


if __name__ == "__main__":
    munis = municipalities()
    bgys = barangays()
    map_location(munis, luzon())
    map_territory(munis)
    map_liberation(munis, bgys)
    map_barangays(bgys, munis)
    map_tectonics()
    map_westbay(munis)
    chart_climate()
    chart_population()
    map_transport(munis, bgys)
    map_economy(bgys, munis)
    map_land(bgys, munis)
    map_migrations()
    map_prehistoric_sites(munis)
    map_polities(munis)
