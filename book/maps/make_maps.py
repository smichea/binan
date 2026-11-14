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
FONT = "EB Garamond, Georgia, serif"


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
    for name in ("Malamig", "Mampalasan"):
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
          f'<text x="{lx + 8}" y="{ly + 9.8}">Industrial parks (LTI, LIIP)</text>'
          f'<line x1="{lx + 2}" y1="{ly + 14}" x2="{lx + 6}" y2="{ly + 14}" stroke="{INK}" stroke-width="0.5"/>'
          f'<text x="{lx + 8}" y="{ly + 14.8}">City limits</text></g>')
    m.scale_bar(2, x=4, y=m.height - 5)
    m.north_arrow(x=7, y=8)
    m.save("barangays")


if __name__ == "__main__":
    munis = municipalities()
    bgys = barangays()
    map_location(munis, luzon())
    map_territory(munis)
    map_liberation(munis, bgys)
    map_barangays(bgys, munis)
