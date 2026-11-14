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
              f'stroke-width="{0.9 if is_wvf else 0.3}" opacity="{1 if is_wvf else 0.8}"/>')
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
          f'<line x1="{lx + 2}" y1="{ly + 13}" x2="{lx + 8}" y2="{ly + 13}" stroke="{ACCENT}" stroke-width="1"/>'
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
    curved_arrow(m, [(121.0, 23.3), (122, 20.4), (121.7, 17.9)], austro, 1.0)
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
          f'<line x1="{lx + 2}" y1="{ly + 11.1}" x2="{lx + 8}" y2="{ly + 11.1}" stroke="{austro}" stroke-width="1"/>'
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
    map_migrations()
    map_prehistoric_sites(munis)
