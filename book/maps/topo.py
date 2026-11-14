"""Minimal TopoJSON reader returning shapely geometries."""
import json
from shapely.geometry import Polygon, MultiPolygon


def load(path):
    topo = json.load(open(path, encoding="utf-8"))
    tr = topo.get("transform")
    arcs = []
    for arc in topo["arcs"]:
        if tr:
            (sx, sy), (tx, ty) = tr["scale"], tr["translate"]
            x = y = 0
            pts = []
            for dx, dy in arc:
                x += dx
                y += dy
                pts.append((x * sx + tx, y * sy + ty))
        else:
            pts = [tuple(p) for p in arc]
        arcs.append(pts)

    def ring(indices):
        pts = []
        for i in indices:
            a = arcs[i] if i >= 0 else arcs[~i][::-1]
            pts.extend(a if not pts else a[1:])
        return pts

    def geom(g):
        if g["type"] == "Polygon":
            rings = [ring(r) for r in g["arcs"]]
            return Polygon(rings[0], rings[1:])
        if g["type"] == "MultiPolygon":
            return MultiPolygon([Polygon(ring(p[0]), [ring(r) for r in p[1:]]) for p in g["arcs"]])
        return None

    obj = next(iter(topo["objects"].values()))
    return [(g.get("properties", {}), geom(g)) for g in obj["geometries"] if geom(g) is not None]
