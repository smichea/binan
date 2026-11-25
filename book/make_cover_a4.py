"""Print the cover of the hand-bound proof on two A4 sheets.

The cover is composed again at the size of the A5 book block made with
make_booklet.py (panels 150 x 212 mm, 1 mm larger than the block all round),
then split: sheet 1 carries the back cover, the spine and a 10 mm glue tab,
sheet 2 the front cover. Cut both along the solid marks, fold the spine along
the dashed marks, then glue the front cover over the tab, its left edge on the
spine fold.

    uv run --with weasyprint python3 book/make_cover_a4.py [pages]

pages: number of pages of the sewn block (default 176, i.e. 11 signatures).
"""
import subprocess
import sys
import tempfile
from pathlib import Path

from weasyprint import HTML

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
import build_pdf  # noqa: E402

PRINT = HERE / "word" / "print"
PANEL, HEIGHT, BLEED, TAB = 150.0, 212.0, build_pdf.BLEED_MM, 10.0  # mm
DPI = 300
SCALE = 1.0
pages = int(sys.argv[1]) if len(sys.argv) > 1 else 176
spine = round(pages / 2 * build_pdf.PAPER_MM + 0.5, 1)
total_h = HEIGHT + 2 * BLEED

pieces = [
    ("back", 0.0, BLEED + PANEL + spine + TAB,
     [BLEED], [BLEED + PANEL, BLEED + PANEL + spine]),
    ("front", BLEED + PANEL + spine, 2 * BLEED + 2 * PANEL + spine,
     [PANEL], []),
]


def mm(px):
    return px / DPI * 25.4


def piece_html(png, x0, x1, cuts, folds):
    w, h = (x1 - x0) * SCALE, total_h * SCALE
    ox, oy = (210 - w) / 2, (297 - h) / 2
    top, bottom = oy + BLEED * SCALE, oy + h - BLEED * SCALE
    marks = []
    for x in cuts + folds:
        style = "solid" if x in cuts else "dashed"
        X = ox + x * SCALE
        marks += [f'<div class="v {style}" style="left:{X:.2f}mm; top:{oy - 9:.2f}mm"></div>',
                  f'<div class="v {style}" style="left:{X:.2f}mm; top:{oy + h + 2:.2f}mm"></div>']
    for Y in (top, bottom):
        marks += [f'<div class="h" style="top:{Y:.2f}mm; left:{ox - 9:.2f}mm"></div>',
                  f'<div class="h" style="top:{Y:.2f}mm; left:{ox + w + 2:.2f}mm"></div>']
    return (f'<section><img src="{png}" style="left:{ox:.2f}mm; top:{oy:.2f}mm; '
            f'width:{w:.2f}mm; height:{h:.2f}mm">{"".join(marks)}</section>')


with tempfile.TemporaryDirectory(dir=PRINT) as tmp:
    SRC = Path(tmp) / "spread.pdf"
    HTML(string=build_pdf.cover_spread(build_pdf.DOCUMENTS[0], spine, PANEL, HEIGHT),
         base_url=str(build_pdf.LAYOUT) + "/").write_pdf(SRC)
    sections = []
    for name, x0, x1, cuts, folds in pieces:
        to_px = DPI / 25.4
        base = Path(tmp) / name
        subprocess.run(["pdftoppm", "-r", str(DPI), "-png", "-singlefile",
                        "-x", str(round(x0 * to_px)), "-y", "0",
                        "-W", str(round((x1 - x0) * to_px)), "-H", str(round(total_h * to_px)),
                        str(SRC), str(base)], check=True)
        sections.append(piece_html(f"{base}.png", x0, x1, [c - x0 for c in cuts], [f - x0 for f in folds]))
    html = f"""<!doctype html><html><head><meta charset="utf-8"><style>
@page {{ size: A4; margin: 0; }}
body {{ margin: 0; }}
section {{ position: relative; width: 210mm; height: 297mm; break-after: page; }}
img {{ position: absolute; }}
.v {{ position: absolute; width: 0; height: 7mm; border-left: 0.3pt solid #000; }}
.v.dashed {{ border-left-style: dashed; }}
.h {{ position: absolute; height: 0; width: 7mm; border-top: 0.3pt solid #000; }}
</style></head><body>{"".join(sections)}</body></html>"""
    dest = PRINT / "proof_cover_a4.pdf"
    HTML(string=html, base_url="/").write_pdf(dest)
print(f"{dest}: panels {PANEL:.0f} x {HEIGHT:.0f} mm, spine {spine} mm")
