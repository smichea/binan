"""Impose the English interior for a home-printed, hand-sewn proof copy.

Each A4 sheet, printed on both sides and folded in half, carries four A5 pages.
Sheets are nested into sections (signatures) of SHEETS sheets each, to be sewn
one by one and then bound together. The 6 x 9 in pages are scaled down to fit
A5 (about 92 %), so the page breaks stay those of the printed edition.

    uv run --with pypdf python3 book/make_booklet.py [--sheets 4] [--colour]

Writes book/word/print/proof_a4_booklet.pdf: one portrait A4 page per sheet
side (the spread is turned a quarter turn), front then back, signature after
signature. For one signature n, print its fronts (pages 8n-7, 8n-5, 8n-3,
8n-1), turn the stack over into the tray, then print its backs (8n-6 ... 8n).
"""
import argparse
from pathlib import Path

from pypdf import PageObject, PdfReader, PdfWriter, Transformation

HERE = Path(__file__).resolve().parent
PRINT = HERE / "word" / "print"
A4_W, A4_H = 841.89, 595.28  # landscape, points
HALF = A4_W / 2

parser = argparse.ArgumentParser()
parser.add_argument("--sheets", type=int, default=4, help="sheets per signature (4 = 16 pages)")
parser.add_argument("--colour", action="store_true", help="use the colour interior instead of black and white")
parser.add_argument("--rotate-backs", action="store_true", help="turn the backs 180° if the test sheet came out upside down")
args = parser.parse_args()

src = PRINT / ("history_english_interior.pdf" if args.colour else "history_english_interior_bw.pdf")
reader = PdfReader(src)
pages = list(reader.pages)
pw, ph = float(pages[0].mediabox.width), float(pages[0].mediabox.height)
scale = min(HALF / pw, A4_H / ph)
dx, dy = (HALF - pw * scale) / 2, (A4_H - ph * scale) / 2

per_sig = 4 * args.sheets
n_sigs = -(-len(pages) // per_sig)
pages += [None] * (n_sigs * per_sig - len(pages))  # blank pages at the end


def side(left, right, rotate=False):
    sheet = PageObject.create_blank_page(width=A4_W, height=A4_H)
    for slot, page in enumerate((left, right)):
        if page is None:
            continue
        t = Transformation().scale(scale).translate(slot * HALF + dx, dy)
        if rotate:
            t = t.rotate(180).translate(A4_W, A4_H)
        sheet.merge_transformed_page(page, t)
    # Hand the printer a portrait A4 page with the spread turned a quarter
    # turn, so that it does not shrink the landscape sheet to fit.
    portrait = PageObject.create_blank_page(width=A4_H, height=A4_W)
    portrait.merge_transformed_page(sheet, Transformation().rotate(90).translate(A4_H, 0))
    return portrait


out = PdfWriter()
for s in range(n_sigs):
    sig = pages[s * per_sig:(s + 1) * per_sig]
    n = len(sig)
    for k in range(args.sheets):
        out.add_page(side(sig[n - 1 - 2 * k], sig[2 * k]))  # front: outer pages
        out.add_page(side(sig[2 * k + 1], sig[n - 2 - 2 * k], args.rotate_backs))  # back
out.compress_identical_objects()
dest = PRINT / "proof_a4_booklet.pdf"
with open(dest, "wb") as f:
    out.write(f)

total = n_sigs * args.sheets
print(f"{dest}: {len(reader.pages)} pages -> {n_sigs} signatures of {args.sheets} sheets "
      f"({total} A4 sheets, {2 * total} sides), scale {scale:.1%}")
