"""Build the print-ready PDFs of the Biñan book.

Markdown (book/raw_files) -> HTML (pandoc) -> PDF (WeasyPrint), with a cover
page, a table of contents with page numbers, source links turned into
footnotes, and "Did You Know?" boxes.

Requirements: pandoc, and `pip install weasyprint`.
Usage (from anywhere):  python3 book/build_pdf.py
"""

import html
import re
import subprocess
import sys
from pathlib import Path
from urllib.parse import unquote, urlsplit

from weasyprint import HTML

BOOK = Path(__file__).resolve().parent
LAYOUT = BOOK / "layout"

DOCUMENTS = [
    {
        "source": BOOK / "raw_files" / "history.md",
        "output": BOOK / "word" / "history.pdf",
        "kicker": "The Book of Biñan",
        "title": "History of Biñan City",
        "subtitle": "Kasaysayan ng Lungsod ng Biñan",
        "tagline": "From Tabuco to the City of Biñan, 1571 – today",
        "languages": "English · Tagalog",
    },
    {
        "source": BOOK / "raw_files" / "history_tagalog_explanation.md",
        "output": BOOK / "word" / "history_tagalog_explanation.pdf",
        "kicker": "The Book of Biñan · Companion",
        "title": "Reading the Tagalog Text",
        "subtitle": "Vocabulary, grammar and verbs of the History of Biñan City",
        "tagline": "For English-only readers",
        "languages": "Tagalog → English",
    },
]

# Photos from Wikimedia Commons used in the layout, with their credits.
IMAGES = {
    "cover": {
        "file": "images/san-isidro-church.jpg",
        "credit": "San Isidro Labrador Parish Church, Biñan (cover). Photo: "
                  "patrickroque01, CC BY-SA 4.0, via Wikimedia Commons.",
    },
    "File:Alberto Mansion Biñan City, Laguna.jpg": {
        "file": "images/alberto-mansion.jpg",
        "credit": "Alberto Mansion, Biñan. Photo: Bengoeswhere, "
                  "CC BY-SA 3.0, via Wikimedia Commons.",
    },
}

# Maps drawn for the book (book/maps/make_maps.py) and archival documents,
# placed in a chapter of the history after its n-th English paragraph.
FIGURES = [
    dict(chapter=1, after=1, kind="Map", id="location", src="maps/location.svg",
         title="Biñan and Laguna de Bay",
         caption="Biñan on the western shore of Laguna de Bay, south of Manila. Before "
                 "the Spaniards came, the area belonged to the Tagalog settlement of "
                 "Tabuco, today Cabuyao.",
         tagalog="Ang Biñán sa kanlurang pampang ng Lawa ng Laguna."),
    dict(chapter=1, after=1, kind="Photo", id="copperplate",
         src="images/archives/copperplate-c900.jpg",
         title="The Laguna Copperplate Inscription, c. 900",
         caption="The Laguna Copperplate Inscription (c. 900 CE), the oldest known written "
                 "document of the Philippines, found in 1989 near Lumban on Laguna de Bay. "
                 "It shows that the lake shore was home to organized, literate communities "
                 "centuries before the Spaniards arrived.",
         tagalog="Ang Inskripsiyon sa Binatbat na Tanso ng Laguna, mga 900.",
         credit="Photo: Gary Todd, CC0, via Wikimedia Commons. National Museum of the Philippines."),
    dict(chapter=2, after=1, kind="Map", id="territory", src="maps/territory.svg",
         title="The territory of Biñan, 1688–1907",
         caption="How Biñan’s territory took shape: separated from Tabuco in 1688, it lost "
                 "its barrio of Santa Rosa in 1791–92. Under American rule, San Pedro "
                 "(1903–1907) and Muntinlupa (1903–1905) were briefly joined to it. "
                 "Present-day boundaries.",
         tagalog="Ang paghubog ng teritoryo ng Biñán, 1688–1907."),
    dict(chapter=2, after=2, kind="Map", id="murillo", src="images/archives/murillo-velarde-1734.jpg",
         title="“Viñan” on the Murillo Velarde map, 1734",
         caption="The south-western shore of Laguna de Bay on Pedro Murillo Velarde’s 1734 map of the Philippines, the "
                 "first scientific map of the islands. “Viñan” appears at the left edge, on the shore "
                 "near Tunasan and Calamba.",
         tagalog="Ang “Viñan” sa mapa ni Murillo Velarde, 1734.",
         credit="Pedro Murillo Velarde, Carta hydrographica y chorographica de las Yslas "
                "Filipinas (Manila, 1734). Library of Congress, Geography and Map Division. "
                "Public domain."),
    dict(chapter=2, after=3, kind="Map", id="manuscript", src="images/archives/manuscript-laguna-18th-c.jpg",
         title="“Biñang” on an 18th-century Spanish map",
         caption="Detail of an 18th-century Spanish manuscript map of the surroundings of "
                 "Laguna de Bay: the town of “Biñang” with its church (red square), the river "
                 "mouth (“Bª de Biñang”), Malabag (today Malaban), Sinalsan, Landayan, San "
                 "Pedro Tunasan and Santa Rosa.",
         tagalog="Ang “Biñang” sa isang mapang Espanyol noong ika-18 siglo.",
         credit="Fuente: Archivo Cartográfico de Estudios Geográficos del Centro Geográfico "
                "del Ejército / Biblioteca Virtual de Defensa. CC BY 4.0."),
    dict(chapter=3, after=1, kind="Map", id="atlas", src="images/archives/atlas-filipinas-1899.jpg",
         title="Biñan in the Atlas de Filipinas, 1899–1900",
         caption="“Biñang” in the Atlas de Filipinas of the Manila Observatory (Map No. 9, "
                 "U.S. Coast and Geodetic Survey edition, 1900), on the road from Manila to "
                 "Calamba, as the Americans found it.",
         tagalog="Ang Biñán sa Atlas de Filipinas, 1899–1900.",
         credit="Observatorio de Manila / U.S. Coast and Geodetic Survey, Atlas of the "
                "Philippine Islands (Washington, 1900), Map No. 9. Internet Archive. Public domain."),
    dict(chapter=4, after=1, kind="Map", id="liberation", src="maps/liberation.svg",
         title="The liberation of Biñan, 1945",
         caption="The liberation of Biñan, 3 February 1945: the Hunters ROTC guerrillas "
                 "converged on the town centre from Carmona, San Pedro and Santa Rosa. "
                 "Arrows are schematic.",
         tagalog="Ang paglaya ng Biñán, Pebrero 3, 1945."),
    dict(chapter=5, after=1, kind="Photo", id="bridge", src="images/archives/binan-bridge-1951.jpg",
         title="The new Biñan bridge, c. 1951",
         caption="“Biñan Bridge completed”, about 1951: one of the bridges of the national "
                 "road rebuilt after the war under the Philippine–American public roads "
                 "rehabilitation programme.",
         tagalog="Ang bagong tulay ng Biñán, mga 1951.",
         credit="U.S. National Archives, Records of the Bureau of Public Roads (RG 30), "
                "30-N-51-4622. Public domain."),
    dict(chapter=5, after=1, kind="Map", id="ams1956", src="images/archives/ams-san-pedro-1956.jpg",
         title="The town of Biñan on a 1956 survey map",
         caption="The town of Biñan on the 1956 U.S. Army Map Service sheet “San Pedro” "
                 "(1:50,000): the street grid of the Poblacion and the neighbouring barrios "
                 "of Canlalay, Malaban, De La Paz and San Antonio, between the national road "
                 "and the lake.",
         tagalog="Ang bayan ng Biñán sa mapa ng U.S. Army, 1956.",
         credit="U.S. Army Map Service, San Pedro (sheet 3263 III), 1:50,000, 1956. "
                "Courtesy of the University of Texas Libraries. Public domain."),
    dict(chapter=5, after=2, kind="Map", id="barangays", src="maps/barangays.svg",
         title="The barangays of Biñan today",
         caption="Biñan today and its 24 barangays. The Poblacion keeps the Spanish-era "
                 "heart of the town; the Laguna Technopark (Malamig) and the Laguna "
                 "International Industrial Park (Mampalasan) mark its industrial growth "
                 "since 1990.",
         tagalog="Ang 24 na barangay ng Biñán ngayon."),
]


# Documents still under copyright are marked rights="pending" in FIGURES: they
# appear in the draft with a notice, and are left out by `build_pdf.py --final`.
FINAL = "--final" in sys.argv


def active_figures():
    return [f for f in FIGURES if not (FINAL and f.get("rights") == "pending")]


def figure_html(fig, number):
    credit = f' <span class="credit">{fig["credit"]}</span>' if fig.get("credit") else ""
    if fig.get("rights") == "pending":
        credit += ' <span class="rights-pending">Reproduction rights being requested.</span>'

    css = "map" if fig["src"].endswith(".svg") else "archive"
    return (
        f'<figure class="{css}" id="fig-{fig["id"]}"><img src="{fig["src"]}" alt="">'
        f'<figcaption><span class="map-no">{fig["kind"]} {number}.</span> '
        f'{fig["caption"]} <em class="tl">{fig["tagalog"]}</em>{credit}</figcaption></figure>'
    )


# A photo caption in the source: "([File:...jpg - Wikimedia Commons](url)) *caption*"
FIGURE = re.compile(r"^\(\[(File:[^\]]+?\.jpg)[^\]]*\]\([^)]+\)\)\s*\*(.+)\*$", re.S)

# Words that are very frequent in Tagalog and rare in English: used to tell
# the Tagalog paragraphs of each chapter apart from the English ones.
TAGALOG_MARKERS = {
    "ang", "ng", "mga", "sa", "na", "ay", "at", "ito", "noong", "nang",
    "kanilang", "isang", "bilang", "dahil", "mula", "pa", "rin", "din",
}


# --------------------------------------------------------------------------
# Footnotes
# --------------------------------------------------------------------------

def short_url(url):
    """Readable form of a URL for a footnote: no scheme, no #:~:text= part."""
    parts = urlsplit(url)
    host = parts.netloc.removeprefix("www.")
    path = unquote(parts.path).rstrip("/")
    query = f"?{parts.query}" if parts.query else ""
    return f"{host}{path}{query}"


def clean_url(url):
    return url.split("#:~:text=")[0]


def footnote(label, url):
    url = clean_url(url)
    label = html.escape(label.strip())
    shown = html.escape(short_url(url))
    return (
        f'<span class="fn"><span class="fn-label">{label}</span>, '
        f'<a href="{html.escape(url)}">{shown}</a></span>'
    )


# A citation in parentheses: "text ([Label](url))"  -> footnote after text.
CITATION = re.compile(r"\s*\(\[([^\]]+)\]\((https?://[^)\s]+)\)\)")
# A plain link inside the sentence: "[words](url)" -> words + footnote.
INLINE_LINK = re.compile(r"\[([^\]]+)\]\((https?://[^)\s]+)\)")


def links_to_footnotes(text):
    text = CITATION.sub(lambda m: footnote(m.group(1), m.group(2)), text)
    return INLINE_LINK.sub(
        lambda m: m.group(1) + footnote(m.group(1), m.group(2)), text
    )


# --------------------------------------------------------------------------
# Markdown pre-processing
# --------------------------------------------------------------------------

def is_tagalog(paragraph):
    words = re.findall(r"[a-zñ’']+", paragraph.lower())
    if len(words) < 12:
        return False
    hits = sum(1 for w in words if w in TAGALOG_MARKERS)
    return hits / len(words) > 0.12


def split_heading(line):
    """'## English / Tagalog' -> heading with the Tagalog part as subtitle."""
    level, text = line.split(" ", 1)
    if " /" in text:
        english, tagalog = (part.strip() for part in text.split("/", 1))
        return (
            f'{level} {english} {{.chapter data-tl="{html.escape(tagalog)}"}}'
        )
    return line


def prepare_history(markdown):
    """Mark up the bilingual history text for the book layout."""
    blocks = markdown.split("\n\n")
    out = []
    in_tagalog = False
    chapter = 0
    english_paragraphs = 0
    numbers = {}
    for block in blocks:
        block = block.strip()
        if not block or block.startswith("# "):
            continue  # the H1 title is on the cover
        if block.startswith("## "):
            if in_tagalog:
                out.append(":::")
                in_tagalog = False
            out.append(split_heading(block))
            chapter += 1
            english_paragraphs = 0
            continue
        if block.startswith("**Did You Know?**"):
            if in_tagalog:
                out.append(":::")
                in_tagalog = False
            body = block.removeprefix("**Did You Know?**").strip()
            out.append(
                '::: {.didyouknow}\n<p class="dyk-title">Did you know?</p>\n\n'
                f"{links_to_footnotes(body)}\n:::"
            )
            continue
        figure = FIGURE.match(block)
        if figure and figure.group(1) in IMAGES:
            image = IMAGES[figure.group(1)]
            caption = links_to_footnotes(figure.group(2))
            out.append(
                f'<figure><img src="{image["file"]}" alt="">'
                f"<figcaption>\n\n{caption}\n\n</figcaption></figure>"
            )
            continue
        tagalog = is_tagalog(block)
        if tagalog and not in_tagalog:
            out.append('::: {.tagalog}\n<p class="lang-label">Sa Tagalog</p>')
            in_tagalog = True
        elif not tagalog and in_tagalog:
            out.append(":::")
            in_tagalog = False
        out.append(links_to_footnotes(block))
        if not tagalog:
            english_paragraphs += 1
            for fig in active_figures():
                if (fig["chapter"], fig["after"]) == (chapter, english_paragraphs):
                    numbers[fig["kind"]] = numbers.get(fig["kind"], 0) + 1
                    out.append(figure_html(fig, numbers[fig["kind"]]))
    if in_tagalog:
        out.append(":::")
    return "\n\n".join(out)


def prepare_explanation(markdown):
    """Mark up the vocabulary companion for the book layout."""
    out = []
    for line in markdown.split("\n"):
        stripped = line.strip()
        if stripped.startswith("# ") or stripped == "---":
            continue  # title is on the cover; the layout separates sections
        if stripped.lower().startswith("explanation of tagalog text"):
            continue  # already the cover tagline
        if stripped.startswith("## "):
            out.extend(["", split_heading(stripped), ""])
            continue
        out.append(links_to_footnotes(line))
    return "\n".join(out)


def to_html(markdown, source_format="markdown-implicit_figures+raw_html"):
    result = subprocess.run(
        ["pandoc", f"--from={source_format}",
         "--to=html5", "--wrap=none", "--shift-heading-level-by=-1"],
        input=markdown, capture_output=True, text=True, check=True,
    )
    return result.stdout


# --------------------------------------------------------------------------
# Book assembly
# --------------------------------------------------------------------------

def table_of_contents(body):
    entries = re.findall(
        r'<h1 class="chapter" data-tl="([^"]*)"\s+id="([^"]+)">(.*?)</h1>',
        body, flags=re.S,
    )
    items = "\n".join(
        f'<li><a class="toc-en" href="#{anchor}">{" ".join(title.split())}</a>'
        f'<span class="toc-tl">{tagalog}</span></li>'
        for tagalog, anchor, title in entries
    )
    present = set(re.findall(r'<figure class="[^"]+" id="fig-([^"]+)"', body))
    figures = [f for f in active_figures() if f["id"] in present]
    map_list = ""
    if figures:
        counts, rows = {}, []
        for fig in figures:
            counts[fig["kind"]] = counts.get(fig["kind"], 0) + 1
            rows.append(f'<li><a class="toc-map" href="#fig-{fig["id"]}">'
                        f'<span class="toc-kind">{fig["kind"]} {counts[fig["kind"]]}.</span> '
                        f'{fig["title"]}</a></li>')
        map_list = ('<p class="toc-maps-title">Maps and documents · Mga mapa at dokumento</p>'
                    f'<ol class="toc-maps">{"".join(rows)}</ol>')
    return (
        '<section class="toc"><h1 class="toc-title">Contents</h1>'
        f'<p class="toc-sub">Nilalaman</p><ol>{items}</ol>{map_list}</section>'
    )


def cover(doc):
    return f"""
<section class="cover">
  <div class="cover-frame">
    <p class="cover-kicker">{doc['kicker']}</p>
    <div class="cover-rule"></div>
    <h1 class="cover-title">{doc['title']}</h1>
    <p class="cover-subtitle">{doc['subtitle']}</p>
    <div class="cover-photo" style="background-image: url({IMAGES['cover']['file']})"></div>
    <p class="cover-tagline">{doc['tagline']}</p>
    <p class="cover-languages">{doc['languages']}</p>
  </div>
</section>"""


def colophon(doc):
    credits = "".join(f"<p>{image['credit']}</p>" for image in IMAGES.values())
    archive_credits = ""
    if doc["source"].name == "history.md":
        archive_credits = '<p class="colophon-head">Archival maps and photographs</p>' + "".join(
            f"<p>{fig['title']}: {fig['credit']}</p>" for fig in active_figures() if fig.get("credit"))
    return f"""
<section class="colophon">
  <p class="colophon-title">{doc['title']}</p>
  <p>{doc['subtitle']}</p>
  <p>Part of <em>The Book of Biñan</em>, an open project: text and sources at
     github.com/smichea/binan.</p>
  <p class="colophon-head">Photographs</p>
  {credits}
  {archive_credits}
  <p class="colophon-head">Maps</p>
  <p>Drawn for this book from the Philippine Standard Geographic Code boundaries
     (PSA, via OCHA and github.com/faeldon/philippines-json-maps, MIT licence).</p>
  <p class="colophon-head">Typography</p>
  <p>Set in EB Garamond and Cormorant Garamond (SIL Open Font License).</p>
  <p class="colophon-head">Sources</p>
  <p>Sources are given as footnotes. Web addresses are shortened for print.</p>
</section>"""


def build(doc):
    markdown = doc["source"].read_text(encoding="utf-8")
    if doc["source"].name == "history.md":
        prepared = prepare_history(markdown)
        body_class = "history"
    else:
        prepared = prepare_explanation(markdown)
        body_class = "companion"
    # The companion is written like GitHub markdown (lists and quotes right
    # after a line of text), which CommonMark reads correctly.
    source_format = "commonmark_x" if body_class == "companion" else (
        "markdown-implicit_figures+raw_html")
    body = to_html(prepared, source_format)
    page = f"""<!doctype html>
<html lang="en"><head><meta charset="utf-8">
<title>{doc['title']}</title>
<link rel="stylesheet" href="book.css">
</head><body class="{body_class}">
{cover(doc)}
{colophon(doc)}
{table_of_contents(body)}
<main>{body}</main>
</body></html>"""
    HTML(string=page, base_url=str(LAYOUT) + "/").write_pdf(doc["output"])
    print(f"wrote {doc['output'].relative_to(BOOK.parent)}")


if __name__ == "__main__":
    for document in DOCUMENTS:
        build(document)
