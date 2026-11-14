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

# Maps (drawn by book/maps/make_maps.py), placed after the first paragraph
# of a chapter of the history, numbered from 1.
MAPS = {
    1: ("location",
        "Biñan on the western shore of Laguna de Bay, south of Manila. Before the "
        "Spaniards came, the area belonged to the Tagalog settlement of Tabuco, "
        "today Cabuyao.",
        "Ang Biñán sa kanlurang pampang ng Lawa ng Laguna."),
    2: ("territory",
        "How Biñan’s territory took shape: separated from Tabuco in 1688, it lost its "
        "barrio of Santa Rosa in 1791–92. Under American rule, San Pedro (1903–1907) "
        "and Muntinlupa (1903–1905) were briefly joined to it. Present-day boundaries.",
        "Ang paghubog ng teritoryo ng Biñán, 1688–1907."),
    4: ("liberation",
        "The liberation of Biñan, 3 February 1945: the Hunters ROTC guerrillas "
        "converged on the town centre from Carmona, San Pedro and Santa Rosa. "
        "Arrows are schematic.",
        "Ang paglaya ng Biñán, Pebrero 3, 1945."),
    5: ("barangays",
        "Biñan today and its 24 barangays. The Poblacion keeps the Spanish-era heart "
        "of the town; the Laguna Technopark (Malamig) and the Laguna International "
        "Industrial Park (Mampalasan) mark its industrial growth since 1990.",
        "Ang 24 na barangay ng Biñán ngayon."),
}

MAP_TITLES = {
    "location": "Biñan and Laguna de Bay",
    "territory": "The territory of Biñan, 1688–1907",
    "liberation": "The liberation of Biñan, 1945",
    "barangays": "The barangays of Biñan today",
}

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
    map_pending = False
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
            map_pending = chapter in MAPS
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
        if map_pending and not tagalog:
            name, caption, tagalog_caption = MAPS[chapter]
            out.append(
                f'<figure class="map" id="map-{name}"><img src="maps/{name}.svg" alt="">'
                f'<figcaption><span class="map-no">Map {list(MAPS).index(chapter) + 1}.</span> '
                f'{caption} <em class="tl">{tagalog_caption}</em></figcaption></figure>'
            )
            map_pending = False
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
    maps = [(f"map-{name}", MAP_TITLES[name])
            for name in re.findall(r'<figure class="map" id="map-([^"]+)"', body)]
    map_list = ""
    if maps:
        map_items = "\n".join(
            f'<li><a class="toc-map" href="#{anchor}">{title}</a></li>'
            for anchor, title in maps)
        map_list = f'<p class="toc-maps-title">Maps · Mga Mapa</p><ol class="toc-maps">{map_items}</ol>'
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
    return f"""
<section class="colophon">
  <p class="colophon-title">{doc['title']}</p>
  <p>{doc['subtitle']}</p>
  <p>Part of <em>The Book of Biñan</em>, an open project: text and sources at
     github.com/smichea/binan.</p>
  <p class="colophon-head">Photographs</p>
  {credits}
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
