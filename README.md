# Biñan

The book of Biñan: the history of the Philippines through one town, in an English and a Tagalog edition.

* Read the books (PDF): [The Book of Biñan](book/word/history_english.pdf) (English edition) · [Ang Aklat ng Biñan](book/word/history_tagalog.pdf) (Tagalog edition) · [Reading the Tagalog Text](book/word/history_tagalog_explanation.pdf), the companion for learners of Tagalog
* Sources of the text: [introduction](book/raw_files/introduction.md) · [history, chapters I–XI](book/raw_files/history.md) · [appendices](book/raw_files/appendices.md) · [companion](book/raw_files/history_tagalog_explanation.md)
* Research notes, with every source: [book/research](book/research)
* Working documents (in French): [rights to obtain](book/droits_a_obtenir.md) · [AI illustrations to review](book/illustrations_revue.md) · [interview guide](book/guide_entretiens.md) · [publication checklist](book/publication_checklist.md) · [Tagalog proofreading](book/relecture_tagalog.md)

Rebuild the PDFs with `python3 book/maps/make_maps.py` then `python3 book/build_pdf.py` (needs pandoc and `pip install weasyprint shapely beautifulsoup4`). Add `--final` to leave out the pictures whose rights are still pending, or `--print` to write the files for a printer (English edition: interior and cover spread, in `book/word/print`). Index entries are listed in `book/index_terms.txt`. For a hand-sewn proof on a home printer (A5 pages, four per A4 sheet, in signatures of 16 pages), run `uv run --with pypdf python3 book/make_booklet.py` after `--print`.

Proposals for Biñan city

* [Innovative Strategies for a Low-Cost, High-Impact City College in Biñan](proposals/College.md)

Night market stand

* [Secret-drawer treasure chest to 3D print](night-market/secret-chest/README.md)
* [Tsinelas keychain to 3D print](night-market/tsinelas-keychain/README.md)
* [Puto Biñan trinket box to 3D print](night-market/puto-box/README.md)
