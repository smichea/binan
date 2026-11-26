"""Build the print-ready PDFs of the Biñan book.

Markdown (book/raw_files) -> HTML (pandoc) -> PDF (WeasyPrint), with a cover
page, a table of contents with page numbers, source links turned into
footnotes, and "Did You Know?" boxes.

Requirements: pandoc, and `pip install weasyprint`.
Usage (from anywhere):  python3 book/build_pdf.py
"""

import html
import unicodedata
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
        "kind": "history", "lang": "en", "author": "Sébastien Michéa",
        "source": BOOK / "raw_files" / "history.md",
        "front": BOOK / "raw_files" / "introduction.md",
        "back": BOOK / "raw_files" / "appendices.md",
        "output": BOOK / "word" / "history_english.pdf",
        "kicker": "",
        "title": "The Book of Biñan",
        "subtitle": "Laguna, Philippines",
        "tagline": "The history of the Philippines through one town",
        "languages": "English edition",
    },
    {
        "kind": "history", "lang": "tl", "author": "Sébastien Michéa",
        "source": BOOK / "raw_files" / "history.md",
        "front": BOOK / "raw_files" / "introduction.md",
        "back": BOOK / "raw_files" / "appendices.md",
        "output": BOOK / "word" / "history_tagalog.pdf",
        "kicker": "Ang Aklat ng Biñan",
        "title": "Kasaysayan ng Lungsod ng Biñan",
        "subtitle": "Laguna, Pilipinas",
        "tagline": "Ang kasaysayan ng Pilipinas sa pamamagitan ng isang bayan",
        "languages": "Edisyong Tagalog",
    },
    {
        "kind": "companion", "lang": "en", "author": "Sébastien Michéa",
        "source": BOOK / "raw_files" / "history_tagalog_explanation.md",
        "output": BOOK / "word" / "history_tagalog_explanation.pdf",
        "kicker": "The Book of Biñan · Companion",
        "title": "Reading the Tagalog Text",
        "subtitle": "Vocabulary, grammar and verbs of Kasaysayan ng Lungsod ng Biñan",
        "tagline": "For learners of Tagalog",
        "languages": "Tagalog → English",
    },
]

# Titles of the figures in the Tagalog edition (list of maps and documents).
TITLES_TL = {
    'location': 'Ang Biñan at ang Laguna de Bay',
    'iss-lake': 'Ang Look ng Maynila at ang Laguna de Bay mula sa kalawakan',
    'tectonics': 'Ang Pilipinas sa pagitan ng dalawang plate',
    'taal-ash': 'Ang abo ng Taal patungong Laguna de Bay, 2020',
    'makiling': 'Ang Bundok Makiling sa kabila ng lawa',
    'westbay': 'Apoy, fault at lawa sa paligid ng Biñan',
    'fault-plane': 'Ang West Valley Fault sa Carmona',
    'lagoon': 'Ang laguna ng tubig-alat, 6,000 taon na ang nakararaan',
    'climate': 'Ang klima ng Biñan',
    'migrations': 'Ang mga unang paglalakbay ng tao sa Timog-Silangang Asyang kapuluan',
    'rhino': 'Ang kinatay na rhinoceros ng Kalinga',
    'luzonensis': 'Homo luzonensis',
    'ill-luzonensis': 'Homo luzonensis, isang guhit-paglalarawan',
    'tabon': 'Ang Kuweba ng Tabon, Palawan',
    'ill-austronesians': 'Dumating sa Luzon ang mga Austronesyong manlalayag',
    'prehistoric-sites': 'Mga sinaunang pook sa Pilipinas',
    'petroglyphs': 'Ang mga Petroglyph ng Angono–Binangonan',
    'manunggul': 'Ang Banga ng Manunggul',
    'ill-village': 'Isang nayon sa baybayin ng lawa, mga 1,000 taon na ang nakararaan',
    'balangay': 'Muling naglayag ang isang balangay',
    'polities': 'Ang Laguna de Bay noong mga 1570',
    'copperplate': 'Ang Inskripsiyon sa Binatbat na Tanso ng Laguna, 900',
    'ill-trade': 'Pakikipagkalakalan sa mga junk ng Tsina',
    'boxer-couple': 'Mag-asawang Tagalog sa Boxer Codex',
    'ill-catalonan': 'Isang pista para sa anito',
    'doctrina': 'Ang baybayin sa Doctrina Christiana, 1593',
    'ill-salcedo': 'Si Salcedo sa lawa, 1571',
    'territory': 'Ang nasasakupan ng Biñan, 1688–1907',
    'murillo': '“Viñan” sa mapa ni Murillo Velarde, 1734',
    'manuscript': '“Biñang” sa isang mapang Espanyol ng ika-18 siglo',
    'map1897': '“Biñang” sa isang mapang militar ng Espanya, 1897',
    'alberto': 'Ang bahay ng mga Alberto',
    'school': 'Kung saan nag-aral si Rizal sa Biñan',
    'tablet': 'Ang plake ng alaala ni José Alberto',
    'lamco': 'Ang talaan ng binyag ni Domingo Lam-co',
    'ill-hacienda': 'Ang asyenda ng mga Dominiko',
    'timbao': 'Ang prinsa ng Timbao',
    'ill-plaza': 'Isang plaza ng bayan noong mga 1870',
    'church2010': 'Ang simbahan ni San Isidro Labrador',
    'ill-cascos': 'Mga kasko sa lawa',
    'gallera': 'Isang sabungan, 1859',
    'rizal-sketch': 'Ang guhit ni Rizal ng rehiyon ng lawa',
    'marker1948': 'Ang panandang “Rizal in Biñan” ng 1948',
    'bautista': 'Si Ambrosio Rianzares Bautista',
    'puck': '“School Begins”, Puck, 1899',
    'kill-every-one': '“Kill every one over ten”, 1902',
    'ill-troops': 'Mga sundalong Amerikano sa Laguna, 1900',
    'gunboat': 'Ang barkong pandigmang Laguna de Bay, 1899',
    'ill-classroom': 'Isang pampublikong paaralan, mga 1905',
    'ill-station': 'Dumating ang tren sa Biñan, 1908',
    'atlas': 'Ang Biñan sa Atlas de Filipinas, 1899–1900',
    'monument': 'Ang bantayog ni Rizal sa plaza, 1918',
    'cinema': 'Ang bahay ng mga Alberto bilang sinehan, dekada 1910',
    'southroad': 'Mga kalabaw sa Manila South Road, 1934',
    'ill-rice': 'Ang bigas sa panahon ng okupasyon, 1943',
    'peso1000': 'Isang perang papel ng okupasyon na 1,000 piso, 1945',
    'limaco': 'Si Rogelio Limaco',
    'ill-guerrillas': 'Isang pulong ng mga gerilya, 1944',
    'liberation': 'Ang paglaya ng Biñan, Pebrero 3, 1945',
    'memorial': 'Ang bantayog ng Hunters-ROTC sa Bulwagang Panlungsod ng Biñan',
    'memorial-text': 'Ang salaysay ni Mojica sa bantayog',
    'mabitac': 'Mabitac, Laguna, Abril 1945',
    'bridge': 'Ang bagong tulay ng Biñan, mga 1951',
    'ams1956': 'Ang bayan ng Biñan sa isang mapa ng 1956',
    'population': 'Ang populasyon ng Biñan, 1903–2024',
    'slex': 'Ang Biñan mula sa himpapawid, 2024',
    'barangays': 'Ang mga barangay ng Biñan ngayon',
    'technopark': 'Ang Laguna Technopark',
    'cityhall': 'Ang Bulwagang Panlungsod ng Biñan',
    'sentro': 'Ang Sentrong Pangkultura ng Biñan',
    'alberto-inside': 'Sa loob ng muling itinayong bahay ng mga Alberto, 2023',
    'view2023': 'Ang Poblacion mula sa bahay ng mga Alberto, 2023',
    'marketstreet': 'Sa paligid ng pampublikong pamilihan, 2013',
    'fruit': 'Isang tindahan ng prutas sa pamilihan, 2013',
    'bandstand': 'Ang kiyosko ng Los Maduros, 2023',
    'bakya': 'Bakya',
    'puto': 'Puto Biñan',
    'pospas': 'Pospas',
    'church2021': 'Ang dambanang pandiyosesis ni San Isidro Labrador, 2021',
    'mass': 'Misa sa San Isidro Labrador, 2013',
    'flowers': 'Mga tindera ng bulaklak, 2013',
    'maglalatik': 'Mga mananayaw ng maglalatik',
    'festival': 'Ang Puto Latik Festival',
    'arawngbinan': 'Ang Araw ng Biñan sa gabi',
    'hdp1953': 'Ang kasaysayan ng Biñan na isinulat noong 1953',
    'pasig1917': 'Mga kasko at barkong de-singaw sa Pasig, mga 1917',
    'station2023': 'Ang istasyon ng Biñan, Hunyo 2023',
    'hybridtrain': 'Ang Hybrid Electric Train sa Biñan, 2019',
    'transport': 'Mga daan, riles at lawa sa paligid ng Biñan',
    'jeepney2013': 'Isang dyip sa Biñan, 2013',
    'n1-2023': 'Ang pambansang daan sa Santo Domingo, 2023',
    'economy': 'Saan naghahanapbuhay ang Biñan',
    'calaxtechno': 'Ang labasan ng CALAX sa Laguna Technopark, 2024',
    'southwoodsmall': 'Southwoods Mall, 2024',
    'pavilion': 'Pavilion Mall, 2013',
    'fishpens': 'Mga baklad sa Laguna de Bay',
    'land': 'Kanino ang lupa?',
}

# The language of the edition being built: "en" or "tl".
LANG = {"code": "en"}


def T(english, tagalog):
    """The English or the Tagalog wording, for the edition being built."""
    return tagalog if LANG["code"] == "tl" else english



# Photos from Wikimedia Commons used in the layout, with their credits.
IMAGES = {
    "cover": {
        "file": "images/san-isidro-church.jpg",
        "credit": "San Isidro Labrador Parish Church, Biñan (cover). Photo: "
                  "patrickroque01, CC BY-SA 4.0, via Wikimedia Commons.",
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
         tagalog="Ang Biñan sa kanlurang baybayin ng Laguna de Bay, sa timog ng Maynila. Bago dumating ang mga Espanyol, bahagi ang lugar ng pamayanang Tagalog ng Tabuco, ang Cabuyao ngayon."),
    dict(chapter=1, after=1, kind="Photo", id="iss-lake", src="images/geography/iss-manila-laguna-2016.jpg",
         title="Manila Bay and Laguna de Bay from space",
         caption="Metro Manila between Manila Bay (left) and Laguna de Bay (right), photographed "
                 "from the International Space Station on 6 May 2016. The Pasig River links the "
                 "lake to the sea.",
         tagalog="Ang Kamaynilaan sa pagitan ng Look ng Maynila (kaliwa) at ng Laguna de Bay (kanan), kuha mula sa International Space Station noong Mayo 6, 2016. Iniuugnay ng Ilog Pasig ang lawa sa dagat.",
         credit="Astronaut photograph ISS047-E-99713. NASA, ISS Crew Earth Observations Facility. Public domain."),
    dict(chapter=1, after=2, kind="Map", id="tectonics", src="maps/tectonics.svg",
         title="The Philippines between two plates",
         caption="The Philippines between the Philippine Sea Plate and the Sunda Plate: the "
                 "subduction trenches, the active faults and some of the active volcanoes. The heavy "
                 "double line is the West Valley Fault, which crosses Biñan.",
         tagalog="Ang Pilipinas sa pagitan ng Philippine Sea Plate at ng Sunda Plate: ang mga trench ng subduction, ang mga aktibong fault at ilan sa mga aktibong bulkan. Ang makapal na dobleng guhit ang West Valley Fault, na tumatawid sa Biñan.",
         credit="Plate boundaries: P. Bird (2003), PB2002, via H. Ahlenius (ODC-By). Faults: GEM "
                "Global Active Faults Database (Styron & Pagani 2020), CC BY-SA."),
    dict(chapter=1, after=3, kind="Photo", id="taal-ash", src="images/geography/iss-taal-ash-2020.jpg",
         title="Taal’s ash drifting toward Laguna de Bay, 2020",
         caption="Ash from Taal Volcano (top left) drifting over Batangas, Cavite and Laguna "
                 "toward Laguna de Bay (top right), seen from the International Space Station on "
                 "22 January 2020.",
         tagalog="Ang abo ng Bulkang Taal (itaas, kaliwa) na tinatangay sa ibabaw ng Batangas, Cavite at Laguna patungo sa Laguna de Bay (itaas, kanan), mula sa International Space Station noong Enero 22, 2020.",
         credit="Expedition 61 crew. NASA, ISS Crew Earth Observations Facility. Public domain."),
    dict(chapter=1, after=3, kind="Photo", id="makiling", src="images/geography/mount-makiling.jpg",
         title="Mount Makiling across the lake",
         caption="Mount Makiling, the extinct volcano of the legend of Maria Makiling, seen "
                 "across Laguna de Bay with fish pens in the foreground.",
         tagalog="Ang Bundok Makiling, ang patay na bulkan ng alamat ni Maria Makiling, na tanaw sa kabila ng Laguna de Bay, kasama ang mga baklad ng isda sa harapan.",
         credit="Photo: Zian Colt, 2021, CC BY-SA 4.0, via Wikimedia Commons."),
    dict(chapter=1, after=4, kind="Map", id="westbay", src="maps/westbay.svg",
         title="Fire, fault and lake around Biñan",
         caption="Biñan between the West Valley Fault and the lake: the approximate outline of "
                 "the Laguna caldera, the volcanoes of the Macolod Corridor, and the former link "
                 "with Manila Bay, when the West Bay was an arm of the sea (6,600–4,700 years ago; "
                 "schematic).",
         tagalog="Ang Biñan sa pagitan ng West Valley Fault at ng lawa: ang tinatayang hugis ng kaldera ng Laguna, ang mga bulkan ng Macolod Corridor, at ang dating ugnayan sa Look ng Maynila, noong bahagi pa ng dagat ang West Bay (6,600–4,700 taon na ang nakararaan; eskematiko).",
         credit="Fault trace: GEM Global Active Faults Database, CC BY-SA. Sea link after "
                "Jaraula et al. 2014."),
    dict(chapter=1, after=4, kind="Photo", id="fault-plane", src="images/geography/wvf-fault-plane-carmona.jpg",
         title="The West Valley Fault at Carmona",
         caption="The scarp of the West Valley Fault, the long grassy slope on the right, in "
                 "Barangay Mabuhay, Carmona, a few kilometres from Biñan.",
         tagalog="Ang dalisdis ng West Valley Fault, ang mahabang madamong libis sa kanan, sa Barangay Mabuhay, Carmona, ilang kilometro mula sa Biñan.",
         credit="Photo: Ervin Malicdem, 2017, CC BY-SA 4.0, via Wikimedia Commons."),
    dict(chapter=1, after=6, kind="Illustration", id="lagoon", src="images/illustrations/salt-lagoon-6000-years.jpg",
         title="The salt lagoon, 6,000 years ago",
         caption="An impression of the shore of Biñan about 6,000 years ago, when the western part "
                 "of the lake was an arm of the sea: mangroves, tidal flats and shellfish gatherers.",
         tagalog="Guhit-paglalarawan ng baybayin ng Biñan mga 6,000 taon na ang nakararaan, noong bahagi pa ng dagat ang kanlurang bahagi ng lawa: mga bakawan, mga putikang inaabot ng taog, at mga nangunguha ng kabibe.", credit="Illustration created for this book (AI-generated, Figma). Artist’s impression."),
    dict(chapter=1, after=7, kind="Chart", id="climate", src="maps/climate.svg",
         title="The climate of Biñan",
         caption="Left: the monsoon year – normal monthly rainfall at Sangley Point, the nearest "
                 "PAGASA station; the wet months are in blue, with darker bars. Right: mean yearly temperature for "
                 "Biñan’s coordinates, 1981–2024, from NASA’s reconstruction, with its trend. The "
                 "model averages a wide area, so its values are lower than those measured on the "
                 "lakeshore; it is the trend that matters.",
         tagalog="Kaliwa: ang taon ng habagat – ang karaniwang buwanang ulan sa Sangley Point, ang pinakamalapit na istasyon ng PAGASA; kulay asul at mas madilim ang mga bar ng mga buwan ng tag-ulan. Kanan: ang katamtamang temperatura bawat taon sa mga koordinada ng Biñan, 1981–2024, ayon sa rekonstruksiyon ng NASA, kasama ang takbo nito. Malawak na lugar ang kinukuhanan ng karaniwan ng modelo, kaya mas mababa ang mga halaga nito kaysa sa nasusukat sa baybayin ng lawa; ang takbo ang mahalaga.",
         credit="PAGASA Climatological Normals; NASA POWER (MERRA-2)."),
    dict(chapter=2, after=1, kind="Map", id="migrations", src="maps/migrations.svg",
         title="Early human journeys to Island Southeast Asia",
         caption="The long journeys of early humans and of our own species from Africa to "
                 "Island Southeast Asia, with some of the oldest dated sites, and the later "
                 "voyages of Austronesian seafarers from Taiwan. Routes are simplified.",
         tagalog="Ang mahahabang paglalakbay ng mga sinaunang tao at ng ating sariling uri mula Aprika hanggang sa mga pulo ng Timog-Silangang Asya, kasama ang ilan sa pinakamatatandang lugar na napetsahan, at ang mga huling paglalayag ng mga Austronesyong manlalayag mula Taiwan. Pinasimple ang mga ruta.",
         credit="Land: Natural Earth (public domain), via world-atlas. Dates from the sources "
                "cited in the text."),
    dict(chapter=2, after=2, kind="Photo", id="rhino", src="images/prehistory/kalinga-rhinoceros.jpg",
         title="The butchered rhinoceros of Kalinga",
         caption="Bones of the extinct Philippine rhinoceros found at Rizal, Kalinga, about 709,000 "
                 "years old: the cut marks left by stone tools are the oldest trace of humans in "
                 "the Philippines.",
         tagalog="Mga buto ng naglahong rhinoceros ng Pilipinas na natagpuan sa Rizal, Kalinga, mga 709,000 taon na ang tanda: ang mga gasgas na iniwan ng mga kasangkapang bato ang pinakamatandang bakas ng tao sa Pilipinas.",
         credit="Photo: Darwgon0801, CC BY-SA 4.0, via Wikimedia Commons."),
    dict(chapter=2, after=2, kind="Photo", id="luzonensis", src="images/prehistory/homo-luzonensis.jpg",
         title="Homo luzonensis",
         caption="Teeth and foot bones of Homo luzonensis from Callao Cave, Cagayan: a human "
                 "species found only on Luzon.",
         tagalog="Mga ngipin at buto ng paa ng Homo luzonensis mula sa Kuweba ng Callao, Cagayan: isang uri ng tao na sa Luzon lamang natagpuan.",
         credit="Diagram: Homo schollensis, CC BY-SA 4.0, via Wikimedia Commons (after Détroit et al. 2019)."),
    dict(chapter=2, after=2, kind="Illustration", id="ill-luzonensis", src="images/illustrations/homo-luzonensis.jpg",
         title="Homo luzonensis, an impression",
         caption="An impression of Homo luzonensis at a cave mouth in Cagayan. Only teeth and a "
                 "few small bones have been found: its true appearance is unknown.",
         tagalog="Guhit-paglalarawan ng Homo luzonensis sa bunganga ng isang kuweba sa Cagayan. Mga ngipin at ilang maliliit na buto lamang ang natagpuan: hindi alam ang tunay nitong anyo.", credit="Illustration created for this book (AI-generated, Figma). Artist’s impression."),
    dict(chapter=2, after=3, kind="Photo", id="tabon", src="images/prehistory/tabon-cave.jpg",
         title="Tabon Cave, Palawan",
         caption="Tabon Cave in Palawan, where the oldest known bones of modern humans in the "
                 "Philippines were found.",
         tagalog="Ang Kuweba ng Tabon sa Palawan, kung saan natagpuan ang pinakamatatandang kilalang buto ng makabagong tao sa Pilipinas.",
         credit="Photo: Jimaggro, 2014, CC BY-SA 4.0, via Wikimedia Commons."),
    dict(chapter=2, after=5, kind="Illustration", id="ill-austronesians", src="images/illustrations/austronesian-arrival.jpg",
         title="Austronesian seafarers reach Luzon",
         caption="An impression of Austronesian seafarers landing in northern Luzon about 4,000 "
                 "years ago, with their outrigger canoes, red-slipped pots, plants and animals.",
         tagalog="Guhit-paglalarawan ng mga Austronesyong manlalayag na dumaraong sa hilagang Luzon mga 4,000 taon na ang nakararaan, kasama ang kanilang mga bangkang may katig, mga palayok na may pulang pahid, mga halaman at mga hayop.", credit="Illustration created for this book (AI-generated, Figma). Artist’s impression."),
    dict(chapter=2, after=5, kind="Map", id="prehistoric-sites", src="maps/prehistoric-sites.svg",
         title="Prehistoric sites of the Philippines",
         caption="Some prehistoric sites of the Philippines, and, in detail, the places around "
                 "Laguna de Bay where ancient traces have been found. None has yet been found in "
                 "Biñan itself.",
         tagalog="Ilang sinaunang pook sa Pilipinas, at, sa detalye, ang mga lugar sa paligid ng Laguna de Bay kung saan natagpuan ang mga sinaunang bakas. Wala pang natatagpuan sa Biñan mismo.",
         credit="Boundaries: PSA / OCHA via philippines-json-maps (MIT)."),
    dict(chapter=2, after=6, kind="Photo", id="petroglyphs", src="images/prehistory/angono-petroglyphs.jpg",
         title="The Angono–Binangonan Petroglyphs",
         caption="Figures carved into a wall of volcanic tuff at Binangonan, Rizal, on the "
                 "northern shore of the lake: among the oldest known artworks of the Philippines.",
         tagalog="Mga pigurang inukit sa isang pader ng bulkanikong tuff sa Binangonan, Rizal, sa hilagang baybayin ng lawa: kabilang sa pinakamatatandang kilalang likhang-sining ng Pilipinas.",
         credit="Photo: Raffy-ish, CC BY-SA 3.0, via Wikimedia Commons."),
    dict(chapter=2, after=6, kind="Photo", id="manunggul", src="images/prehistory/manunggul-jar.jpg",
         title="The Manunggul Jar",
         caption="The Manunggul Jar (Palawan, about 890–710 BCE): on its lid, two figures in a "
                 "boat carry a soul to the afterlife. National Museum of the Philippines.",
         tagalog="Ang Banga ng Manunggul (Palawan, mga 890–710 BCE): sa takip nito, dalawang pigura sa isang bangka ang naghahatid ng isang kaluluwa sa kabilang buhay. Pambansang Museo ng Pilipinas.",
         credit="Photo: Gary Todd, CC0, via Wikimedia Commons."),
    dict(chapter=2, after=8, kind="Illustration", id="ill-village", src="images/illustrations/lakeshore-village.jpg",
         title="A lakeshore village, about 1,000 years ago",
         caption="An impression of a village on Laguna de Bay about a thousand years ago: houses on "
                 "stilts, outrigger canoes, fishing nets, taro and earthen jars, with Mount "
                 "Makiling across the water.",
         tagalog="Guhit-paglalarawan ng isang nayon sa Laguna de Bay mga isang libong taon na ang nakararaan: mga bahay na nakatayo sa mga haligi, mga bangkang may katig, mga lambat, gabi at mga banga, at ang Bundok Makiling sa kabila ng tubig.", credit="Illustration created for this book (AI-generated, Figma). Artist’s impression."),
    dict(chapter=2, after=8, kind="Photo", id="balangay", src="images/prehistory/balangay-replica.jpg",
         title="A balangay sails again",
         caption="A modern replica of a balangay, the plank boat of ancient Butuan, sailing on "
                 "Manila Bay. The word barangay comes from this boat.",
         tagalog="Isang makabagong replika ng balangay, ang bangkang tabla ng sinaunang Butuan, na naglalayag sa Look ng Maynila. Mula sa bangkang ito ang salitang barangay.",
         credit="Photo: Fung360, CC BY-SA 4.0, via Wikimedia Commons."),
    dict(chapter=3, after=1, kind="Map", id="polities", src="maps/polities.svg",
         title="Laguna de Bay around 1570",
         caption="The communities around Laguna de Bay on the eve of the Spanish conquest, with the "
                 "places named in this chapter. No document mentions a settlement on the site of "
                 "Biñan before 1571.",
         tagalog="Ang mga pamayanan sa paligid ng Laguna de Bay sa bisperas ng pananakop ng Espanya, kasama ang mga lugar na binanggit sa kabanatang ito. Walang dokumentong bumabanggit ng pamayanan sa kinaroroonan ng Biñan bago ang 1571.",
         credit="Boundaries: PSA / OCHA via philippines-json-maps (MIT). Places after Blair & "
                "Robertson, vols. 3, 7 and 8."),
    dict(chapter=3, after=2, kind="Photo", id="copperplate",
         src="images/archives/copperplate-c900.jpg",
         title="The Laguna Copperplate Inscription, 900",
         caption="The Laguna Copperplate Inscription, dated 21 April 900: the oldest known written "
                 "document of the Philippines, found at the end of the 1980s near the mouth of the "
                 "Lumbang River, on the other shore of our lake.",
         tagalog="Ang Inskripsiyon sa Binatbat na Tanso ng Laguna, na may petsang Abril 21, 900: ang pinakamatandang kilalang nakasulat na dokumento ng Pilipinas, na natagpuan sa pagtatapos ng dekada 1980 malapit sa bukana ng Ilog Lumbang, sa kabilang pampang ng ating lawa.",
         credit="Photo: Gary Todd, CC0, via Wikimedia Commons. National Museum of the Philippines."),
    dict(chapter=3, after=3, kind="Illustration", id="ill-trade", src="images/illustrations/chinese-trade.jpg",
         title="Trading with the Chinese junks",
         caption="An impression of trade at a river mouth in the 13th century, after Zhao Rugua’s "
                 "description of Ma-i: porcelain, iron pots and glass beads against cotton, wax and "
                 "cloth, and a white parasol offered to the datu.",
         tagalog="Guhit-paglalarawan ng kalakalan sa isang bukana ng ilog noong ika-13 siglo, batay sa paglalarawan ni Zhao Rugua sa Ma-i: porselana, mga kalderong bakal at mga butil na salamin kapalit ng bulak, pagkit at tela, at isang puting payong na inialay sa datu.", credit="Illustration created for this book (AI-generated, Figma). Artist’s impression."),
    dict(chapter=3, after=5, kind="Photo", id="boxer-couple", small=True, src="images/archives/boxer-codex-tagalog-couple.png",
         title="A Tagalog couple in the Boxer Codex",
         caption="A Tagalog couple of rank, painted in Manila around 1590 for the Boxer Codex: the "
                 "man in a kangan jacket and putong headcloth, the woman in baro and saya, both "
                 "wearing gold.",
         tagalog="Isang mag-asawang Tagalog na may mataas na katayuan, ipininta sa Maynila noong mga 1590 para sa Boxer Codex: nakasuot ang lalaki ng kangan at putong, nakabaro at saya ang babae, at kapwa sila may suot na ginto.",
         credit="Boxer Codex, c. 1590, Lilly Library, Indiana University. Public domain."),
    dict(chapter=3, after=6, kind="Illustration", id="ill-catalonan", src="images/illustrations/catalonan-feast.jpg",
         title="A feast for the anito",
         caption="An impression of a feast for the anito, the ancestor spirits, led by a catalonan, "
                 "after the descriptions of 1572 and 1589. Clothing is simplified.",
         tagalog="Guhit-paglalarawan ng isang pagdiriwang para sa mga anito, ang mga espiritu ng mga ninuno, na pinamumunuan ng isang katalonan, batay sa mga paglalarawan ng 1572 at 1589. Pinasimple ang mga kasuotan.", credit="Illustration created for this book (AI-generated, Figma). Artist’s impression."),
    dict(chapter=3, after=6, kind="Photo", id="doctrina", src="images/archives/doctrina-christiana-1593.jpg",
         title="Baybayin in the Doctrina Christiana, 1593",
         caption="A page of the Doctrina Christiana printed in Manila in 1593, the first book of the "
                 "Philippines: Tagalog prayers in baybayin script above, the Ave Maria in Spanish below.",
         tagalog="Isang pahina ng Doctrina Christiana na inilimbag sa Maynila noong 1593, ang unang aklat ng Pilipinas: mga dasal na Tagalog sa sulat-baybayin sa itaas, at ang Ave Maria sa Espanyol sa ibaba.",
         credit="Library of Congress, Rare Book and Special Collections, Rosenwald Collection 1302. "
                "Public domain."),
    dict(chapter=3, after=7, kind="Illustration", id="ill-salcedo", src="images/illustrations/salcedo-galley-1571.jpg",
         title="Salcedo on the lake, 1571",
         caption="An impression of Juan de Salcedo’s galley reaching a lakeshore village in the "
                 "second half of 1571.",
         tagalog="Guhit-paglalarawan ng galera ni Juan de Salcedo na dumarating sa isang nayon sa baybayin ng lawa sa ikalawang hati ng 1571.", credit="Illustration created for this book (AI-generated, Figma). Artist’s impression."),
    dict(chapter=4, after=4, kind="Map", id="territory", src="maps/territory.svg",
         title="The territory of Biñan, 1688–1907",
         caption="How Biñan’s territory took shape: separated from Tabuco in 1688, it lost "
                 "its barrio of Santa Rosa in 1791–92. Under American rule, San Pedro "
                 "(1903–1907) and Muntinlupa (1903–1905) were briefly joined to it. "
                 "Present-day boundaries.",
         tagalog="Kung paano nabuo ang teritoryo ng Biñan: inihiwalay ito sa Tabuco noong 1688, at nawala rito ang baryo ng Santa Rosa noong 1791–92. Sa ilalim ng mga Amerikano, sandaling idinugtong dito ang San Pedro (1903–1907) at ang Muntinlupa (1903–1905). Kasalukuyang mga hangganan."),
    dict(chapter=4, after=1, kind="Map", id="murillo", src="images/archives/murillo-velarde-1734.jpg",
         title="“Viñan” on the Murillo Velarde map, 1734",
         caption="The south-western shore of Laguna de Bay on Pedro Murillo Velarde’s 1734 map of the Philippines, the "
                 "first scientific map of the islands. “Viñan” appears at the left edge, on the shore "
                 "near Tunasan and Calamba.",
         tagalog="Ang timog-kanlurang baybayin ng Laguna de Bay sa mapa ng Pilipinas ni Pedro Murillo Velarde (1734), ang unang siyentipikong mapa ng kapuluan. Makikita ang “Viñan” sa kaliwang gilid, sa baybayin malapit sa Tunasan at Calamba.",
         credit="Pedro Murillo Velarde, Carta hydrographica y chorographica de las Yslas "
                "Filipinas (Manila, 1734). Library of Congress, Geography and Map Division. "
                "Public domain."),
    dict(chapter=4, after=2, kind="Map", id="manuscript", src="images/archives/manuscript-laguna-18th-c.jpg",
         title="“Biñang” on an 18th-century Spanish map",
         caption="Detail of an 18th-century Spanish manuscript map of the surroundings of "
                 "Laguna de Bay: the town of “Biñang” with its church (small square), the river "
                 "mouth (“Bª de Biñang”), Malabag (today Malaban), Sinalsan, Landayan, San "
                 "Pedro Tunasan and Santa Rosa.",
         tagalog="Detalye ng isang manuskritong mapang Espanyol ng ika-18 siglo ng paligid ng Laguna de Bay: ang bayan ng “Biñang” at ang simbahan nito (maliit na parisukat), ang bukana ng ilog (“Bª de Biñang”), ang Malabag (Malaban ngayon), Sinalsan, Landayan, San Pedro Tunasan at Santa Rosa.",
         credit="Fuente: Archivo Cartográfico de Estudios Geográficos del Centro Geográfico "
                "del Ejército / Biblioteca Virtual de Defensa. CC BY 4.0."),
    dict(chapter=4, after=8, kind="Map", id="map1897", src="images/archives/cavite-region-map-1897.jpg",
         title="“Biñang” on a Spanish military map, 1897",
         caption="“Biñang” and its barrios on the Spanish military map of the province of Cavite "
                 "and the neighbouring region, drawn in 1897 during the Philippine Revolution.",
         tagalog="Ang “Biñang” at ang mga baryo nito sa mapang militar ng Espanya ng lalawigan ng Cavite at ng karatig na rehiyon, na iginuhit noong 1897, sa panahon ng Rebolusyong Pilipino.",
         credit="Provincia de Cavite y región limítrofe, 1897. Fuente: Archivo Histórico de la "
                "Armada Juan Sebastián de Elcano / Biblioteca Virtual de Defensa. CC BY 4.0."),
    dict(chapter=4, after=6, kind="Photo", id="alberto", src="images/alberto-mansion.jpg",
         title="The Alberto house",
         caption="The Alberto house on the plaza of Biñan, about two hundred years old: the home of the family of José Rizal’s mother, Teodora Alonso, where she grew up. Its later fate is told in chapter VII.",
         tagalog="Ang bahay ng mga Alberto sa plaza ng Biñan, na mga dalawandaang taon na: ang tahanan ng pamilya ng ina ni José Rizal, si Teodora Alonso, kung saan siya lumaki. Isinasalaysay sa kabanata VII ang sinapit nito pagkatapos.",
         credit="Photo: Bengoeswhere, CC BY-SA 3.0, via Wikimedia Commons."),
    dict(chapter=4, after=7, kind="Photo", id="school", src="images/archives/justiniano-school-site-pre1918.jpg",
         title="Where Rizal went to school in Biñan",
         caption="“Where Rizal went to school in Biñan: site of Master Justiniano’s house, which burned down many years ago.” The young José Rizal studied here in 1870; the site is now the School of Dr. José Rizal museum.",
         tagalog="“Kung saan nag-aral si Rizal sa Biñan: ang kinatayuan ng bahay ni Maestro Justiniano, na nasunog maraming taon na ang nakalipas.” Dito nag-aral ang batang si José Rizal noong 1870; museo na ngayon ang lugar, ang School of Dr. José Rizal.",
         credit="Austin Craig, Rizal’s Own Story of His Life (Manila, 1918). Internet Archive. "
                "Public domain."),
    dict(chapter=4, after=6, kind="Photo", id="tablet", small=True,
         src="images/archives/jose-alberto-tablet-pre1913.jpg",
         title="Memorial tablet of José Alberto",
         caption="Memorial tablet of José Alberto, Rizal’s uncle and owner of the Alberto house, "
                 "who died on 29 October 1875, in the church of Biñan.",
         tagalog="Ang lapida ni José Alberto, tiyo ni Rizal at may-ari ng bahay ng mga Alberto, na namatay noong Oktubre 29, 1875, sa simbahan ng Biñan.",
         credit="Austin Craig, Lineage, Life and Labors of José Rizal (Manila, 1913). "
                "Internet Archive. Public domain."),
    dict(chapter=4, after=2, kind="Photo", id="lamco", small=True, src="images/spanish/lamco-baptism.jpg",
         title="The baptism record of Domingo Lam-co",
         caption="The record of the baptism of Domingo Lam-co, a Chinese convert who settled on the "
                 "Dominican estate of Biñan and became an ancestor of José Rizal.",
         tagalog="Ang talaan ng binyag ni Domingo Lam-co, isang Tsinong nagbalik-loob na nanirahan sa asyenda ng mga Dominiko sa Biñan at naging ninuno ni José Rizal.",
         credit="Austin Craig, Lineage, Life and Labors of José Rizal (1913). Project Gutenberg. Public domain."),
    dict(chapter=4, after=5, kind="Illustration", id="ill-hacienda", src="images/illustrations/dominican-hacienda.jpg",
         title="The Dominican hacienda",
         caption="An impression of the Dominican estate of Biñan: irrigated rice fields fed by a stone "
                 "dam, sugar cane and a carabao-driven press in the gardens, tenants and their "
                 "sharecroppers.",
         tagalog="Guhit-paglalarawan ng asyenda ng mga Dominiko sa Biñan: mga palayang pinatutubigan mula sa isang prinsang bato, tubuhan at isang gilingang pinaiikot ng kalabaw sa mga taniman, at ang mga inkilino at kanilang mga kasama.", credit="Illustration created for this book (AI-generated, Figma). Artist’s impression."),
    dict(chapter=4, after=5, kind="Photo", id="timbao", src="images/spanish/timbao-dam.jpg",
         title="The Timbao dam",
         caption="The Timbao dam, one of the two surviving dams of the old hacienda’s irrigation "
                 "system, declared local heritage in 2023.",
         tagalog="Ang prinsa ng Timbao, isa sa dalawang natitirang prinsa ng sistema ng patubig ng lumang asyenda, na idineklarang lokal na pamana noong 2023.",
         credit="Photo: Totoyba2, CC BY-SA 4.0, via Wikimedia Commons."),
    dict(chapter=4, after=6, kind="Illustration", id="ill-plaza", src="images/illustrations/binan-plaza-1870.jpg",
         title="A town plaza around 1870",
         caption="An impression of a Laguna town plaza around 1870, with its stone church, "
                 "bahay-na-bato houses and market. The church is not a portrait of Biñan’s.",
         tagalog="Guhit-paglalarawan ng plaza ng isang bayan sa Laguna noong mga 1870, kasama ang simbahang bato, mga bahay-na-bato at palengke. Hindi larawan ng simbahan ng Biñan ang simbahang ito.", credit="Illustration created for this book (AI-generated, Figma). Artist’s impression."),
    dict(chapter=4, after=6, kind="Photo", id="church2010", src="images/spanish/san-isidro-church-2010.jpg",
         title="The church of San Isidro Labrador",
         caption="The church of San Isidro Labrador, heart of the town since Spanish times (photographed in 2010). The present building was completed in 1970, after a fire in 1968.",
         tagalog="Ang simbahan ni San Isidro Labrador, puso ng bayan mula pa noong panahon ng Espanyol (kuha noong 2010). Natapos ang kasalukuyang gusali noong 1970, matapos ang isang sunog noong 1968.",
         credit="Photo: Monbrye, public domain, via Wikimedia Commons."),
    dict(chapter=4, after=6, kind="Illustration", id="ill-cascos", src="images/illustrations/cascos-on-the-lake.jpg",
         title="Cascos on the lake",
         caption="An impression of cascos, the roofed cargo barges that carried rice, sugar and cloth "
                 "from the lake towns to Manila.",
         tagalog="Guhit-paglalarawan ng mga kasko, ang mga bangkang pangkargang may bubong na nagdadala ng palay, asukal at tela mula sa mga bayan sa baybayin ng lawa patungong Maynila.", credit="Illustration created for this book (AI-generated, Figma). Artist’s impression."),
    dict(chapter=4, after=6, kind="Photo", id="gallera", small=True, src="images/spanish/bowring-gallera.jpg",
         title="A cockpit, 1859",
         caption="A gallera, or cockpit, as drawn for Sir John Bowring’s account of his journey to "
                 "La Laguna in 1858–59.",
         tagalog="Isang gallera, o sabungan, ayon sa guhit para sa salaysay ni Sir John Bowring tungkol sa kanyang paglalakbay sa La Laguna noong 1858–59.",
         credit="John Bowring, A Visit to the Philippine Islands (1859). Project Gutenberg. Public domain."),
    dict(chapter=4, after=7, kind="Photo", id="rizal-sketch", small=True, src="images/spanish/rizal-sketch-map.jpg",
         title="Rizal’s sketch of the lake region",
         caption="A sketch map of the Laguna de Bay region drawn by José Rizal.",
         tagalog="Isang mapa ng rehiyon ng Laguna de Bay na iginuhit ni José Rizal.",
         credit="Austin Craig, Lineage, Life and Labors of José Rizal (1913). Project Gutenberg. Public domain."),
    dict(chapter=4, after=7, kind="Photo", id="marker1948", src="images/spanish/rizal-marker-1948.jpg",
         title="The 1948 marker “Rizal in Biñan”",
         caption="The historical marker placed in 1948 at the site of Rizal’s school. It gives the "
                 "year 1871; Rizal’s own memoirs point to 1870.",
         tagalog="Ang makasaysayang panandang inilagay noong 1948 sa kinatayuan ng paaralan ni Rizal. Taong 1871 ang nakasulat dito; ayon sa sariling mga alaala ni Rizal, 1870 ito.",
         credit="Photo: Andre Ryan B. Maceda (perspective correction: Eugene Alvin Villar), CC BY-SA 4.0, via Wikimedia Commons."),
    dict(chapter=4, after=8, kind="Photo", id="bautista", small=True, src="images/spanish/ambrosio-rianzares-bautista.jpg",
         title="Ambrosio Rianzares Bautista",
         caption="Ambrosio Rianzares Bautista (1830–1903), born in Biñan, author of the Act of the "
                 "Proclamation of Philippine Independence of 12 June 1898.",
         tagalog="Si Ambrosio Rianzares Bautista (1830–1903), na isinilang sa Biñan, ang sumulat ng Akta ng Pagpapahayag ng Kasarinlan ng Pilipinas noong Hunyo 12, 1898.",
         credit="Public domain, via Wikimedia Commons."),
    dict(chapter=5, after=1, kind="Photo", id="puck", src="images/american/puck-school-begins-1899.jpg",
         title="“School Begins”, Puck, 1899",
         caption="“School Begins”, a cartoon published in the American magazine Puck on 25 January "
                 "1899, a week before the war: Uncle Sam lectures his “new class in civilization”. "
                 "Its caricatures show the racist view of the colonized peoples common in the United "
                 "States at the time.",
         tagalog="“School Begins”, isang karikaturang inilathala sa magasing Amerikanong Puck noong Enero 25, 1899, isang linggo bago ang digmaan: tinuturuan ni Uncle Sam ang kanyang “bagong klase sa sibilisasyon”. Ipinakikita ng mga karikatura nito ang rasistang pagtingin sa mga sinakop na bayan na karaniwan noon sa Estados Unidos.",
         credit="Louis Dalrymple, Puck, 1899. Library of Congress, Prints and Photographs Division. Public domain."),
    dict(chapter=5, after=2, kind="Photo", id="kill-every-one", small=True,
         src="images/american/kill-every-one-over-ten-1902.jpg",
         title="“Kill every one over ten”, 1902",
         caption="“Kill every one over ten”: a cartoon in the New York Evening Journal of 5 May 1902, "
                 "denouncing General Jacob Smith’s order in Samar. Its caption reads: “Criminals "
                 "because they were born ten years before we took the Philippines.”",
         tagalog="“Kill every one over ten” (Patayin ang lahat ng lampas sampung taon): isang karikatura sa New York Evening Journal noong Mayo 5, 1902, na tumutuligsa sa utos ni Heneral Jacob Smith sa Samar. Ayon sa kapsiyon nito: “Mga kriminal dahil isinilang sila sampung taon bago namin kinuha ang Pilipinas.”",
         credit="Homer Davenport, New York Evening Journal, 1902. Public domain."),
    dict(chapter=5, after=3, kind="Illustration", id="ill-troops", src="images/illustrations/us-troops-laguna-1900.jpg",
         title="American troops in Laguna, 1900",
         caption="An impression of an American column on a Laguna road in January 1900, during the "
                 "campaign that took Biñan.",
         tagalog="Guhit-paglalarawan ng isang hanay ng mga sundalong Amerikano sa isang daan sa Laguna noong Enero 1900, sa panahon ng kampanyang sumakop sa Biñan.", credit="Illustration created for this book (AI-generated, Figma). Artist’s impression."),
    dict(chapter=5, after=3, kind="Photo", id="gunboat", src="images/american/gunboat-laguna-de-bay-1899.jpg",
         title="The gunboat Laguna de Bay, 1899",
         caption="The American river gunboat Laguna de Bay, 1899. Gunboats on the lake supported the "
                 "attacks on the lakeshore towns, and in January 1900 the Napindan laid a telegraph "
                 "cable from Taguig to Biñan.",
         tagalog="Ang Amerikanong barkong pandigma na Laguna de Bay, 1899. Tumulong ang mga barkong pandigma sa lawa sa pagsalakay sa mga bayan sa baybayin, at noong Enero 1900, naglatag ang Napindan ng kable ng telegrapo mula Taguig hanggang Biñan.",
         credit="Unknown photographer, 1899. Public domain, via Wikimedia Commons."),
    dict(chapter=5, after=7, kind="Illustration", id="ill-classroom", src="images/illustrations/thomasite-classroom.jpg",
         title="A public school, around 1905",
         caption="An impression of a public school classroom in a Laguna town around 1905, with an "
                 "American teacher and a Filipino assistant.",
         tagalog="Guhit-paglalarawan ng isang silid-aralan ng pampublikong paaralan sa isang bayan sa Laguna noong mga 1905, kasama ang isang gurong Amerikano at isang katulong na Pilipino.", credit="Illustration created for this book (AI-generated, Figma). Artist’s impression."),
    dict(chapter=5, after=7, kind="Illustration", id="ill-station", src="images/illustrations/binan-station-1908.jpg",
         title="The train reaches Biñan, 1908",
         caption="An impression of the arrival of the Manila Railroad at Biñan, opened on 2 September "
                 "1908, with puto vendors on the platform.",
         tagalog="Guhit-paglalarawan ng pagdating ng Manila Railroad sa Biñan, na binuksan noong Setyembre 2, 1908, kasama ang mga nagtitinda ng puto sa plataporma.", credit="Illustration created for this book (AI-generated, Figma). Artist’s impression."),
    dict(chapter=5, after=3, kind="Map", id="atlas", src="images/archives/atlas-filipinas-1899.jpg",
         title="Biñan in the Atlas de Filipinas, 1899–1900",
         caption="“Biñang” in the Atlas de Filipinas of the Manila Observatory (Map No. 9, "
                 "U.S. Coast and Geodetic Survey edition, 1900), on the road from Manila to "
                 "Calamba, as the Americans found it.",
         tagalog="Ang “Biñang” sa Atlas de Filipinas ng Obserbatoryo ng Maynila (Mapa Blg. 9, edisyon ng U.S. Coast and Geodetic Survey, 1900), sa daan mula Maynila hanggang Calamba, gaya ng nadatnan ng mga Amerikano.",
         credit="Observatorio de Manila / U.S. Coast and Geodetic Survey, Atlas of the "
                "Philippine Islands (Washington, 1900), Map No. 9. Internet Archive. Public domain."),
    dict(chapter=5, after=7, kind="Photo", id="monument", src="images/archives/rizal-monument-plaza-pre1918.jpg",
         title="The Rizal monument on the plaza, 1918",
         caption="The Rizal monument on the town plaza, newly built, with the Alberto house – the "
                 "home of Rizal’s mother’s family – behind it.",
         tagalog="Ang bantayog ni Rizal sa plaza ng bayan, bagong tayo, at sa likod nito ang bahay ng mga Alberto – ang tahanan ng pamilya ng ina ni Rizal.",
         credit="Austin Craig, Rizal’s Own Story of His Life (Manila, 1918). Internet Archive. "
                "Public domain."),
    dict(chapter=5, after=7, kind="Photo", id="cinema", src="images/archives/alberto-house-cinematograph-pre1918.jpg",
         title="The Alberto house as a cinema, 1910s",
         caption="“Rizal’s uncle’s home in Biñan”: in the 1910s the ground floor of the Alberto "
                 "house served as the town’s cinematograph – a sign of the new American-era "
                 "leisure, from movies to baseball.",
         tagalog="“Rizal’s uncle’s home in Biñan” (Ang bahay ng tiyo ni Rizal sa Biñan): noong dekada 1910, nagsilbing sinematograpo ng bayan ang unang palapag ng bahay ng mga Alberto – tanda ng mga bagong libangan sa panahon ng Amerikano, mula sa sine hanggang sa beysbol.",
         credit="Austin Craig, Rizal’s Own Story of His Life (Manila, 1918). Internet Archive. "
                "Public domain."),
    dict(chapter=5, after=8, kind="Photo", id="southroad", rights="pending",
         src="images/archives/pendleton-south-road-1934.jpg",
         title="Carabaos on the Manila South Road, 1934",
         caption="Carabaos on the Manila South Road between Muntinlupa and Biñan, looking toward "
                 "Biñan, 6 December 1934. Rain trees line the national road; rice fields and a "
                 "telegraph line run alongside.",
         tagalog="Mga kalabaw sa Manila South Road sa pagitan ng Muntinlupa at Biñan, nakaharap sa Biñan, Disyembre 6, 1934. May mga punong akasya sa tabi ng pambansang daan; may mga palayan at linya ng telegrapo sa gilid nito.",
         credit="Robert L. Pendleton, 1934. American Geographical Society Library, University "
                "of Wisconsin–Milwaukee Libraries (negative 409.11)."),
    dict(chapter=6, after=2, kind="Illustration", id="ill-rice", src="images/illustrations/occupation-rice-1943.jpg",
         title="Rice under the occupation, 1943",
         caption="An impression of rice being collected for the Japanese army in a Laguna town, 1943, "
                 "while families queue for their ration.",
         tagalog="Guhit-paglalarawan ng pagkuha ng bigas para sa hukbong Hapones sa isang bayan sa Laguna, 1943, habang nakapila ang mga pamilya para sa kanilang rasyon.", credit="Illustration created for this book (AI-generated, Figma). Artist’s impression."),
    dict(chapter=6, after=2, kind="Photo", id="peso1000", small=True,
         src="images/wwii/occupation-1000-pesos-1945.jpg",
         title="A 1,000-peso occupation note, 1945",
         caption="Front and back of a 1,000-peso note issued by the Japanese government in the "
                 "Philippines in 1945. Filipinos called this money “Mickey Mouse money”: by the "
                 "end of the war it was almost worthless.",
         tagalog="Harap at likod ng isang perang papel na 1,000 piso na inilabas ng pamahalaang Hapones sa Pilipinas noong 1945. “Mickey Mouse money” ang tawag ng mga Pilipino sa perang ito: sa pagtatapos ng digmaan, halos wala na itong halaga.",
         credit="National Numismatic Collection, National Museum of American History (Smithsonian). Public domain, via Wikimedia Commons."),
    dict(chapter=6, after=4, kind="Photo", id="limaco", small=True, css="portrait", rights="pending",
         src="images/wwii/rogelio-limaco-portrait.jpg",
         title="Rogelio Limaco",
         caption="Rogelio Limaco (1913–1944), mayor of Biñan, seized by the Japanese on 24 October 1944 "
                 "and never seen again. His portrait hangs with those of the other mayors in the city museum.",
         tagalog="Si Rogelio Limaco (1913–1944), alkalde ng Biñan, na dinakip ng mga Hapones noong Oktubre 24, 1944 at hindi na muling nakita. Nakasabit ang kanyang larawan kasama ng sa iba pang mga alkalde sa museo ng lungsod.",
         credit="Portrait by Cristobal, 1999, Biñan City Museum. Photo: Judgefloro, 2018, CC0, via Wikimedia Commons."),
    dict(chapter=6, after=5, kind="Illustration", id="ill-guerrillas", src="images/illustrations/guerrillas-1944.jpg",
         title="A guerrilla meeting, 1944",
         caption="An impression of a night meeting of guerrillas in a lakeshore house, 1944: most of "
                 "the resistance was made of civilians who carried food, messages and information.",
         tagalog="Guhit-paglalarawan ng isang pulong ng mga gerilya sa gabi sa isang bahay sa baybayin ng lawa, 1944: karamihan sa paglaban ay binubuo ng mga sibilyang nagdadala ng pagkain, mensahe at impormasyon.", credit="Illustration created for this book (AI-generated, Figma). Artist’s impression."),
    dict(chapter=6, after=6, kind="Map", id="liberation", src="maps/liberation.svg",
         title="The liberation of Biñan, 3 February 1945",
         caption="The liberation of Biñan, 3 February 1945. According to Proculo Mojica, the Hunters-ROTC "
                 "came down from Carmona in three columns: one attacked the garrison, one sealed the "
                 "roads from San Pedro, one blocked the way to Santa Rosa. Arrows are schematic.",
         tagalog="Ang paglaya ng Biñan, Pebrero 3, 1945. Ayon kay Proculo Mojica, bumaba mula Carmona ang Hunters-ROTC sa tatlong hanay: sinalakay ng isa ang garison, isinara ng isa ang mga daan mula San Pedro, at hinarangan ng isa ang daan patungong Santa Rosa. Eskematiko ang mga palaso."),
    dict(chapter=6, after=6, kind="Photo", id="memorial", src="images/wwii/hunters-rotc-memorial-binan.jpg",
         title="The Hunters-ROTC memorial at Biñan City Hall",
         caption="The memorial to the Hunters-ROTC guerrillas, “the army that freed Biñan in the time of "
                 "the Japanese”, unveiled at City Hall on 3 February 2018.",
         tagalog="Ang bantayog ng mga gerilyang Hunters-ROTC, “ang hukbong nagpalaya sa Biñan noong panahon ng Hapon”, na pinasinayaan sa Bulwagang Panlungsod noong Pebrero 3, 2018.",
         credit="Photo: Judgefloro, 2018. CC0, via Wikimedia Commons."),
    dict(chapter=6, after=6, kind="Photo", id="memorial-text", src="images/wwii/hunters-rotc-memorial-text.jpg",
         title="Mojica’s account on the memorial",
         caption="The plaque quotes Proculo Mojica’s Terry’s Hunters (1965): “In one day, the entire "
                 "garrison was annihilated and Biñan fell to the Hunters.”",
         tagalog="Sinisipi ng plake ang Terry’s Hunters (1965) ni Proculo Mojica: “Sa loob ng isang araw, nalipol ang buong garison at bumagsak ang Biñan sa mga Hunters.”",
         credit="Photo: Judgefloro, 2018. CC0, via Wikimedia Commons."),
    dict(chapter=6, after=8, kind="Photo", id="mabitac", src="images/wwii/mabitac-1945.jpg",
         title="Mabitac, Laguna, April 1945",
         caption="The ruins of Mabitac, on the eastern shore of Laguna de Bay, photographed by the U.S. Army "
                 "on 4 April 1945. Biñan, behind the front since February, was spared such destruction.",
         tagalog="Ang mga guho ng Mabitac, sa silangang baybayin ng Laguna de Bay, kuha ng hukbong Amerikano noong Abril 4, 1945. Nakaligtas sa ganitong pagkawasak ang Biñan, na nasa likod na ng labanan mula noong Pebrero.",
         credit="U.S. Army Signal Corps, 1945. Public domain, via Wikimedia Commons."),
    dict(chapter=7, after=1, kind="Photo", id="bridge", src="images/archives/binan-bridge-1951.jpg",
         title="The new Biñan bridge, c. 1951",
         caption="“Biñan Bridge completed”, about 1951: one of the bridges of the national "
                 "road rebuilt after the war under the Philippine–American public roads "
                 "rehabilitation programme.",
         tagalog="“Biñan Bridge completed” (Tapos na ang Tulay ng Biñan), mga 1951: isa sa mga tulay ng pambansang daan na muling itinayo matapos ang digmaan sa ilalim ng programang Pilipino-Amerikano para sa rehabilitasyon ng mga daan.",
         credit="U.S. National Archives, Records of the Bureau of Public Roads (RG 30), "
                "30-N-51-4622. Public domain."),
    dict(chapter=7, after=1, kind="Map", id="ams1956", src="images/archives/ams-san-pedro-1956.jpg",
         title="The town of Biñan on a 1956 survey map",
         caption="The town of Biñan on the 1956 U.S. Army Map Service sheet “San Pedro” "
                 "(1:50,000): the street grid of the Poblacion and the neighbouring barrios "
                 "of Canlalay, Malaban, De La Paz and San Antonio, between the national road "
                 "and the lake.",
         tagalog="Ang bayan ng Biñan sa mapang “San Pedro” ng U.S. Army Map Service, 1956 (1:50,000): ang ayos ng mga kalye ng Poblacion at ang mga karatig na baryo ng Canlalay, Malaban, De La Paz at San Antonio, sa pagitan ng pambansang daan at ng lawa.",
         credit="U.S. Army Map Service, San Pedro (sheet 3263 III), 1:50,000, 1956. "
                "Courtesy of the University of Texas Libraries. Public domain."),
    dict(chapter=7, after=4, kind="Chart", id="population", src="maps/population.svg",
         title="Population of Biñan, 1903–2024",
         caption="Population of Biñan at each census. The town grew slowly until the war, then "
                 "doubled every fifteen years or so; the expressway, the industrial parks of the "
                 "1990s and the subdivisions turned it into the most populous city of Laguna.",
         tagalog="Ang populasyon ng Biñan sa bawat senso. Dahan-dahang lumaki ang bayan hanggang sa digmaan, saka dumoble halos tuwing labinlimang taon; ang expressway, ang mga industrial park ng dekada 1990 at ang mga subdibisyon ang gumawa rito na pinakamataong lungsod ng Laguna."),
    dict(chapter=7, after=4, kind="Photo", id="slex", src="images/modern/slex-from-air-2024.jpg",
         title="Biñan from the air, 2024",
         caption="Biñan from the air in April 2024: the dense city between Laguna de Bay (top) and "
                 "the South Luzon Expressway, with industrial buildings on the right.",
         tagalog="Ang Biñan mula sa himpapawid noong Abril 2024: ang siksik na lungsod sa pagitan ng Laguna de Bay (itaas) at ng South Luzon Expressway, kasama ang mga gusaling pang-industriya sa kanan.",
         credit="Photo: Patrickroque01, 2024, CC BY-SA 4.0, via Wikimedia Commons."),
    dict(chapter=7, after=5, kind="Map", id="barangays", src="maps/barangays.svg",
         title="The barangays of Biñan today",
         caption="Biñan today and its 24 barangays (boundaries approximate). The Poblacion keeps the Spanish-era heart of the town; hatching marks the barangays where the Laguna Technopark (Biñan, Loma, Timbao) and the Laguna International Industrial Park (Mamplasan, Ganado) were built from 1990.",
         tagalog="Ang Biñan ngayon at ang 24 na barangay nito (tinatayang mga hangganan). Nasa Poblacion ang puso ng bayan mula pa noong panahon ng Espanyol; minarkahan ng guhit-guhit ang mga barangay kung saan itinayo mula 1990 ang Laguna Technopark (Biñan, Loma, Timbao) at ang Laguna International Industrial Park (Mamplasan, Ganado)."),
    dict(chapter=7, after=5, kind="Photo", id="technopark", src="images/modern/technopark-sign-2018.jpg",
         title="The Laguna Technopark",
         caption="The entrance sign of the Laguna Technopark, opened in 1990 on former sugarcane "
                 "land in Santa Rosa and Biñan, listing some of its manufacturers.",
         tagalog="Ang karatula sa pasukan ng Laguna Technopark, na binuksan noong 1990 sa dating tubuhan sa Santa Rosa at Biñan, na naglilista ng ilan sa mga pabrika nito.", credit="Photo: Judgefloro, 2018, CC0, via Wikimedia Commons."),
    dict(chapter=7, after=6, kind="Photo", id="cityhall", src="images/modern/city-hall-2018.jpg",
         title="Biñan City Hall",
         caption="The Biñan City Hall complex in Zapote, seat of the city government since Biñan "
                 "became a city in 2010.",
         tagalog="Ang Bulwagang Panlungsod ng Biñan sa Zapote, luklukan ng pamahalaang panlungsod mula nang maging lungsod ang Biñan noong 2010.", credit="Photo: Judgefloro, 2018, CC0, via Wikimedia Commons."),
    dict(chapter=7, after=7, kind="Photo", id="sentro", src="images/modern/sentrong-pangkultura-2023.jpg",
         title="The Sentrong Pangkultura ng Biñan",
         caption="The old municipal hall – the Dominicans’ casa hacienda, then a school, then the "
                 "town hall from the 1950s – restored as the Sentrong Pangkultura ng Biñan under the "
                 "Balik-Biñan programme.",
         tagalog="Ang lumang munisipyo – ang casa hacienda ng mga Dominiko, na naging paaralan at, mula noong dekada 1950, munisipyo – na ipinanumbalik bilang Sentrong Pangkultura ng Biñan sa ilalim ng programang Balik-Biñan.", credit="Photo: Ralff Nestor Nacor, 2023, CC BY-SA 4.0, via Wikimedia Commons."),
    dict(chapter=7, after=7, kind="Photo", id="alberto-inside", src="images/modern/alberto-mansion-interior-2023.jpg",
         title="Inside the rebuilt Alberto house, 2023",
         caption="The arcaded courtyard of the Alberto house, rebuilt by the city and reopened in "
                 "December 2019.",
         tagalog="Ang patyong may mga arko ng bahay ng mga Alberto, na muling itinayo ng lungsod at binuksang muli noong Disyembre 2019.",
         credit="Photo: Ralff Nestor Nacor, 2023, CC BY-SA 4.0, via Wikimedia Commons."),
    dict(chapter=7, after=8, kind="Photo", id="view2023", src="images/modern/view-from-alberto-mansion-2023.jpg",
         title="The Poblacion from the Alberto house, 2023",
         caption="The heart of Biñan seen from the balcony of the Alberto house: the plaza, the old "
                 "municipal hall and the church of San Isidro Labrador, rebuilt after the fire of 1968.",
         tagalog="Ang puso ng Biñan na tanaw mula sa balkonahe ng bahay ng mga Alberto: ang plaza, ang lumang munisipyo at ang simbahan ni San Isidro Labrador, na muling itinayo matapos ang sunog noong 1968.", credit="Photo: Ralff Nestor Nacor, 2023, CC BY-SA 4.0, via Wikimedia Commons."),
    dict(chapter=8, after=1, kind="Photo", id="marketstreet", src="images/modern/market-street-2013.jpg",
         title="Around the public market, 2013",
         caption="Shops and shoppers around the public market of Biñan on a rainy morning.",
         tagalog="Mga tindahan at mamimili sa paligid ng pampublikong pamilihan ng Biñan sa isang maulang umaga.",
         credit="Photo: Ramon F. Velasquez, 2013, CC BY-SA 3.0, via Wikimedia Commons."),
    dict(chapter=8, after=1, kind="Photo", id="fruit", src="images/modern/market-fruit-stall-2013.jpg",
         title="A fruit stall in the market, 2013",
         caption="A fruit stall at the Biñan market: bananas, citrus and local produce.",
         tagalog="Isang tindahan ng prutas sa pamilihan ng Biñan: saging, dalandan at mga lokal na produkto.",
         credit="Photo: Ramon F. Velasquez, 2013, CC BY-SA 3.0, via Wikimedia Commons."),
    dict(chapter=8, after=2, kind="Photo", id="bandstand", src="images/today/los-maduros-bandstand-2023.jpg",
         title="The Los Maduros bandstand, 2023",
         caption="The Los Maduros bandstand on Plaza Rizal, given by a local gentlemen’s club in the 1950s or 1960s; its shape recalls a wooden clog or a slipper, in memory of the town’s shoemakers.",
         tagalog="Ang kiyosko ng Los Maduros sa Plaza Rizal, handog ng isang samahan ng mga ginoo noong dekada 1950 o 1960; ipinaaalala ng hugis nito ang bakya o tsinelas, bilang pag-alala sa mga sapatero ng bayan.",
         credit="Photo: Ralff Nestor Nacor, 2023, CC BY-SA 4.0, via Wikimedia Commons."),
    dict(chapter=8, after=2, kind="Photo", id="bakya", small=True, src="images/today/bakya.png",
         title="Bakya, wooden clogs",
         caption="Bakya, wooden clogs like those once made in Casile and San Jose (photographed elsewhere in the Philippines).",
         tagalog="Mga bakya, tulad ng mga dating ginagawa sa Casile at San Jose (kuha sa ibang lugar sa Pilipinas).",
         credit="Photo: Joan Villarante-Kavianifar, 2007, CC BY-SA 2.0, via Wikimedia Commons."),
    dict(chapter=8, after=3, kind="Photo", id="puto", rights="pending", src="images/culture/puto-binan.jpg",
         title="Puto Biñan",
         caption="Puto Biñan, topped with grated cheese and salted egg.",
         tagalog="Ang Puto Biñan, na nilagyan ng ginadgad na keso at itlog na maalat.",
         credit="Experience Biñan, City Government of Biñan (BCHATO), experiencebinan.com."),
    dict(chapter=8, after=3, kind="Photo", id="pospas", rights="pending", src="images/culture/pospas.jpg",
         title="Pospas",
         caption="Pospas, Biñan’s yellow porridge of glutinous rice with chicken, fried garlic, spring onion and pork crackling.",
         tagalog="Ang pospas, ang dilaw na lugaw na malagkit ng Biñan na may manok, pritong bawang, dahon ng sibuyas at chicharon.",
         credit="Experience Biñan, City Government of Biñan (BCHATO), experiencebinan.com."),
    dict(chapter=8, after=4, kind="Photo", id="church2021", src="images/today/san-isidro-church-2021.jpg",
         title="The diocesan shrine of San Isidro Labrador, 2021",
         caption="The church of San Isidro Labrador, completed in 1970 after the fire of 1968 and declared a diocesan shrine in 2022.",
         tagalog="Ang simbahan ni San Isidro Labrador, na natapos noong 1970 matapos ang sunog noong 1968 at idineklarang dambanang pandiyosesis noong 2022.",
         credit="Photo: patrickroque01, 2021, CC BY-SA 4.0, via Wikimedia Commons."),
    dict(chapter=8, after=4, kind="Photo", id="mass", src="images/modern/church-mass-2013.jpg",
         title="Mass at San Isidro Labrador, 2013",
         caption="A full church for Mass at San Isidro Labrador, the heart of the town’s religious life.",
         tagalog="Punong simbahan sa isang misa sa San Isidro Labrador, ang puso ng buhay-pananampalataya ng bayan.",
         credit="Photo: Ramon F. Velasquez, 2013, CC BY-SA 3.0, via Wikimedia Commons."),
    dict(chapter=8, after=4, kind="Photo", id="flowers", src="images/modern/market-flowers-2013.jpg",
         title="Flower vendors, 2013",
         caption="Flower vendors outside the market, preparing offerings for the church and for family celebrations.",
         tagalog="Mga nagtitinda ng bulaklak sa labas ng pamilihan, na naghahanda ng mga handog para sa simbahan at sa mga pagdiriwang ng pamilya.",
         credit="Photo: Ramon F. Velasquez, 2013, CC BY-SA 3.0, via Wikimedia Commons."),
    dict(chapter=8, after=5, kind="Photo", id="maglalatik", rights="pending", src="images/culture/maglalatik.jpg",
         title="Maglalatik dancers",
         caption="Maglalatik dancers, coconut shells strapped to their bodies: the “Christians” in blue and the “Moros” in red.",
         tagalog="Mga mananayaw ng maglalatik na may mga bao ng niyog na nakatali sa katawan: ang mga “Kristiyano” na nakaasul at ang mga “Moro” na nakapula.",
         credit="Experience Biñan, City Government of Biñan (BCHATO), experiencebinan.com."),
    dict(chapter=8, after=5, kind="Photo", id="festival", rights="pending", src="images/culture/puto-latik-festival.jpg",
         title="The Puto Latik Festival",
         caption="Street dancers in the Puto Latik Festival, held each May since 2017.",
         tagalog="Mga mananayaw sa kalye sa Puto Latik Festival, na idinaraos tuwing Mayo mula noong 2017.",
         credit="Experience Biñan, City Government of Biñan (BCHATO), experiencebinan.com."),
    dict(chapter=8, after=5, kind="Photo", id="arawngbinan", rights="pending", src="images/culture/araw-ng-binan-night.jpg",
         title="Araw ng Biñan by night",
         caption="Araw ng Biñan in February: the illuminated Sentrong Pangkultura and crowds on the plaza.",
         tagalog="Ang Araw ng Biñan tuwing Pebrero: ang naiilawang Sentrong Pangkultura at ang mga tao sa plaza.",
         credit="Experience Biñan, City Government of Biñan (BCHATO), experiencebinan.com."),
    dict(chapter=8, after=7, kind="Photo", id="hdp1953", rights="pending",
         src="images/today/hdp-binan-town-1953.jpg",
         title="The history of Biñan written in 1953",
         caption="The first page of the “History and Cultural Life of the Town of Biñan”, typed by the "
                 "town’s teachers for the Historical Data Papers of 1952–53. Section III gives three "
                 "explanations of the name: the banian tree, binyagan and the nickname Viñiang.",
         tagalog="Ang unang pahina ng “History and Cultural Life of the Town of Biñan”, na makinilyado ng mga guro ng bayan para sa Historical Data Papers ng 1952–53. Ibinibigay ng bahaging III ang tatlong paliwanag sa pangalan: ang punong banian, ang binyagan at ang palayaw na Viñiang.",
         credit="Historical Data Papers, Laguna: Biñan (scan NLP00NNHDBIL29001). National Library of the Philippines."),
    dict(chapter=9, after=1, kind="Photo", id="pasig1917", src="images/today/pasig-steamers-1917.jpg",
         title="Barges and steamers on the Pasig, c. 1917",
         caption="Barges, steamers and sailing boats on the Pasig River in Manila, around 1917: the waterway that linked the lake towns, Biñan among them, to the capital.",
         tagalog="Mga bangkang pangkarga, barkong de-singaw at bangkang may layag sa Ilog Pasig sa Maynila noong mga 1917: ang daang-tubig na nag-uugnay sa mga bayan sa lawa, kabilang ang Biñan, sa kabisera.",
         credit="Photo: Henry S. Townsend, c. 1917. Public domain, via Wikimedia Commons."),
    dict(chapter=9, after=2, kind="Photo", id="station2023", src="images/today/pnr-binan-station-2023.jpg",
         title="Biñan station, June 2023",
         caption="Biñan station on General Malvar Street, a few days before it closed on 2 July 2023 to make way for the North–South Commuter Railway.",
         tagalog="Ang istasyon ng Biñan sa Kalye General Malvar, ilang araw bago ito isinara noong Hulyo 2, 2023 upang bigyang-daan ang North–South Commuter Railway.",
         credit="Photo: Patrickroque01, 2023, CC BY-SA 4.0, via Wikimedia Commons."),
    dict(chapter=9, after=2, kind="Photo", id="hybridtrain", src="images/today/dost-hybrid-train-2019.jpg",
         title="The Hybrid Electric Train at Biñan, 2019",
         caption="The inauguration of the Hybrid Electric Train of the Department of Science and Technology, the first train made in the Philippines, at Biñan station on 24 April 2019.",
         tagalog="Ang pagpapasinaya sa Hybrid Electric Train ng Kagawaran ng Agham at Teknolohiya, ang unang tren na ginawa sa Pilipinas, sa istasyon ng Biñan noong Abril 24, 2019.",
         credit="Photo: JC John Sese Cuneta, 2019, CC BY-SA 4.0, via Wikimedia Commons."),
    dict(chapter=9, after=3, kind="Map", id="transport", src="maps/transport.svg",
         title="Roads, railway and lake around Biñan",
         caption="Roads, railway and lake around Biñan. The expressways (SLEX and CALAX) with the exit numbers, the national roads N1 and N65, the railway being rebuilt as the North–South Commuter Railway, and the planned lakeshore road (schematic).",
         tagalog="Mga daan, riles at lawa sa paligid ng Biñan: ang mga expressway (SLEX at CALAX) at ang bilang ng mga labasan, ang mga pambansang daang N1 at N65, ang riles na muling itinatayo bilang North–South Commuter Railway, at ang planong daan sa baybayin (eskematiko).",
         credit="Roads and railway: © OpenStreetMap contributors (ODbL), via Nominatim and OSRM; stations: Wikidata."),
    dict(chapter=9, after=4, kind="Photo", id="jeepney2013", src="images/today/jeepney-binan-2013.jpg",
         title="A jeepney in Biñan, 2013",
         caption="A Sarao jeepney near Biñan Central Mall, 2013.",
         tagalog="Isang dyip na Sarao malapit sa Biñan Central Mall, 2013.",
         credit="Photo: Ramon F. Velasquez, 2013, CC BY-SA 3.0, via Wikimedia Commons."),
    dict(chapter=9, after=4, kind="Photo", id="n1-2023", src="images/today/n1-santo-domingo-2023.jpg",
         title="The national road at Santo Domingo, 2023",
         caption="The national road (N1) at Santo Domingo, with traditional and modern jeepneys, seen from a footbridge in 2023.",
         tagalog="Ang pambansang daan (N1) sa Santo Domingo, kasama ang mga tradisyonal at modernong dyip, tanaw mula sa isang tulay-tawiran noong 2023.",
         credit="Photo: Patrickroque01, 2023, CC BY-SA 4.0, via Wikimedia Commons."),
    dict(chapter=10, after=1, kind="Map", id="economy", src="maps/economy.svg",
         title="Where Biñan makes a living",
         caption="Where Biñan makes a living: the industrial parks of the south, the market and shops of the old centre, the malls, the fishing barangays on the lake and the barangays where some farmland remains. Boundaries and fish-pen area are approximate.",
         tagalog="Kung saan naghahanapbuhay ang Biñan: ang mga industrial park sa timog, ang pamilihan at mga tindahan sa lumang kabayanan, ang mga mall, ang mga barangay ng mangingisda sa lawa at ang mga barangay na may natitira pang bukirin. Tinatayang hangganan at lugar ng baklad lamang ang nasa mapa.",
         credit="Places: © OpenStreetMap contributors (ODbL), via Nominatim; land use after the city (Experience Biñan)."),
    dict(chapter=9, after=3, kind="Photo", id="calaxtechno", small=True, src="images/today/calax-technopark-2024.jpg",
         title="The Laguna Technopark exit of the CALAX, 2024",
         caption="The Laguna Technopark exit of the Cavite–Laguna Expressway, opened in 2020 in Timbao.",
         tagalog="Ang labasan ng Laguna Technopark sa Cavite–Laguna Expressway, na binuksan noong 2020 sa Timbao.",
         credit="Photo: LMP 2001, 2024, CC BY-SA 4.0, via Wikimedia Commons."),
    dict(chapter=10, after=2, kind="Photo", id="southwoodsmall", src="images/today/southwoods-mall-2024.jpg",
         title="Southwoods Mall, 2024",
         caption="Southwoods Mall, opened in September 2017 in barangay San Francisco.",
         tagalog="Ang Southwoods Mall, na binuksan noong Setyembre 2017 sa barangay San Francisco.",
         credit="Photo: Ralff Nestor Nacor, 2024, CC BY-SA 4.0, via Wikimedia Commons."),
    dict(chapter=10, after=2, kind="Photo", id="pavilion", src="images/today/pavilion-mall-2013.jpg",
         title="Pavilion Mall, 2013",
         caption="The entrance of Pavilion Mall on the national road, Biñan’s first large mall.",
         tagalog="Ang pasukan ng Pavilion Mall sa pambansang daan, ang unang malaking mall ng Biñan.",
         credit="Photo: Ramon F. Velasquez, 2013, CC BY-SA 3.0, via Wikimedia Commons."),
    dict(chapter=10, after=5, kind="Photo", id="fishpens", src="images/today/fish-pens-laguna-de-bay-2010.jpg",
         title="Fish pens on Laguna de Bay",
         caption="Fish pens on Laguna de Bay, photographed off Cardona, Rizal, in 2010: the same kind of bamboo enclosures stand off Biñan.",
         tagalog="Mga baklad sa Laguna de Bay, kuha sa harap ng Cardona, Rizal, noong 2010: ganito rin ang mga kulungang kawayan sa harap ng Biñan.",
         credit="Photo: Shubert Ciencia, 2010, CC BY 2.0, via Wikimedia Commons."),
    dict(chapter=11, after=3, kind="Map", id="land", src="maps/land.svg",
         title="Who owns the land?",
         caption="The largest documented landholdings of Biñan, numbered as in the table. The circles show each area to scale at an approximate place; Southwoods City (2) extends into Carmona. The hatched strip along the shore is the lake bed, public land; the railway right of way also belongs to the State.",
         tagalog="Ang pinakamalalaking naitalang lupain sa Biñan, na may bilang na gaya ng sa talahanayan. Ipinakikita ng mga bilog ang lawak ng bawat isa sa tamang sukat, sa tinatayang lugar; umaabot sa Carmona ang Southwoods City (2). Ang guhit-guhit na bahagi sa baybayin ay ang ilalim ng lawa, lupang pampubliko; pag-aari rin ng Estado ang daanan ng riles.",
         credit="Areas: proclamations, court decisions and company reports (see the chapter); positions approximate."),
]


# Documents still under copyright are marked rights="pending" in FIGURES: they
# appear in the draft with a notice, and are left out by `build_pdf.py --final`.
FINAL = "--final" in sys.argv

ROMAN = ["I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX", "X", "XI", "XII",
         "XIII", "XIV", "XV", "XVI", "XVII", "XVIII", "XIX", "XX"]
LETTERS = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"


def active_figures():
    return [f for f in FIGURES if not (FINAL and f.get("rights") == "pending")]


KIND_TL = {"Map": "Mapa", "Photo": "Larawan", "Illustration": "Guhit", "Chart": "Tsart"}


def kind_label(kind):
    return T(kind, KIND_TL.get(kind, kind))


def figure_html(fig, number):
    credit = f' <span class="credit">{fig["credit"]}</span>' if fig.get("credit") else ""
    if fig.get("rights") == "pending":
        credit += (' <span class="rights-pending">'
                   + T("Reproduction rights being requested.",
                       "Hinihingi pa ang pahintulot na mailathala ito.") + '</span>')

    css = "map" if fig["src"].endswith(".svg") else "archive"
    if fig.get("small"):
        css += " small"
    if fig.get("css"):
        css += " " + fig["css"]
    if fig["kind"] == "Illustration":
        css += " impression"
    caption = T(fig["caption"], fig.get("tagalog") or fig["caption"])
    return (
        f'<figure class="{css}" id="fig-{fig["id"]}"><img src="{fig["src"]}" alt="">'
        f'<figcaption><span class="map-no">{kind_label(fig["kind"])} {number}.</span> '
        f'{caption}{credit}</figcaption></figure>'
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
# Footnotes and the list of sources
# --------------------------------------------------------------------------

# Every source cited in the history, in order: (section, label, url). They are
# listed by chapter in the "Sources" section at the end of the book.
SOURCES = []
CURRENT = {"section": "", "record": False}


def short_url(url):
    """Readable form of a URL: no scheme, no #:~:text= part."""
    parts = urlsplit(url)
    host = parts.netloc.removeprefix("www.")
    path = unquote(parts.path).rstrip("/")
    query = f"?{parts.query}" if parts.query else ""
    return f"{host}{path}{query}"


def clean_url(url):
    return url.split("#:~:text=")[0]


def split_label(label):
    """'Title | Site' -> ('Title', 'Site')."""
    if " | " in label:
        title, site = label.rsplit(" | ", 1)
        return title.strip(), site.strip()
    return label.strip(), ""


def footnote(label, url):
    """A footnote naming the source; the link stays clickable in the e-book,
    and the full address is printed in the list of sources."""
    url = clean_url(url)
    if CURRENT["record"]:
        SOURCES.append((CURRENT["section"], label.strip(), url))
    title, site = split_label(label)
    text = html.escape(title) + (f", {html.escape(site)}" if site else "")
    return f'<span class="fn"><a href="{html.escape(url)}">{text}</a></span>'


# A citation in parentheses: "text ([Label](url))"  -> footnote after text.
CITATION = re.compile(r"\s*\(\[([^\]]+)\]\((https?://[^)\s]+)\)\)")
# A plain link inside the sentence: "[words](url)" -> words + footnote.
INLINE_LINK = re.compile(r"\[([^\]]+)\]\((https?://[^)\s]+)\)")


def links_to_footnotes(text):
    text = CITATION.sub(lambda m: footnote(m.group(1), m.group(2)), text)
    return INLINE_LINK.sub(
        lambda m: m.group(1) + footnote(m.group(1), m.group(2)), text
    )


def sort_key(text):
    plain = unicodedata.normalize("NFKD", re.sub(r"<[^>]+>", "", text))
    plain = "".join(c for c in plain if not unicodedata.combining(c))
    return re.sub(r"^[^0-9A-Za-z]+", "", plain).casefold()


def inline_markdown(text):
    """Escape a label for HTML and turn *x* into italics."""
    return re.sub(r"\*([^*]+)\*", r"<i>\1</i>", html.escape(text))


def sources_section():
    groups = {}
    for section, label, url in SOURCES:
        groups.setdefault(section, {}).setdefault(url, label)
    parts = []
    for section, entries in groups.items():
        rows = []
        for url, label in sorted(entries.items(), key=lambda kv: sort_key(kv[1])):
            title, site = split_label(label)
            site = f", {inline_markdown(site)}" if site else ""
            rows.append(f'<li><span class="bib-title">{inline_markdown(title)}</span>{site}. '
                        f'<a href="{html.escape(url)}">{html.escape(short_url(url))}</a></li>')
        parts.append(f'<h2 class="bib-group">{html.escape(section)}</h2>'
                     f'<ul class="bib">{"".join(rows)}</ul>')
    intro = T("The sources cited in the footnotes, chapter by chapter, with their web addresses "
              "(shortened: add https:// in front). In the electronic edition, every footnote is "
              "also a link.",
              "Ang mga sangguniang binanggit sa mga talababa, ayon sa kabanata, kasama ang "
              "kanilang web address (pinaikli: idagdag ang https:// sa unahan). Sa elektronikong "
              "edisyon, bawat talababa ay isa ring link.")
    return (
        '<section class="sources">'
        f'<h1 class="backmatter" id="sources">{T("Sources", "Mga Sanggunian")}</h1>'
        f'<p class="bm-intro">{intro}</p>'
        + "".join(parts) + "</section>"
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


def bilingual(text):
    """'English / Tagalog' -> ('English', 'Tagalog')."""
    if " / " in text:
        english, tagalog = (part.strip() for part in text.split(" / ", 1))
        return english, tagalog
    return text.strip(), ""


def split_heading(line, css="chapter", both=False):
    """'## English / Tagalog' -> the heading in the language of the edition
    (both=True: English, with the Tagalog part as subtitle)."""
    level, text = line.split(" ", 1)
    english, tagalog = bilingual(text)
    if both and tagalog:
        return f'{level} {english} {{.{css} data-tl="{html.escape(tagalog)}"}}'
    return f"{level} {T(english, tagalog or english)} {{.{css}}}"


def part_page(text):
    """'# Part One · History / Unang Bahagi · Kasaysayan' -> a part title page."""
    english, tagalog = bilingual(text)
    anchor = "part-" + re.sub(r"[^a-z]+", "-", english.partition(" · ")[0].lower()).strip("-")
    number, _, title = T(english, tagalog or english).partition(" · ")
    return (f'<section class="part-page"><h1 class="part" id="{anchor}" '
            f'data-no="{html.escape(number)}">{html.escape(title)}</h1></section>')


FIGURE_NUMBERS = {}


def place_figures(chapter, paragraph, out):
    """Insert the figures that follow the n-th paragraph of a chapter."""
    here = [f for f in active_figures() if (f["chapter"], f["after"]) == (chapter, paragraph)]
    row = []  # consecutive photos are set side by side, two per row
    for fig in here:
        FIGURE_NUMBERS[fig["kind"]] = FIGURE_NUMBERS.get(fig["kind"], 0) + 1
        html_fig = figure_html(fig, FIGURE_NUMBERS[fig["kind"]])
        if fig["kind"] == "Photo" and not fig.get("small"):
            row.append(html_fig)
            if len(row) == 2:
                out.append(f'<div class="fig-row">{"".join(row)}</div>')
                row = []
        else:
            if row:
                out.append(row.pop())
            out.append(html_fig)
    if row:
        out.append(row.pop())


MONTHS_TL = {"January": "Enero", "February": "Pebrero", "March": "Marso", "April": "Abril",
             "May": "Mayo", "June": "Hunyo", "July": "Hulyo", "August": "Agosto",
             "September": "Setyembre", "October": "Oktubre", "November": "Nobyembre",
             "December": "Disyembre"}
MONTH = "(" + "|".join(MONTHS_TL) + ")"


def date_tl(text):
    """English dates of the appendices -> Tagalog ('15 May 1955' -> 'Mayo 15, 1955')."""
    text = re.sub(rf"(\d{{1,2}}) and (\d{{1,2}}) {MONTH} (\d{{3,4}})",
                  lambda m: f"{MONTHS_TL[m[3]]} {m[1]} at {m[2]}, {m[4]}", text)
    text = re.sub(rf"(\d{{1,2}}) {MONTH} (\d{{3,4}})",
                  lambda m: f"{MONTHS_TL[m[2]]} {m[1]}, {m[3]}", text)
    text = re.sub(rf"{MONTH} (\d{{3,4}})", lambda m: f"{MONTHS_TL[m[1]]} {m[2]}", text)
    for english, tagalog in [("late (\\d{4})s", r"huling bahagi ng dekada \1"),
                             ("late (\\d{4})", r"huling bahagi ng \1"),
                             ("(\\d{4})s", r"dekada \1"),
                             ("at least ", "hindi bababa sa "), ("c\\. ", "mga "),
                             (" million", " milyon"), ("years ago", "taon na ang nakararaan")]:
        text = re.sub(english, tagalog, text)
    return text


# Table columns kept in one edition only, and Tagalog column headings.
COLUMNS_EN_ONLY = {"Event", "Note", "English"}
COLUMNS_TL_ONLY = {"Pangyayari", "Tala", "Tagalog"}
HEADINGS_TL = {"Date": "Petsa", "Years": "Mga taon", "Head of the town": "Pinuno ng bayan",
               "Year": "Taon", "Population": "Populasyon", "Source": "Sanggunian",
               "Word": "Salita", "Tagalog": "Kahulugan"}
CELLS_TL = [("Census of ", "Senso ng "), ("Census (PSA)", "Senso (PSA)"),
            ("Proclamation No. 973 of 2025", "Proklamasyon Blg. 973 ng 2025"),
            ("Buzeta and Bravo", "Buzeta at Bravo"),
            ("“souls” of the parish", "mga “kaluluwa” ng parokya"),
            ("officer-in-charge, not yet known", "officer-in-charge, hindi pa alam"),
            ("not yet known", "hindi pa alam")]


def table_for_edition(block):
    """A bilingual table of the appendices -> the columns of one edition."""
    rows = [[c.strip() for c in line.strip().strip("|").split("|")] for line in block.splitlines()]
    header, rule, body = rows[0], rows[1], rows[2:]
    drop = COLUMNS_TL_ONLY if LANG["code"] == "en" else COLUMNS_EN_ONLY
    keep = [k for k, name in enumerate(header) if name not in drop]
    out = []
    for n, row in enumerate([header, rule] + body):
        cells = [row[k] if k < len(row) else "" for k in keep]
        if LANG["code"] == "tl" and n == 0:
            cells = [HEADINGS_TL.get(c, c) for c in cells]
        elif LANG["code"] == "tl" and n > 1:
            cells = [date_tl(c) for c in cells]
            for english, tagalog in CELLS_TL:
                cells = [c.replace(english, tagalog) for c in cells]
        out.append("| " + " | ".join(cells) + " |")
    return "\n".join(out)


def citations(text):
    """The sources cited in a paragraph, in order: [(label, url)]."""
    found = [(m.start(), m.group(1), m.group(2)) for m in CITATION.finditer(text)]
    rest = CITATION.sub("", text)
    found += [(10 ** 6 + m.start(), m.group(1), m.group(2)) for m in INLINE_LINK.finditer(rest)]
    seen, result = set(), []
    for _, label, url in sorted(found):
        if url not in seen:
            seen.add(url)
            result.append((label, url))
    return result


def prepare_text(markdown, kind="chapter"):
    """Mark up the text of one edition (English or Tagalog) for the book layout.

    The source is bilingual: in each chapter the English sections come first,
    then the same sections in Tagalog. Only one language is kept. The sources
    are cited in the English paragraphs; in the Tagalog edition, each paragraph
    receives the footnotes of its English twin.

    kind: "chapter" (the history: numbered chapters with figures), "front"
    (introduction) or "appendix" (back matter)."""
    blocks = [b.strip() for b in markdown.split("\n\n") if b.strip()]
    out = []
    want_tl = LANG["code"] == "tl"
    chapter = paragraphs = 0
    english_sources, tagalog_count = [], 0
    current_tl = False

    def next_text(i):
        for later in blocks[i + 1:]:
            if not later.startswith("#"):
                return later
        return ""

    skip = set()
    for i, block in enumerate(blocks):
        if i in skip:
            continue
        if block.startswith("# "):
            if block[2:].startswith("Part "):
                out.append(part_page(block[2:]))
            continue  # the H1 title of the book is on the cover
        if block.startswith("## "):
            current_tl = False
            english_sources, tagalog_count = [], 0
            english, tagalog = bilingual(block[3:])
            title = T(english, tagalog or english)
            if kind == "chapter":
                chapter += 1
                paragraphs = 0
                CURRENT["section"] = f"{ROMAN[chapter - 1]}. {title}"
                out.append(split_heading(block, "chapter"))
            else:
                CURRENT["section"] = title
                out.append(split_heading(block, "front" if kind == "front" else "appendix"))
            continue
        if block.startswith("### "):
            current_tl = is_tagalog(next_text(i))
            if current_tl == want_tl:
                out.append(f"### {block[4:].strip()} {{.section}}")
            continue
        if block.startswith("**Did You Know?**"):
            english = block.removeprefix("**Did You Know?**").strip()
            text = links_to_footnotes(english)
            if i + 1 < len(blocks) and blocks[i + 1].startswith("**Alam Mo Ba?**"):
                skip.add(i + 1)
                if want_tl:
                    tagalog = blocks[i + 1].removeprefix("**Alam Mo Ba?**").strip()
                    text = links_to_footnotes(tagalog) + "".join(
                        footnote(label, url) for label, url in citations(english))
            out.append("\n".join(['::::: {.didyouknow}',
                                  f'<p class="dyk-title">{T("Did you know?", "Alam mo ba?")}</p>',
                                  "", text, ":::::"]))
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
        if block.startswith("<"):
            out.append(block)
            continue
        if block.startswith("|"):
            if kind == "appendix":
                out.append(table_for_edition(block))
            elif current_tl == want_tl:
                out.append(block)  # the tables of the chapters exist in each language
            continue
        current_tl = is_tagalog(block)
        if not current_tl:
            english_sources.append(citations(block))
        if current_tl != want_tl:
            continue
        if current_tl:
            twin = english_sources[tagalog_count] if tagalog_count < len(english_sources) else []
            tagalog_count += 1
            out.append(links_to_footnotes(block) + "".join(footnote(l, u) for l, u in twin))
        else:
            out.append(links_to_footnotes(block))
        if kind == "chapter":
            paragraphs += 1
            place_figures(chapter, paragraphs, out)
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
            css = "front" if stripped.startswith("## Introduction") else "chapter"
            out.extend(["", split_heading(stripped, css, both=True), ""])
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
# Index
# --------------------------------------------------------------------------

INDEX_TERMS = BOOK / "index_terms.txt"
SKIP_INDEX = {"fn", "credit", "lang-label", "dyk-title", "map-no", "rights-pending"}


def index_entries():
    """Lines 'Heading = variant | variant' (or just 'Heading') -> [(heading, variants)]."""
    entries = []
    if not INDEX_TERMS.exists():
        return entries
    for line in INDEX_TERMS.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#"):
            continue
        heading, _, variants = line.partition(" = ")
        names = [v.strip() for v in variants.split(" | ")] if variants else [heading.strip()]
        entries.append((heading.strip(), [n for n in names if n]))
    return entries


def mark_index(body):
    """Wrap every occurrence of an index term in the text with an anchor.
    Returns the new HTML and {heading: [anchor ids]}."""
    from bs4 import BeautifulSoup, NavigableString

    entries = index_entries()
    if not entries:
        return body, {}
    variants = sorted({(v, h) for h, vs in entries for v in vs}, key=lambda x: -len(x[0]))
    heading_of = {v: h for v, h in variants}
    pattern = re.compile(r"(?<!\w)(" + "|".join(re.escape(v) for v, _ in variants) + r")(?!\w)")
    soup = BeautifulSoup(body, "html.parser")
    anchors, counter = {}, [0]

    def excluded(node):
        for parent in node.parents:
            if parent.name in ("h1", "a", "figcaption", "table", "style", "script"):
                return True
            if set(parent.get("class") or []) & SKIP_INDEX:
                return True
        return False

    for text in list(soup.find_all(string=True)):
        if not isinstance(text, NavigableString) or not pattern.search(text) or excluded(text):
            continue
        pieces, last = [], 0
        for match in pattern.finditer(text):
            pieces.append(html.escape(text[last:match.start()]))
            counter[0] += 1
            anchor = f"ix{counter[0]}"
            anchors.setdefault(heading_of[match.group(1)], []).append(anchor)
            pieces.append(f'<span class="ix" id="{anchor}">{html.escape(match.group(1))}</span>')
            last = match.end()
        pieces.append(html.escape(text[last:]))
        text.replace_with(BeautifulSoup("".join(pieces), "html.parser"))
    return str(soup), anchors


INDEX_PLACEHOLDER = "<!--INDEX-->"


def index_section(anchors, pages):
    rows = []
    for heading in sorted(anchors, key=sort_key):
        numbers = sorted({pages[a] for a in anchors[heading] if a in pages})
        if numbers:
            rows.append(f'<li><span class="ix-head">{html.escape(heading)}</span> '
                        f'{", ".join(str(n) for n in numbers)}</li>')
    return f'<ul class="index">{"".join(rows)}</ul>'


# --------------------------------------------------------------------------
# Book assembly
# --------------------------------------------------------------------------

H1 = re.compile(r"<h1\b([^>]*)>(.*?)</h1>", re.S)


def attribute(attrs, name):
    match = re.search(rf'\b{name}="([^"]*)"', attrs)
    return match.group(1) if match else ""


def table_of_contents(html_text):
    items, chapter, appendix = [], 0, 0
    for attrs, inner in H1.findall(html_text):
        css, anchor = attribute(attrs, "class"), attribute(attrs, "id")
        if not anchor or css not in ("chapter", "front", "appendix", "backmatter", "part"):
            continue
        title = " ".join(re.sub(r"<[^>]+>", "", inner).split())
        tagalog = attribute(attrs, "data-tl")
        tl = f' <span class="toc-tl">{tagalog}</span>' if tagalog else ""
        if css == "part":
            items.append(f'<li class="toc-part"><a href="#{anchor}">'
                         f'{attribute(attrs, "data-no")} · {title}</a>{tl}</li>')
            continue
        number = ""
        if css == "chapter":
            chapter += 1
            number = f'<span class="toc-no">{ROMAN[chapter - 1]}.</span>'
        elif css == "appendix":
            appendix += 1
            number = f'<span class="toc-no">{LETTERS[appendix - 1]}.</span>'
        items.append(f'<li class="toc-{css}"><a class="toc-link" href="#{anchor}">{number}'
                     f'<span class="toc-en">{title}</span>{tl}</a></li>')
    return (
        f'<section class="toc"><h1 class="toc-title">{T("Contents", "Nilalaman")}</h1>'
        f'<ol class="toc-main">{"".join(items)}</ol></section>'
    )


def figures_section(html_text):
    """The list of maps, photos and documents, at the end of the book."""
    present = set(re.findall(r'<figure class="[^"]+" id="fig-([^"]+)"', html_text))
    figures = [f for f in active_figures() if f["id"] in present]
    if not figures:
        return ""
    counts, rows = {}, []
    for fig in figures:
        counts[fig["kind"]] = counts.get(fig["kind"], 0) + 1
        title = TITLES_TL.get(fig["id"], fig["title"]) if LANG["code"] == "tl" else fig["title"]
        rows.append(f'<li><a class="toc-map" href="#fig-{fig["id"]}">'
                    f'<span class="toc-kind">{kind_label(fig["kind"])} {counts[fig["kind"]]}.</span> '
                    f'{title}</a></li>')
    return ('<section class="figure-list">'
            f'<h1 class="backmatter" id="figures">{T("Maps and Documents", "Mga Mapa at Dokumento")}</h1>'
            f'<ol class="toc-maps">{"".join(rows)}</ol></section>')


def cover(doc):
    return f"""
<section class="cover">
  <div class="cover-frame">
    {f'<p class="cover-kicker">{doc["kicker"]}</p>' if doc['kicker'] else ''}
    <div class="cover-rule"></div>
    <h1 class="cover-title">{doc['title']}</h1>
    <p class="cover-subtitle">{doc['subtitle']}</p>
    <p class="cover-author">{doc.get('author', '')}</p>
    <div class="cover-photo" style="background-image: url({IMAGES['cover']['file']})"></div>
    <p class="cover-tagline">{doc['tagline']}</p>
    <p class="cover-languages">{doc['languages']}</p>
  </div>
</section>"""


# Striking facts for the back cover of the history, one per period
# (sources in the chapters and in book/research/).
BACK_COVER_FACTS = [
    ("Some 6,600 to 4,700 years ago, the lake in front of Biñan was an arm of the sea, "
     "connected to Manila Bay.",
     "Mga 6,600 hanggang 4,700 taon na ang nakararaan, bahagi ng dagat na karugtong ng "
     "Look ng Maynila ang lawa sa harap ng Biñan."),
    ("Early humans were already butchering a rhinoceros on Luzon 709,000 years ago – and "
     "the island had its own human species, <em>Homo luzonensis</em>.",
     "May mga sinaunang tao nang kumakatay ng rhinoceros sa Luzon 709,000 taon na ang "
     "nakararaan – at may sariling uri pa ng tao ang pulo, ang <em>Homo luzonensis</em>."),
    ("The oldest written document of the Philippines, from the year 900, was found on the "
     "shore of the same lake.",
     "Ang pinakamatandang nakasulat na dokumento ng Pilipinas, mula pa noong taong 900, "
     "ay natagpuan sa baybayin ng lawa ring ito."),
    ("In 1870 a nine-year-old José Rizal came to Biñan to study; years later he put the "
     "town’s church in his novel <em>Noli me tangere</em>.",
     "Noong 1870, dumating sa Biñan ang siyam na taong gulang na si José Rizal upang mag-aral; "
     "makalipas ang ilang taon, isinama niya ang simbahan ng bayan sa nobelang "
     "<em>Noli me tangere</em>."),
    ("In January 1900, American troops captured Biñan twice in five days.",
     "Noong Enero 1900, dalawang beses na nasakop ng mga sundalong Amerikano ang Biñan sa "
     "loob ng limang araw."),
    ("On 3 February 1945, Filipino guerrillas freed Biñan in a single day – before any "
     "American soldier had arrived.",
     "Noong Pebrero 3, 1945, pinalaya ng mga gerilyang Pilipino ang Biñan sa loob ng isang "
     "araw – bago pa dumating ang sinumang sundalong Amerikano."),
    ("In 2024 Biñan became the most populous city of Laguna, with 584,479 people – "
     "twenty-eight times more than in 1948.",
     "Noong 2024, naging pinakamataong lungsod ng Laguna ang Biñan, na may 584,479 katao – "
     "dalawampu’t walong ulit na higit kaysa noong 1948."),
]


def back_cover(doc):
    if doc["kind"] != "history":
        return ""
    items = "".join(f'<li><p class="bc-en">{T(en, tl)}</p></li>' for en, tl in BACK_COVER_FACTS)
    return f"""
<section class="back-cover">
  <div class="cover-frame back-frame">
    <p class="bc-title">{T("Did you know?", "Alam mo ba?")}</p>
    <div class="cover-rule"></div>
    <ul class="bc-facts">{items}</ul>
    <p class="bc-footer">{doc['title']} · {doc.get('author', '')}</p>
  </div>
</section>"""


def colophon(doc):
    """The short notice on the back of the title page."""
    other = T("A Tagalog edition, <em>Kasaysayan ng Lungsod ng Biñan</em>, has the same text "
              "and the same numbered sections, and a companion, <em>Reading the Tagalog Text</em>, "
              "explains its vocabulary and grammar.",
              "May edisyong Ingles, ang <em>The Book of Biñan</em>, na pareho ang teksto at ang "
              "mga bahaging may bilang, at isang kasamang aklat, ang <em>Reading the Tagalog "
              "Text</em>, na nagpapaliwanag sa talasalitaan at balarila nito.")
    return f"""
<section class="colophon">
  <p class="colophon-title">{doc['title']}</p>
  <p>© {doc.get('author', '')}, 2026</p>
  <p>{doc['subtitle']}</p>
  <p>{T("Draft edition, November 2026.", "Borador na edisyon, Nobyembre 2026.")}</p>
  <p>{other}</p>
  <p class="colophon-head">{T("An open book", "Isang bukas na aklat")}</p>
  <p>{T("This book is an open project: the text, its sources and every "
        "correction are public at https://github.com/smichea/binan. Corrections and memories from the "
        "people of Biñan are welcome.",
        "Bahagi ng <em>Ang Aklat ng Biñan</em>, isang bukas na proyekto: nasa "
        "https://github.com/smichea/binan ang teksto, ang mga sanggunian at ang bawat pagwawasto. Malugod "
        "na tinatanggap ang mga pagwawasto at alaala ng mga taga-Biñan.")}</p>
  <p>{T("The electronic edition of this book is free.",
        "Libre ang elektronikong edisyon ng aklat na ito.")}</p>
  <p>{T("Sources, picture credits, the list of maps and documents and the index are at the end "
        "of the book.",
        "Nasa dulo ng aklat ang mga sanggunian, ang mga kredito ng mga larawan, ang talaan ng "
        "mga mapa at dokumento at ang talatuntunan.")}</p>
</section>"""


def credits_section(doc):
    photos = "".join(f"<p>{image['credit']}</p>" for image in IMAGES.values())
    figures = "".join(
        f"<p><span class=\"cr-title\">{T(fig['title'], TITLES_TL.get(fig['id'], fig['title']))}</span>: {fig['credit']}</p>"
        for fig in active_figures() if fig.get("credit") and fig["kind"] != "Illustration")
    maps = T(
        """Maps and charts were drawn for this book (book/maps/make_maps.py) from: the
     Philippine Standard Geographic Code boundaries (PSA, via OCHA and
     github.com/faeldon/philippines-json-maps, MIT licence; barangay boundaries are
     approximate); Natural Earth land (public domain, via world-atlas, ISC licence);
     plate boundaries of P. Bird (2003), PB2002 (ODC-By); the GEM Global Active Faults
     Database (Styron &amp; Pagani 2020, CC BY-SA 4.0); NASA POWER and PAGASA climate data;
     the Philippine censuses (NSO and PSA); and, for roads, railway and places of Part Two,
     OpenStreetMap (© OpenStreetMap contributors, ODbL 1.0, via Nominatim and OSRM) and
     Wikidata (CC0). The maps that contain GEM fault traces are released under the same
     CC BY-SA 4.0 licence.""",
        """Iginuhit para sa aklat na ito ang mga mapa at tsart (book/maps/make_maps.py) mula sa:
     mga hangganan ng Philippine Standard Geographic Code (PSA, sa pamamagitan ng OCHA at
     github.com/faeldon/philippines-json-maps, lisensiyang MIT; tinatayang lamang ang mga
     hangganan ng mga barangay); ang lupain ng Natural Earth (pampublikong pag-aari, sa
     pamamagitan ng world-atlas, lisensiyang ISC); ang mga hangganan ng plate ni P. Bird
     (2003), PB2002 (ODC-By); ang GEM Global Active Faults Database (Styron &amp; Pagani 2020,
     CC BY-SA 4.0); ang mga datos ng klima ng NASA POWER at PAGASA; ang mga senso ng Pilipinas
     (NSO at PSA); at, para sa mga daan, riles at lugar ng Ikalawang Bahagi, ang OpenStreetMap
     (© mga kontributor ng OpenStreetMap, ODbL 1.0, sa pamamagitan ng Nominatim at OSRM) at
     ang Wikidata (CC0). Inilalathala sa ilalim ng parehong lisensiyang CC BY-SA 4.0 ang mga
     mapang may mga fault ng GEM.""")
    illustrations = T(
        """The pictures labelled “Illustration” are artist’s impressions created for this book
     with an AI image generator (Figma). They show what the sources describe, for periods
     or scenes of which no picture exists; they are not historical documents. They are
     listed for review by people of Biñan in book/illustrations_revue.md.""",
        """Ang mga larawang may tatak na “Guhit” ay guhit-paglalarawan na ginawa para sa aklat na
     ito sa tulong ng isang AI (Figma). Ipinakikita nila ang inilalarawan ng mga sanggunian,
     para sa mga panahon o eksenang walang natitirang larawan; hindi sila makasaysayang
     dokumento. Nakalista sila sa book/illustrations_revue.md upang masuri ng mga taga-Biñan.""")
    return f"""
<section class="credits">
  <h1 class="backmatter" id="credits">{T("Credits", "Mga Kredito")}</h1>
  <p class="colophon-head">{T("Cover and photographs", "Pabalat at mga larawan")}</p>
  {photos}
  {figures}
  <p class="colophon-head">{T("Maps and charts", "Mga mapa at tsart")}</p>
  <p>{maps}</p>
  <p class="colophon-head">{T("Illustrations", "Mga guhit")}</p>
  <p>{illustrations}</p>
  <p class="colophon-head">{T("Typography", "Tipograpiya")}</p>
  <p>{T("Set in Source Serif 4 and Cormorant Garamond (SIL Open Font License).",
        "Nakaayos sa Source Serif 4 at Cormorant Garamond (SIL Open Font License).")}</p>
</section>"""


def index_shell():
    return ('<section class="index-section">'
            f'<h1 class="backmatter" id="index">{T("Index", "Talatuntunan")}</h1>'
            '<p class="bm-intro">'
            + T("People, places and institutions, with the pages where they appear.",
                "Mga tao, lugar at institusyon, at ang mga pahinang kinaroroonan nila.")
            + f'</p>{INDEX_PLACEHOLDER}</section>')


EDITION_CSS = {
    "tl": """body.history { string-set: book-title "Kasaysayan ng Lungsod ng Biñan"; }
h1.chapter::before { content: "Kabanata " counter(chapter, upper-roman); }
h1.appendix::before { content: "Apendise " counter(appendix, upper-alpha); }""",
}


def title_page(doc):
    """The first page of the printed interior (the cover is printed apart)."""
    return f"""
<section class="title-page">
  {f'<p class="tp-kicker">{doc["kicker"]}</p>' if doc['kicker'] else ''}
  <h1 class="tp-title">{doc['title']}</h1>
  <p class="tp-subtitle">{doc['subtitle']}</p>
  <p class="tp-author">{doc.get('author', '')}</p>
  <p class="tp-tagline">{doc['tagline']}</p>
</section>"""


def page_html(doc, front, body, back, body_class, printing=False):
    style = EDITION_CSS.get(LANG["code"], "") if body_class == "history" else ""
    return f"""<!doctype html>
<html lang="{LANG['code']}"><head><meta charset="utf-8">
<title>{doc['title']}</title>
<meta name="author" content="{doc.get('author', '')}">
<link rel="stylesheet" href="book.css">
<style>{style}</style>
</head><body class="{body_class}">
{title_page(doc) if printing else cover(doc)}
{colophon(doc) if body_class == "history" else ""}
{table_of_contents(front + body + back)}
<div class="front-matter">{front}</div>
<main>{body}</main>
<div class="back-matter">{back}</div>
{"" if printing else back_cover(doc)}
</body></html>"""


def build(doc, printing=False):
    """Write the PDF of one document. printing=True writes, for a printer, the
    interior (title page instead of the cover, even number of pages) and the
    cover as one spread: back, spine and front, with bleed."""
    LANG["code"] = doc.get("lang", "en")
    markdown = doc["source"].read_text(encoding="utf-8")
    if doc["kind"] != "history":
        CURRENT["record"] = False
        body = to_html(prepare_explanation(markdown), "commonmark_x")
        page = page_html(doc, "", body, "", "companion")
        HTML(string=page, base_url=str(LAYOUT) + "/").write_pdf(doc["output"])
        print(f"wrote {doc['output'].relative_to(BOOK.parent)}")
        return

    SOURCES.clear()
    FIGURE_NUMBERS.clear()
    CURRENT["record"] = True
    front = ""
    if doc.get("front") and doc["front"].exists():
        front = to_html(prepare_text(doc["front"].read_text(encoding="utf-8"), "front"))
    body = to_html(prepare_text(markdown, "chapter"))
    appendices = ""
    if doc.get("back") and doc["back"].exists():
        appendices = to_html(prepare_text(doc["back"].read_text(encoding="utf-8"), "appendix"))
    CURRENT["record"] = False

    # Index: mark the terms, lay the book out once to learn their pages, then
    # print the index (at the very end, so no earlier page moves).
    indexed, anchors = mark_index(front + "<!--MAIN-->" + body)
    front, body = indexed.split("<!--MAIN-->")
    back = (appendices + sources_section() + credits_section(doc)
            + figures_section(body) + index_shell())
    page = page_html(doc, front, body, back, "history", printing)
    base = str(LAYOUT) + "/"
    if anchors:
        first = HTML(string=page.replace(INDEX_PLACEHOLDER, ""), base_url=base).render()
        pages = {}
        for number, laid_out in enumerate(first.pages, start=1):
            for anchor in laid_out.anchors:
                if anchor.startswith("ix"):
                    pages.setdefault(anchor, number)
        page = page.replace(INDEX_PLACEHOLDER, index_section(anchors, pages))
    if not printing:
        HTML(string=page, base_url=base).write_pdf(doc["output"])
        print(f"wrote {doc['output'].relative_to(BOOK.parent)}")
        return
    PRINT_DIR.mkdir(exist_ok=True)
    stem = doc["output"].stem
    interior = HTML(string=page, base_url=base).render()
    count = len(interior.pages)
    path = PRINT_DIR / f"{stem}_interior.pdf"
    interior.write_pdf(path)
    if count % 2:  # a bound book has an even number of pages
        from pypdf import PdfReader, PdfWriter
        writer = PdfWriter(clone_from=PdfReader(path))
        writer.add_blank_page()
        writer.write(path)
        count += 1
    print(f"wrote {path.relative_to(BOOK.parent)} ({count} pages)")
    spine = round(count / 2 * PAPER_MM + 0.5, 1)
    path = PRINT_DIR / f"{stem}_cover.pdf"
    HTML(string=cover_spread(doc, spine), base_url=base).write_pdf(path)
    print(f"wrote {path.relative_to(BOOK.parent)} (spine {spine} mm, bleed {BLEED_MM} mm)")


# Printing: the files for a printer go to book/word/print. The spine width
# follows the page count; PAPER_MM is the thickness of one sheet (80 g/m²
# offset paper: about 0.1 mm) -- ask the printer for the real value.
PRINT_DIR = BOOK / "word" / "print"
PAPER_MM = 0.1
BLEED_MM = 3


def cover_spread(doc, spine, panel_w=152.4, panel_h=228.6):
    """The printed cover: back cover, spine and front cover on one sheet.
    The panels are 6 x 9 in, unless another trim size is given (in mm)."""
    width = 2 * panel_w + spine + 2 * BLEED_MM
    height = panel_h + 2 * BLEED_MM
    spine_text = f"{doc['title']} · {doc.get('author', '')}"
    return f"""<!doctype html>
<html lang="{LANG['code']}"><head><meta charset="utf-8"><title>{doc['title']} – cover</title>
<link rel="stylesheet" href="book.css">
<style>
@page {{ size: {width}mm {height}mm; margin: 0; @top-left {{ content: none; }} @top-right {{ content: none; }} }}
body {{ margin: 0; }}
.spread {{ position: relative; width: {width}mm; height: {height}mm; background: var(--accent); }}
.panel {{ position: absolute; top: {BLEED_MM}mm; width: {panel_w}mm; height: {panel_h}mm; }}
.panel .cover, .panel .back-cover {{ page: auto; break-before: auto; break-after: auto; height: {panel_h}mm; }}
.spine {{ position: absolute; top: {BLEED_MM}mm; left: {BLEED_MM + panel_w}mm; width: {spine}mm; height: {panel_h}mm; }}
.spine p {{ position: absolute; left: 50%; top: 50%; width: 8in; margin: 0; text-align: center;
           transform: translate(-50%, -50%) rotate(90deg); color: var(--gold);
           font-family: "Cormorant Garamond"; font-weight: 600; font-size: {min(11, spine * 1.6):.1f}pt;
           letter-spacing: 0.08em; white-space: nowrap; }}
</style></head><body class="history">
<div class="spread">
  <div class="panel" style="left: {BLEED_MM}mm">{back_cover(doc)}</div>
  <div class="spine"><p>{spine_text}</p></div>
  <div class="panel" style="left: {BLEED_MM + panel_w + spine}mm">{cover(doc)}</div>
</div>
</body></html>"""


if __name__ == "__main__":
    if "--print" in sys.argv:  # print files of the English edition
        build(DOCUMENTS[0], printing=True)
    else:
        for document in DOCUMENTS:
            build(document)
