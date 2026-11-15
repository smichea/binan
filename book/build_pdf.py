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
    dict(chapter=1, after=1, kind="Photo", id="iss-lake", src="images/geography/iss-manila-laguna-2016.jpg",
         title="Manila Bay and Laguna de Bay from space",
         caption="Metro Manila between Manila Bay (left) and Laguna de Bay (right), photographed "
                 "from the International Space Station on 6 May 2016. The Pasig River links the "
                 "lake to the sea.",
         tagalog="Ang Look ng Maynila at ang Lawa ng Laguna mula sa kalawakan.",
         credit="Astronaut photograph ISS047-E-99713. NASA, ISS Crew Earth Observations Facility. Public domain."),
    dict(chapter=1, after=2, kind="Map", id="tectonics", src="maps/tectonics.svg",
         title="The Philippines between two plates",
         caption="The Philippines between the Philippine Sea Plate and the Sunda Plate: the "
                 "subduction trenches, the active faults and some of the active volcanoes. In red, "
                 "the West Valley Fault, which crosses Biñan.",
         tagalog="Ang Pilipinas sa pagitan ng dalawang tectonic plate.",
         credit="Plate boundaries: P. Bird (2003), PB2002, via H. Ahlenius (ODC-By). Faults: GEM "
                "Global Active Faults Database (Styron & Pagani 2020), CC BY-SA."),
    dict(chapter=1, after=3, kind="Photo", id="taal-ash", src="images/geography/iss-taal-ash-2020.jpg",
         title="Taal’s ash drifting toward Laguna de Bay, 2020",
         caption="Ash from Taal Volcano (top left) drifting over Batangas, Cavite and Laguna "
                 "toward Laguna de Bay (top right), seen from the International Space Station on "
                 "22 January 2020.",
         tagalog="Ang abo ng Taal patungo sa Lawa ng Laguna, 2020.",
         credit="Expedition 61 crew. NASA, ISS Crew Earth Observations Facility. Public domain."),
    dict(chapter=1, after=3, kind="Photo", id="makiling", src="images/geography/mount-makiling.jpg",
         title="Mount Makiling across the lake",
         caption="Mount Makiling, the extinct volcano of the legend of Maria Makiling, seen "
                 "across Laguna de Bay with fish pens in the foreground.",
         tagalog="Ang Bundok Makiling mula sa lawa.",
         credit="Photo: Zian Colt, 2021, CC BY-SA 4.0, via Wikimedia Commons."),
    dict(chapter=1, after=4, kind="Map", id="westbay", src="maps/westbay.svg",
         title="Fire, fault and lake around Biñan",
         caption="Biñan between the West Valley Fault and the lake: the approximate outline of "
                 "the Laguna caldera, the volcanoes of the Macolod Corridor, and the former link "
                 "with Manila Bay, when the West Bay was an arm of the sea (6,600–4,700 years ago; "
                 "schematic).",
         tagalog="Ang apoy, ang fault at ang lawa sa paligid ng Biñán.",
         credit="Fault trace: GEM Global Active Faults Database, CC BY-SA. Sea link after "
                "Jaraula et al. 2014."),
    dict(chapter=1, after=4, kind="Photo", id="fault-plane", src="images/geography/wvf-fault-plane-carmona.jpg",
         title="The West Valley Fault at Carmona",
         caption="The scarp of the West Valley Fault, the long grassy slope on the right, in "
                 "Barangay Mabuhay, Carmona, a few kilometres from Biñan.",
         tagalog="Ang West Valley Fault sa Carmona.",
         credit="Photo: Ervin Malicdem, 2017, CC BY-SA 4.0, via Wikimedia Commons."),
    dict(chapter=1, after=6, kind="Illustration", id="lagoon", src="images/illustrations/salt-lagoon-6000-years.jpg",
         title="The salt lagoon, 6,000 years ago",
         caption="An impression of the shore of Biñan about 6,000 years ago, when the western part "
                 "of the lake was an arm of the sea: mangroves, tidal flats and shellfish gatherers.",
         tagalog="Ang maalat na laguna, 6,000 taon na ang nakararaan.", credit="Illustration created for this book (AI-generated, Figma). Artist’s impression."),
    dict(chapter=1, after=7, kind="Chart", id="climate", src="maps/climate.svg",
         title="The climate of Biñan",
         caption="Left: the monsoon year – normal monthly rainfall at Sangley Point, the nearest "
                 "PAGASA station; the wet months are in blue. Right: mean yearly temperature for "
                 "Biñan’s coordinates, 1981–2024, from NASA’s reconstruction, with its trend. The "
                 "model averages a wide area, so its values are lower than those measured on the "
                 "lakeshore; it is the trend that matters.",
         tagalog="Ang klima ng Biñán.",
         credit="PAGASA Climatological Normals; NASA POWER (MERRA-2)."),
    dict(chapter=2, after=1, kind="Map", id="migrations", src="maps/migrations.svg",
         title="Early human journeys to Island Southeast Asia",
         caption="The long journeys of early humans and of our own species from Africa to "
                 "Island Southeast Asia, with some of the oldest dated sites, and the later "
                 "voyages of Austronesian seafarers from Taiwan. Routes are simplified.",
         tagalog="Ang mahabang paglalakbay ng mga sinaunang tao patungong Timog-Silangang Asya.",
         credit="Land: Natural Earth (public domain), via world-atlas. Dates from the sources "
                "cited in the text."),
    dict(chapter=2, after=2, kind="Photo", id="rhino", src="images/prehistory/kalinga-rhinoceros.jpg",
         title="The butchered rhinoceros of Kalinga",
         caption="Bones of the extinct Philippine rhinoceros found at Rizal, Kalinga, about 709,000 "
                 "years old: the cut marks left by stone tools are the oldest trace of humans in "
                 "the Philippines.",
         tagalog="Ang kinatay na rhinoceros ng Kalinga.",
         credit="Photo: Darwgon0801, CC BY-SA 4.0, via Wikimedia Commons."),
    dict(chapter=2, after=2, kind="Photo", id="luzonensis", src="images/prehistory/homo-luzonensis.jpg",
         title="Homo luzonensis",
         caption="Teeth and foot bones of Homo luzonensis from Callao Cave, Cagayan: a human "
                 "species found only on Luzon.",
         tagalog="Ang Homo luzonensis.",
         credit="Diagram: Homo schollensis, CC BY-SA 4.0, via Wikimedia Commons (after Détroit et al. 2019)."),
    dict(chapter=2, after=2, kind="Illustration", id="ill-luzonensis", src="images/illustrations/homo-luzonensis.jpg",
         title="Homo luzonensis, an impression",
         caption="An impression of Homo luzonensis at a cave mouth in Cagayan. Only teeth and a "
                 "few small bones have been found: its true appearance is unknown.",
         tagalog="Ang Homo luzonensis, ayon sa guhit ng isang artista.", credit="Illustration created for this book (AI-generated, Figma). Artist’s impression."),
    dict(chapter=2, after=3, kind="Photo", id="tabon", src="images/prehistory/tabon-cave.jpg",
         title="Tabon Cave, Palawan",
         caption="Tabon Cave in Palawan, where the oldest known bones of modern humans in the "
                 "Philippines were found.",
         tagalog="Ang Kuweba ng Tabon sa Palawan.",
         credit="Photo: Jimaggro, 2014, CC BY-SA 4.0, via Wikimedia Commons."),
    dict(chapter=2, after=5, kind="Illustration", id="ill-austronesians", src="images/illustrations/austronesian-arrival.jpg",
         title="Austronesian seafarers reach Luzon",
         caption="An impression of Austronesian seafarers landing in northern Luzon about 4,000 "
                 "years ago, with their outrigger canoes, red-slipped pots, plants and animals.",
         tagalog="Ang pagdating ng mga manlalayag na Austronesyo sa Luzon.", credit="Illustration created for this book (AI-generated, Figma). Artist’s impression."),
    dict(chapter=2, after=5, kind="Map", id="prehistoric-sites", src="maps/prehistoric-sites.svg",
         title="Prehistoric sites of the Philippines",
         caption="Some prehistoric sites of the Philippines, and, in detail, the places around "
                 "Laguna de Bay where ancient traces have been found. None has yet been found in "
                 "Biñan itself.",
         tagalog="Mga sinaunang pook sa Pilipinas at sa paligid ng Lawa ng Laguna.",
         credit="Boundaries: PSA / OCHA via philippines-json-maps (MIT)."),
    dict(chapter=2, after=6, kind="Photo", id="petroglyphs", src="images/prehistory/angono-petroglyphs.jpg",
         title="The Angono–Binangonan Petroglyphs",
         caption="Figures carved into a wall of volcanic tuff at Binangonan, Rizal, on the "
                 "northern shore of the lake: among the oldest known artworks of the Philippines.",
         tagalog="Ang Angono–Binangonan Petroglyphs.",
         credit="Photo: Raffy-ish, CC BY-SA 3.0, via Wikimedia Commons."),
    dict(chapter=2, after=6, kind="Photo", id="manunggul", src="images/prehistory/manunggul-jar.jpg",
         title="The Manunggul Jar",
         caption="The Manunggul Jar (Palawan, about 890–710 BCE): on its lid, two figures in a "
                 "boat carry a soul to the afterlife. National Museum of the Philippines.",
         tagalog="Ang Banga ng Manunggul.",
         credit="Photo: Gary Todd, CC0, via Wikimedia Commons."),
    dict(chapter=2, after=8, kind="Illustration", id="ill-village", src="images/illustrations/lakeshore-village.jpg",
         title="A lakeshore village, about 1,000 years ago",
         caption="An impression of a village on Laguna de Bay about a thousand years ago: houses on "
                 "stilts, outrigger canoes, fishing nets, taro and earthen jars, with Mount "
                 "Makiling across the water.",
         tagalog="Isang nayon sa baybayin ng lawa, mga 1,000 taon na ang nakararaan.", credit="Illustration created for this book (AI-generated, Figma). Artist’s impression."),
    dict(chapter=2, after=8, kind="Photo", id="balangay", src="images/prehistory/balangay-replica.jpg",
         title="A balangay sails again",
         caption="A modern replica of a balangay, the plank boat of ancient Butuan, sailing on "
                 "Manila Bay. The word barangay comes from this boat.",
         tagalog="Isang balangay sa Look ng Maynila.",
         credit="Photo: Fung360, CC BY-SA 4.0, via Wikimedia Commons."),
    dict(chapter=3, after=1, kind="Photo", id="copperplate",
         src="images/archives/copperplate-c900.jpg",
         title="The Laguna Copperplate Inscription, c. 900",
         caption="The Laguna Copperplate Inscription (c. 900 CE), the oldest known written "
                 "document of the Philippines, found in 1989 near Lumban on Laguna de Bay. "
                 "It shows that the lake shore was home to organized, literate communities "
                 "centuries before the Spaniards arrived.",
         tagalog="Ang Inskripsiyon sa Binatbat na Tanso ng Laguna, mga 900.",
         credit="Photo: Gary Todd, CC0, via Wikimedia Commons. National Museum of the Philippines."),
    dict(chapter=4, after=1, kind="Map", id="territory", src="maps/territory.svg",
         title="The territory of Biñan, 1688–1907",
         caption="How Biñan’s territory took shape: separated from Tabuco in 1688, it lost "
                 "its barrio of Santa Rosa in 1791–92. Under American rule, San Pedro "
                 "(1903–1907) and Muntinlupa (1903–1905) were briefly joined to it. "
                 "Present-day boundaries.",
         tagalog="Ang paghubog ng teritoryo ng Biñán, 1688–1907."),
    dict(chapter=4, after=2, kind="Map", id="murillo", src="images/archives/murillo-velarde-1734.jpg",
         title="“Viñan” on the Murillo Velarde map, 1734",
         caption="The south-western shore of Laguna de Bay on Pedro Murillo Velarde’s 1734 map of the Philippines, the "
                 "first scientific map of the islands. “Viñan” appears at the left edge, on the shore "
                 "near Tunasan and Calamba.",
         tagalog="Ang “Viñan” sa mapa ni Murillo Velarde, 1734.",
         credit="Pedro Murillo Velarde, Carta hydrographica y chorographica de las Yslas "
                "Filipinas (Manila, 1734). Library of Congress, Geography and Map Division. "
                "Public domain."),
    dict(chapter=4, after=2, kind="Map", id="manuscript", src="images/archives/manuscript-laguna-18th-c.jpg",
         title="“Biñang” on an 18th-century Spanish map",
         caption="Detail of an 18th-century Spanish manuscript map of the surroundings of "
                 "Laguna de Bay: the town of “Biñang” with its church (red square), the river "
                 "mouth (“Bª de Biñang”), Malabag (today Malaban), Sinalsan, Landayan, San "
                 "Pedro Tunasan and Santa Rosa.",
         tagalog="Ang “Biñang” sa isang mapang Espanyol noong ika-18 siglo.",
         credit="Fuente: Archivo Cartográfico de Estudios Geográficos del Centro Geográfico "
                "del Ejército / Biblioteca Virtual de Defensa. CC BY 4.0."),
    dict(chapter=4, after=3, kind="Map", id="map1897", src="images/archives/cavite-region-map-1897.jpg",
         title="“Biñang” on a Spanish military map, 1897",
         caption="“Biñang” and its barrios on the Spanish military map of the province of Cavite "
                 "and the neighbouring region, drawn in 1897 during the Philippine Revolution.",
         tagalog="Ang “Biñang” sa isang mapang militar ng Espanya, 1897.",
         credit="Provincia de Cavite y región limítrofe, 1897. Fuente: Archivo Histórico de la "
                "Armada Juan Sebastián de Elcano / Biblioteca Virtual de Defensa. CC BY 4.0."),
    dict(chapter=4, after=3, kind="Photo", id="school", src="images/archives/justiniano-school-site-pre1918.jpg",
         title="Where Rizal went to school in Biñan",
         caption="“Where Rizal went to school in Biñan: site of Master Justiniano’s house, which "
                 "burned down many years ago.” The young José Rizal studied here in 1869–70; the "
                 "site is now the School of Dr. José Rizal museum.",
         tagalog="Ang dating paaralan ni Maestro Justiniano, kung saan nag-aral si Rizal.",
         credit="Austin Craig, Rizal’s Own Story of His Life (Manila, 1918). Internet Archive. "
                "Public domain."),
    dict(chapter=4, after=3, kind="Photo", id="tablet", small=True,
         src="images/archives/jose-alberto-tablet-pre1913.jpg",
         title="Memorial tablet of José Alberto",
         caption="Memorial tablet of José Alberto, Rizal’s uncle and owner of the Alberto house, "
                 "who died on 29 October 1875, in the church of Biñan.",
         tagalog="Ang lapida ni José Alberto sa simbahan ng Biñán.",
         credit="Austin Craig, Lineage, Life and Labors of José Rizal (Manila, 1913). "
                "Internet Archive. Public domain."),
    dict(chapter=5, after=1, kind="Map", id="atlas", src="images/archives/atlas-filipinas-1899.jpg",
         title="Biñan in the Atlas de Filipinas, 1899–1900",
         caption="“Biñang” in the Atlas de Filipinas of the Manila Observatory (Map No. 9, "
                 "U.S. Coast and Geodetic Survey edition, 1900), on the road from Manila to "
                 "Calamba, as the Americans found it.",
         tagalog="Ang Biñán sa Atlas de Filipinas, 1899–1900.",
         credit="Observatorio de Manila / U.S. Coast and Geodetic Survey, Atlas of the "
                "Philippine Islands (Washington, 1900), Map No. 9. Internet Archive. Public domain."),
    dict(chapter=5, after=2, kind="Photo", id="monument", src="images/archives/rizal-monument-plaza-pre1918.jpg",
         title="The Rizal monument on the plaza, 1918",
         caption="The Rizal monument on the town plaza, newly built, with the Alberto house – the "
                 "home of Rizal’s mother’s family – behind it.",
         tagalog="Ang bantayog ni Rizal sa plasa ng Biñán, mga 1918.",
         credit="Austin Craig, Rizal’s Own Story of His Life (Manila, 1918). Internet Archive. "
                "Public domain."),
    dict(chapter=5, after=3, kind="Photo", id="cinema", src="images/archives/alberto-house-cinematograph-pre1918.jpg",
         title="The Alberto house as a cinema, 1910s",
         caption="“Rizal’s uncle’s home in Biñan”: in the 1910s the ground floor of the Alberto "
                 "house served as the town’s cinematograph – a sign of the new American-era "
                 "leisure, from movies to baseball.",
         tagalog="Ang bahay ng mga Alberto noong naging sinehan, dekada 1910.",
         credit="Austin Craig, Rizal’s Own Story of His Life (Manila, 1918). Internet Archive. "
                "Public domain."),
    dict(chapter=5, after=4, kind="Photo", id="southroad", rights="pending",
         src="images/archives/pendleton-south-road-1934.jpg",
         title="Carabaos on the Manila South Road, 1934",
         caption="Carabaos on the Manila South Road between Muntinlupa and Biñan, looking toward "
                 "Biñan, 6 December 1934. Rain trees line the national road; rice fields and a "
                 "telegraph line run alongside.",
         tagalog="Mga kalabaw sa Manila South Road patungong Biñán, 1934.",
         credit="Robert L. Pendleton, 1934. American Geographical Society Library, University "
                "of Wisconsin–Milwaukee Libraries (negative 409.11)."),
    dict(chapter=6, after=1, kind="Map", id="liberation", src="maps/liberation.svg",
         title="The liberation of Biñan, 1945",
         caption="The liberation of Biñan, 3 February 1945: the Hunters ROTC guerrillas "
                 "converged on the town centre from Carmona, San Pedro and Santa Rosa. "
                 "Arrows are schematic.",
         tagalog="Ang paglaya ng Biñán, Pebrero 3, 1945."),
    dict(chapter=7, after=1, kind="Photo", id="bridge", src="images/archives/binan-bridge-1951.jpg",
         title="The new Biñan bridge, c. 1951",
         caption="“Biñan Bridge completed”, about 1951: one of the bridges of the national "
                 "road rebuilt after the war under the Philippine–American public roads "
                 "rehabilitation programme.",
         tagalog="Ang bagong tulay ng Biñán, mga 1951.",
         credit="U.S. National Archives, Records of the Bureau of Public Roads (RG 30), "
                "30-N-51-4622. Public domain."),
    dict(chapter=7, after=1, kind="Map", id="ams1956", src="images/archives/ams-san-pedro-1956.jpg",
         title="The town of Biñan on a 1956 survey map",
         caption="The town of Biñan on the 1956 U.S. Army Map Service sheet “San Pedro” "
                 "(1:50,000): the street grid of the Poblacion and the neighbouring barrios "
                 "of Canlalay, Malaban, De La Paz and San Antonio, between the national road "
                 "and the lake.",
         tagalog="Ang bayan ng Biñán sa mapa ng U.S. Army, 1956.",
         credit="U.S. Army Map Service, San Pedro (sheet 3263 III), 1:50,000, 1956. "
                "Courtesy of the University of Texas Libraries. Public domain."),
    dict(chapter=7, after=2, kind="Map", id="barangays", src="maps/barangays.svg",
         title="The barangays of Biñan today",
         caption="Biñan today and its 24 barangays. The Poblacion keeps the Spanish-era "
                 "heart of the town; the Laguna Technopark (Malamig) and the Laguna "
                 "International Industrial Park (Mampalasan) mark its industrial growth "
                 "since 1990.",
         tagalog="Ang 24 na barangay ng Biñán ngayon."),
    dict(chapter=7, after=2, kind="Photo", id="technopark", src="images/modern/technopark-sign-2018.jpg",
         title="The Laguna Technopark",
         caption="The entrance sign of the Laguna Technopark, opened in 1990 on land in Biñan "
                 "(Malamig) and Santa Rosa (Don Jose), listing some of its manufacturers.",
         tagalog="Ang Laguna Technopark.", credit="Photo: Judgefloro, 2018, CC0, via Wikimedia Commons."),
    dict(chapter=7, after=4, kind="Photo", id="cityhall", src="images/modern/city-hall-2018.jpg",
         title="Biñan City Hall",
         caption="The Biñan City Hall complex, seat of the city government since Biñan became "
                 "a city in 2010.",
         tagalog="Ang bulwagang panlungsod ng Biñán.", credit="Photo: Judgefloro, 2018, CC0, via Wikimedia Commons."),
    dict(chapter=7, after=5, kind="Photo", id="view2023", src="images/modern/view-from-alberto-mansion-2023.jpg",
         title="The Poblacion from the Alberto Mansion, 2023",
         caption="The heart of Biñan seen from the balcony of the restored Alberto Mansion: the "
                 "plaza, the old municipal hall and the bell tower of San Isidro Labrador.",
         tagalog="Ang Poblacion mula sa Alberto Mansion, 2023.", credit="Photo: Ralff Nestor Nacor, 2023, CC BY-SA 4.0, via Wikimedia Commons."),
    dict(chapter=7, after=6, kind="Photo", id="sentro", src="images/modern/sentrong-pangkultura-2023.jpg",
         title="The Sentrong Pangkultura ng Biñan",
         caption="The Spanish-era municipal hall, restored as the Sentrong Pangkultura ng Biñan "
                 "(Biñan Center for Culture and the Arts) under the Balik-Biñan programme.",
         tagalog="Ang Sentrong Pangkultura ng Biñán.", credit="Photo: Ralff Nestor Nacor, 2023, CC BY-SA 4.0, via Wikimedia Commons."),
    dict(chapter=8, after=1, kind="Photo", id="marketstreet", src="images/modern/market-street-2013.jpg",
         title="Around the public market, 2013",
         caption="Shops and shoppers around the public market of Biñan on a rainy morning.",
         tagalog="Sa paligid ng pamilihang bayan, 2013.", credit="Photo: Ramon FVelasquez, 2013, CC BY-SA 3.0, via Wikimedia Commons."),
    dict(chapter=8, after=1, kind="Photo", id="fruit", src="images/modern/market-fruit-stall-2013.jpg",
         title="A fruit stall in the market, 2013",
         caption="A fruit stall at the Biñan market: bananas, citrus and local produce.",
         tagalog="Isang tindahan ng prutas sa pamilihan, 2013.", credit="Photo: Ramon FVelasquez, 2013, CC BY-SA 3.0, via Wikimedia Commons."),
    dict(chapter=8, after=1, kind="Photo", id="pospas", rights="pending", src="images/culture/pospas.jpg",
         title="Pospas", caption="Pospas, Biñan’s yellow rice porridge with chicken, toasted "
                 "garlic, spring onion and pork crackling.",
         tagalog="Pospas, ang dilaw na lugaw ng Biñán.", credit="Experience Biñan, City Government of Biñan (BCHATO), experiencebinan.com."),
    dict(chapter=8, after=2, kind="Photo", id="puto", rights="pending", src="images/culture/puto-binan.jpg",
         title="Puto Biñan", caption="Puto Biñán, topped with grated cheese and salted egg.",
         tagalog="Ang Puto Biñán.", credit="Experience Biñan, City Government of Biñan (BCHATO), experiencebinan.com."),
    dict(chapter=8, after=3, kind="Photo", id="maglalatik", rights="pending", src="images/culture/maglalatik.jpg",
         title="Maglalatik dancers",
         caption="Maglalatik dancers, coconut shells strapped to their bodies: the “Christians” "
                 "in blue and the “Moros” in red.",
         tagalog="Mga mananayaw ng maglalatik.", credit="Experience Biñan, City Government of Biñan (BCHATO), experiencebinan.com."),
    dict(chapter=8, after=3, kind="Photo", id="festival", rights="pending", src="images/culture/puto-latik-festival.jpg",
         title="The Puto Latik Festival",
         caption="Street dancers in the Puto Latik Festival.",
         tagalog="Ang Puto Latik Festival.", credit="Experience Biñan, City Government of Biñan (BCHATO), experiencebinan.com."),
    dict(chapter=8, after=4, kind="Photo", id="arawngbinan", rights="pending", src="images/culture/araw-ng-binan-night.jpg",
         title="Araw ng Biñan by night",
         caption="Araw ng Biñán: the illuminated Sentrong Pangkultura and crowds on the plaza.",
         tagalog="Ang Araw ng Biñán sa gabi.", credit="Experience Biñan, City Government of Biñan (BCHATO), experiencebinan.com."),
    dict(chapter=8, after=4, kind="Photo", id="mass", src="images/modern/church-mass-2013.jpg",
         title="Sunday Mass at San Isidro Labrador",
         caption="A full church for Mass at San Isidro Labrador, the heart of the town’s "
                 "religious life.",
         tagalog="Misa sa simbahan ni San Isidro Labrador.", credit="Photo: Ramon FVelasquez, 2013, CC BY-SA 3.0, via Wikimedia Commons."),
    dict(chapter=8, after=4, kind="Photo", id="flowers", src="images/modern/market-flowers-2013.jpg",
         title="Flower vendors, 2013",
         caption="Flower vendors outside the market, preparing offerings for the church and for "
                 "family celebrations.",
         tagalog="Mga nagtitinda ng bulaklak, 2013.", credit="Photo: Ramon FVelasquez, 2013, CC BY-SA 3.0, via Wikimedia Commons."),
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
    if fig.get("small"):
        css += " small"
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
            here = [f for f in active_figures()
                    if (f["chapter"], f["after"]) == (chapter, english_paragraphs)]
            row = []  # consecutive photos are set side by side, two per row
            for fig in here:
                numbers[fig["kind"]] = numbers.get(fig["kind"], 0) + 1
                html_fig = figure_html(fig, numbers[fig["kind"]])
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
        map_list = ('</section><section class="toc toc-figures">'
                    '<h1 class="toc-title">Maps and documents</h1>'
                    '<p class="toc-sub">Mga mapa at dokumento</p>'
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


# Striking facts for the back cover of the history (sources in book/research/).
BACK_COVER_FACTS = [
    ("Some 6,600 to 4,700 years ago, the lake in front of Biñan was an arm of the sea, "
     "connected to Manila Bay.",
     "Mga 6,600 hanggang 4,700 taon na ang nakararaan, bahagi ng dagat na karugtong ng "
     "Look ng Maynila ang lawa sa harap ng Biñán."),
    ("An active fault, the West Valley Fault, runs beneath Biñan. Its last great rupture "
     "was in 1658.",
     "Isang aktibong fault, ang West Valley Fault, ang dumaraan sa ilalim ng Biñán. Huli "
     "itong lumindol nang malakas noong 1658."),
    ("Early humans were already butchering a rhinoceros on Luzon 709,000 years ago – and "
     "the island had its own human species, <em>Homo luzonensis</em>.",
     "May mga sinaunang tao nang kumakatay ng rhinoceros sa Luzon 709,000 taon na ang "
     "nakararaan – at may sariling uri pa ng tao ang pulo, ang <em>Homo luzonensis</em>."),
    ("The oldest written document of the Philippines, from the year 900, was found on the "
     "shore of the same lake.",
     "Ang pinakamatandang nakasulat na dokumento ng Pilipinas, mula pa noong taong 900, "
     "ay natagpuan sa baybayin ng lawa ring ito."),
    ("The Tagalog word <em>binyag</em>, “to baptize”, came from Brunei – and tradition says Biñan "
     "takes its name from <em>binyagan</em>, “the place of baptism”.",
     "Galing sa Brunei ang salitang binyag – at ayon sa tradisyon, mula sa “binyagan” "
     "ang pangalan ng Biñán."),
    ("Biñan’s bandstand is shaped like a slipper, a tribute to the town’s old slipper "
     "makers.",
     "Hugis-tsinelas ang entablado ng banda sa Biñán, bilang pagpupugay sa mga dating "
     "gumagawa ng tsinelas."),
]


def back_cover(doc):
    if doc["source"].name != "history.md":
        return ""
    items = "".join(
        f'<li><p class="bc-en">{en}</p><p class="bc-tl">{tl}</p></li>'
        for en, tl in BACK_COVER_FACTS)
    return f"""
<section class="back-cover">
  <div class="cover-frame back-frame">
    <p class="bc-title">Did you know?</p>
    <p class="bc-subtitle">Alam mo ba?</p>
    <div class="cover-rule"></div>
    <ul class="bc-facts">{items}</ul>
    <p class="bc-footer">{doc['title']} · {doc['subtitle']}</p>
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
  <p class="colophon-head">Illustrations</p>
  <p>Illustrations marked as impressions were created for this book with an AI image
     generator (Figma) and checked against the sources cited in the text; they are not
     historical documents.</p>
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
{back_cover(doc)}
</body></html>"""
    HTML(string=page, base_url=str(LAYOUT) + "/").write_pdf(doc["output"])
    print(f"wrote {doc['output'].relative_to(BOOK.parent)}")


if __name__ == "__main__":
    for document in DOCUMENTS:
        build(document)
