# Remplacement des sources faibles dans `history.md`

*Note de recherche pour le livre bilingue sur Biñan. Rédigée le 16/11/2026. Je n'ai modifié aucun autre fichier du dépôt.*

**Version analysée.** `book/raw_files/history.md` tel qu'il était le 16/11/2026 vers 14 h UTC (813 lignes). Le fichier a changé pendant mon travail : trois chapitres sont apparus (*Getting Around*, *Making a Living*, *Who Owns the Land?*), et deux phrases de l'après-guerre ont été corrigées. J'ai traité toutes les citations faibles de cette version. Les numéros de ligne (L…) sont indicatifs.

**Périmètre.**
1. Toutes les citations vers en.wikipedia.org (et es.wikipedia.org), *.blogspot.com et *.wordpress.com.
2. Les autres sources faibles repérées : blog personnel (wixsite), blog de transcriptions (batangashistory.date), reprise d'article (ntucphl.org), agrégateurs (PhilAtlas, City Population, EurekAlert, ChanRobles), lien cassé d'Experience Biñan.
3. À la demande du coordinateur : trois encadrés « Did You Know? » sans source (ombrelles de Ma-i, *Noli me tangere*, statue de Rizal).

**Méthode.** Pour chaque article Wikipédia, j'ai lu le texte source pour voir sa référence. Ensuite j'ai ouvert la source de remplacement et j'ai cherché le passage exact. Quand un site bloquait l'accès depuis cet environnement, je le dis. Project Gutenberg était en maintenance le 16/11/2026. J'ai donc lu les mêmes fichiers Gutenberg sur leurs copies Internet Archive.

**Actions.**
- *remplacer* : la nouvelle source soutient la phrase telle quelle.
- *remplacer + nuancer* ou *corriger* : la nouvelle source oblige à changer la phrase. La phrase anglaise proposée est dans les détails.
- *retirer* : supprimer un détail que rien ne soutient.
- *ajouter* : l'encadré n'avait pas de source, j'en ai trouvé une.
- *garder* : la citation peut rester.

**Format.** Dans le tableau, le signe « \| » à l'intérieur des liens est un échappement de Markdown. Dans le livre, on écrit la citation avec un « | » simple : `([Titre court, année | Éditeur ou site](URL))`.

---

## Les découvertes principales

1. **Laguna Caldera.** Le Global Volcanism Program (Smithsonian) donne aujourd'hui **12 × 24 km**, pas 20 × 10 km.
2. **Plaque de cuivre de Laguna.** Le chiffre de « 865 grams » vient, dans Wikipédia, d'un site personnel (Paul Morrow). Postma ne donne pas de poids en grammes. Il faut retirer ce chiffre (deux fois).
3. **Mateo Capinpin.** Morton dit qu'il est **capturé le 5 avril 1942**, quatre jours avant la reddition. Il n'a donc pas commandé « jusqu'à la reddition ». La rue « Gen. M. Capinpin Street » est confirmée par le NHCP. Sa mort à Biñan (1958) n'est confirmée par aucune source fiable que j'ai pu ouvrir.
4. **Marcial Alintanahin.** Le musée de la ville expose une **plaque de la Hunters ROTC Association, datée du 25 avril 1971, au nom de « Hunter Marcial Alintanahin »** (photo CC0 sur Wikimedia Commons). Il était donc bien membre des Hunters. Son rôle le 3 février 1945 et ses mandats de président municipal restent non confirmés.
5. **Los Baños, « some 1,500 civilians ».** Non soutenu. Le dossier du procès Yamashita parle d'environ **264 civils tués** dans les barrios de Los Baños (21 février – 6 mars 1945). Le chiffre de 1 500 concerne en fait **Calamba** (rapport n° 214 cité par Henderson).
6. **Hôtel de ville (après-guerre).** L'URL Experience Biñan avec « ñ » renvoie une **erreur 404**. La bonne page ne parle ni de 1954–55, ni du 15 mai 1955, ni du maire Jesus M. Garcia.
7. **PhilAtlas.** La page donne **40,27 km²** (chiffre de 2013) et 11,8 m d'altitude, pas 43,5 km².
8. **City Population.** La page citée ne donne **pas** le chiffre de 1948 (20 794).
9. **Zhao Rugua.** « before doing business » et « until the next season » ne sont pas dans le texte. Hirth et Rockhill disent « eight or nine months ». Ils datent aussi la première mention de Ma-i de **982**, alors que le chapitre dit 971.
10. **Statue de Rizal.** « one of the oldest Rizal monuments in the country » n'est soutenu par aucune source. Le reste de l'encadré est confirmé.
11. **SLEX en 1978.** Seul un annuaire imprimé (*Fookien Times Yearbook* 1978, cité par Wikipédia) donne l'inauguration du 4 mai 1978. La JICA confirme le début des travaux en 1974.
12. **Lien SEAFDEC faux (hors périmètre).** Le handle 10862/6120 (« SEAFDEC, LakeCon 2003 », chapitre Géographie) est un article sur le **lac Manguao (Palawan)**, pas sur Laguna de Bay.

---

## Tableau récapitulatif

| N° | Chapitre (ligne) | Affirmation | Ancienne URL | Nouvelle citation | Vérification | Action |
|---|---|---|---|---|---|---|
| 1 | Geography (L9) | Superficie 43,5 km² ; rivage à ~2 m ; Poblacion à ~12 m | https://www.philatlas.com/luzon/r04a/laguna/binan.html | [Biñan City, Laguna Profile \| PhilAtlas](https://www.philatlas.com/luzon/r04a/laguna/binan.html) + [Laguna de Bay \| World Lake Database, ILEC](https://wldb.ilec.or.jp/Display/html/3530) | PhilAtlas : « Area (2013) 40.27 km² », altitude « 11.8 meters ». ILEC : lac à « 1.8 m above sea level ». | corriger |
| 2 | Geography (L9) | Makiling 1 090 m à ~27 km ; Taal à 39 km | https://en.wikipedia.org/wiki/Mount_Makiling | [Mount Makiling Forest Reserve \| ASEAN Clearing House Mechanism](https://asean.chm-cbd.net/node/538) + [Taal (273070) \| Smithsonian Global Volcanism Program](https://volcano.si.edu/volcano.cfm?vn=273070) | ASEAN : « rises to an elevation of 1,090 m ». Distances recalculées : Makiling ≈ 26 km, Taal ≈ 38 km. | remplacer |
| 3 | Geography (L13) | Fosse de Manille, fosse des Philippines, faille de 1 200 km | https://en.wikipedia.org/wiki/Philippine_fault_system | [Large-scale digital mapping of the Philippine fault zone \| PHIVOLCS](https://www.phivolcs.dost.gov.ph/philippine-fault-zone-maps/) + [MMEIRS Final Report, ch. 2, 2004 \| JICA](https://openjicareport.jica.go.jp/pdf/11763745_02.pdf) | PHIVOLCS : « The 1,200-km-long Philippine fault zone ». JICA : fosse de Manille à l'ouest, « East Luzon Trench » à l'est. | remplacer + nuancer |
| 4 | Geography (L17) | Couloir de Macolod, ~40 km de large, rift | https://en.wikipedia.org/wiki/Macolod_Corridor | [Förster et al., The Macolod Corridor: A rift crossing the Philippine island arc, 1990 \| Tectonophysics](https://doi.org/10.1016/0040-1951%2890%2990420-D) + [Sudo et al., New 40Ar/39Ar and K-Ar ages from Macolod Corridor, 2007 \| Goldschmidt Conference Abstracts](https://goldschmidtabstracts.info/2007/984.pdf) | Titre de Förster vérifié (Crossref), article payant non lu. Sudo : champ volcanique « including Taal caldera », caldeira « at the north of Macolod Corridor ». 40 km non vérifié. | remplacer + nuancer |
| 5 | Geography (L17) | Caldeira de 20 × 10 km ; éruptions entre ~2 Ma et 30 ka | https://en.wikipedia.org/wiki/Laguna_Caldera | [Laguna Caldera (273080) \| Smithsonian Global Volcanism Program](https://volcano.si.edu/volcano.cfm?vn=273080) + Sudo et al. 2007 (n° 4) | GVP : « The elliptical 12 x 24 km Laguna Caldera… about 1 million and 27,000-29,000 years ago ». Sudo : dépôts de 1,83 à 1,92 Ma. | corriger |
| 6 | Geography (L25) | ~911 km², 2,5–2,8 m, 21 rivières, bassin 3 820 km², exutoire unique | https://en.wikipedia.org/wiki/Laguna_de_Bay | [Santos-Borja & Nepomuceno, Laguna de Bay: Experience and Lessons Learned Brief, 2006 \| ILEC](https://wldb.ilec.or.jp/itembyid/10202) + [Laguna de Bay \| World Lake Database, ILEC](https://wldb.ilec.or.jp/Display/html/3530) + [Philippines – Environmental tax, Laguna de Bay \| Convention on Biological Diversity](https://www.cbd.int/financial/fiscalenviron/philippines-fiscaltax.pdf) | LLDA : 900 km², 2,5 m, « only one outlet ». ILEC : 2,8 m, bassin « ca. 3,820 ». CBD : « Twenty-one major rivers ». 911 km² introuvable. | remplacer + nuancer |
| 7 | Geography (L41) | Sols profonds, argileux, sur tuf ; terrain presque plat | https://en.wikipedia.org/wiki/Bi%C3%B1an | [CALA East–West National Road Project, EIS, ch. 3, 2006 \| JICA](https://openjicareport.jica.go.jp/pdf/11844016_05.pdf) | Sols Guadalupe et Carmona : argiles profondes, issues de dépôts pyroclastiques. Rien de propre à Biñan. Wikipédia n'a pas de référence ici. | nuancer |
| 8 | Geography, DYK (L79) | Nom « lac de Bay » ; Bay première capitale ; reflux d'eau salée | https://en.wikipedia.org/wiki/Laguna_de_Bay | [Account of the encomiendas, 1591, in Blair & Robertson vol. 8 \| Project Gutenberg](https://www.gutenberg.org/ebooks/13742) + [Santos-Borja & Nepomuceno 2006 \| ILEC](https://wldb.ilec.or.jp/itembyid/10202) | B&R : « the province of La Laguna, which is the alcalde-mayoralty styled by itself Bay ». LLDA : « Pasig River becomes a tributary ». « First capital » : pas de source primaire. | remplacer + nuancer |
| 9 | Geography, DYK (L83) | Taal, « un des plus actifs du monde », à 39 km ; cendres 2020 | https://en.wikipedia.org/wiki/Taal_Volcano | [Delos Reyes et al., A synthesis and review of historical eruptions at Taal Volcano, 2018 \| PHIVOLCS](https://www.phivolcs.dost.gov.ph/tag/taal-volcanoeruption-historytephra-fallpyroclastic-density-currentvolcanic-tsunami/) + [Impacts of Taal Volcano phreatic eruption (12 January 2020), 2020 \| Manila Observatory](https://www.observatory.ph/2020/04/20/impacts-of-taal-volcano-phreatic-eruption-12-january-2020-on-the-environment-and-population-satellite-based-observations-compared-with-historical-records/) | PHIVOLCS : « second most frequently erupting volcano of the 24 active volcanoes in the Philippines ». Cendres jusqu'à la Laguna et Metro Manila. | remplacer + nuancer |
| 10 | Prehistory (L91) | *H. erectus* à Java vers 1,5–1,3 Ma | https://www.eurekalert.org/news-releases/766424 | [Matsu’ura et al., Age control of the first appearance datum for Javanese Homo erectus, 2020 \| Science (PubMed)](https://pubmed.ncbi.nlm.nih.gov/31919224/) | « most likely ~1.3 million years ago and less than 1.5 million years ago ». | remplacer |
| 11 | Prehistory (L111) | Pétroglyphes : Botong Francisco 1965, >100 figures, Trésor 1973, âge inconnu | https://en.wikipedia.org/wiki/Angono_Petroglyphs | [Presidential Decree No. 260, 1973 \| Lawphil](https://lawphil.net/statutes/presdecs/pd1973/pd_260_1973.html) + [Angono Binangonan Petroglyphs, 2022 \| National Museum of the Philippines](https://www.nationalmuseum.gov.ph/regional-museums/angono-binangonan-petroglyphs/) + [Philippines’ oldest artworks in danger of disappearing (AFP), 2014 \| Gulf Times](https://www.gulf-times.com/story/389513/philippines-oldest-artworks-in-danger-of-disappearing) | PD 260 : « Petroglyphs of the Rockshelter in Angono… National Cultural Treasures ». Musée : 179 figures, âge « cannot be established ». AFP : Francisco, 1965. | remplacer + nuancer |
| 12 | Pre-Colonial (L169) | Contenu de la plaque ; dette « about 865 grams » | https://en.wikipedia.org/wiki/Laguna_Copperplate_Inscription | [Postma, The Laguna Copper-Plate Inscription: A Valuable Philippine Document, 1991 \| Bulletin of the Indo-Pacific Prehistory Association](https://journals.lib.washington.edu/index.php/BIPPA/article/view/11381) | p. 162 (langues, écriture kawi) et p. 170 (traduction : Tundun, Namwran, 1 kati 8 suwarna, Puliran, Pailah, Binwangan). Aucun poids en grammes. | remplacer + retirer « 865 g » |
| 13 | Pre-Colonial (L185) | *Boxer Codex*, vers 1590, couples tagalogs | https://en.wikipedia.org/wiki/Boxer_Codex | [Boxer Codex Illustrations \| CCP Encyclopedia of Philippine Art](https://epa.culturalcenter.gov.ph/3/82/2152) + [Boxer Codex, ca. 1590, digital copy \| Lilly Library, Indiana University](https://purl.dlib.indiana.edu/iudl/general/VAB8326) | CCP : « Ca 1590… Philippine illustrations include those of the Tagalog ». | remplacer |
| 14 | Pre-Colonial, DYK (L219) | 865 g ; trouvée en draguant du sable à Lumban | https://en.wikipedia.org/wiki/Laguna_Copperplate_Inscription | [Postma 1991 \| Bulletin of the Indo-Pacific Prehistory Association](https://journals.lib.washington.edu/index.php/BIPPA/article/view/11381) | p. 164 : « during dredging of sand from the delta of the Lumbang River ». | remplacer + retirer « 865 g » |
| 15 | Pre-Colonial, DYK (L223), **sans source** | Ombrelles blanches offertes aux chefs de Ma-i ; crédit | — | [Hirth & Rockhill, Chau Ju-kua, 1911, ch. 40 « Ma-i », p. 159–160 \| Internet Archive](https://archive.org/details/chaujukuahiswork00chauuoft/page/159/mode/2up) + [A Chinese Gazetteer of Foreign Lands (Zhufan zhi) \| Global Middle Ages Project](https://globalmiddleages.org/node/109) | « The chiefs are in the habit of using white umbrellas, for which reason the traders offer them as gifts ». Remboursement après « eight or nine months ». Texte composé en 1224–1225. | ajouter + nuancer |
| 16 | Spanish (L239) | Étendue de l'Hacienda de Biñán ; Santa Rosa > moitié | https://philippinevillas.blogspot.com/2007/06/epilogue-last-aspirants-bin-and-santa.html | [Santiago, Pomp, Pageantry and Gold: The Eight Spanish Villas in the Philippines (1565–1887), 2005 \| Philippine Quarterly of Culture and Society](https://www.jstor.org/stable/29792568) + [Buzeta & Bravo, Diccionario geográfico… de las Islas Filipinas, t. I, 1850, art. « Biñan » \| Bayerische Staatsbibliothek](https://www.digitale-sammlungen.de/en/view/bsb10433378?page=414) | Santiago : trois phases, 1641–1677 ; Santa Rosa « more than half ». Buzeta : limites lac / Santa Rosa / Cavite / San Pedro Tunasan. | remplacer + nuancer |
| 17 | Spanish (L251) | Description de Buzeta & Bravo (riz, fermiers, *casamajanes*) | https://es.wikipedia.org/wiki/Bi%C3%B1%C3%A1n | [Buzeta & Bravo, Diccionario geográfico, estadístico, histórico de las Islas Filipinas, t. I, 1850, art. « Biñan », p. 388–390 \| Bayerische Staatsbibliothek](https://www.digitale-sammlungen.de/en/view/bsb10433378?page=414) | Lu dans l'OCR du scan de 1850 : tous les détails sont là. | remplacer |
| 18 | Spanish (L263) | 1896–97 : bases et hôpitaux ; 19 janv. 1898 : projet de *villa* | https://philippinevillas.blogspot.com/2007/06/epilogue-last-aspirants-bin-and-santa.html | [Santiago, Pomp, Pageantry and Gold, 2005 \| Philippine Quarterly of Culture and Society](https://www.jstor.org/stable/29792568) | « provisional hospitals… By January 19, 1898… a villa with the title "siempre fiel y leal" ». | remplacer |
| 19 | Spanish (L297) | Photo de l'Alberto Mansion | https://commons.wikimedia.org/wiki/File:Alberto_Mansion_Bi%C3%B1an_City,_Laguna.jpg | (inchangé) | C'est un crédit photo, pas la source d'un fait. | garder |
| 20 | Spanish (L297) | Alberto Mansion : début XIXe ; maison ancestrale de Teodora Alonso | https://mikeegungon.wixsite.com/welcome/post/history-reborn-the-alberto-mansion-in-bi%C3%B1an | [Roof of the ancestral house of Rizal’s mother caves in, 2012 \| GMA News](https://www.gmanetwork.com/news/topstories/regions/278933/roof-of-the-ancestral-house-of-rizal-s-mother-caves-in/story) + [Historic Alberto Mansion \| Experience Biñan](https://experiencebinan.com/histo-cultural/historic-alberto-mansion/) | GMA : « the home where Teodora Alonso… grew up », « 200-year-old ». Experience Biñan : « perhaps two hundred years ago ». | remplacer + nuancer |
| 21 | Spanish, DYK (L299) | Projet de *villa* ; Bautista lit l'Acte d'indépendance | https://philippinevillas.blogspot.com/2007/06/epilogue-last-aspirants-bin-and-santa.html | [Santiago 2005 \| Philippine Quarterly of Culture and Society](https://www.jstor.org/stable/29792568) + [Ambrosio Rianzares Bautista (1830–1903) \| NHCP](https://philhistoricsites.nhcp.gov.ph/registry_database/ambrosio-rianzares-bautista-1830-1903-2/) | NHCP : auteur de l'Acta, « binasa sa harap ng madla, Kawit, Cavite, 12 Hunyo 1898 ». | remplacer |
| 22 | Spanish, DYK (L307), **sans source** | Tasio et la foudre sur l'église de Biñan « en l'an 70 » | — | [Rizal, Noli me tángere, 1887, cap. XIV « Tasio el loco ó el filósofo » \| Project Gutenberg](https://www.gutenberg.org/ebooks/47584) + [Rizal, The Social Cancer, trans. Derbyshire, 1912, ch. XIV « Tasio: Lunatic or Sage » \| Project Gutenberg](https://www.gutenberg.org/ebooks/6737) | « ¿por qué el año 70 cuando cayó un rayo en Biñan, cayó precisamente en la torre y destrozó reloj y un altar? » | ajouter |
| 23 | American (L315) | Traité de Paris, 20 millions de dollars | https://www.chanrobles.com/treatyofparis.htm | [Treaty of Peace Between the United States and Spain, 1898, art. III \| Avalon Project, Yale Law School](https://avalon.law.yale.edu/19th_century/sp1898.asp) | « The United States will pay to Spain the sum of twenty million dollars ($20,000,000) ». | remplacer |
| 24 | WWII (L393) | Capinpin commande la 21ᵉ division « until the surrender… and the Death March » | https://en.wikipedia.org/wiki/Mateo_M._Capinpin | [Morton, The Fall of the Philippines, 1953, ch. XXIV \| HyperWar](https://www.ibiblio.org/hyperwar/USA/USA-P-PI/USA-P-PI-24.html) | « the 21st Division, led by Brig. Gen. Mateo Capinpin… been captured by the Japanese » le 5 avril 1942 (dimanche de Pâques). | corriger |
| 25 | WWII (L405) | 45ᵉ régiment : 80 combattants, ~5 000 *home guards* | https://www.batangashistory.date/2020/04/commanders-report-on-strength-and-composition-of-hunters-rotc-dec-44.html | [Adevoso, Commander’s report on the Hunters-ROTC, Dec. 1944, RG 407, Entry 1087, Box 246 \| U.S. National Archives](https://www.archives.gov/research/military/ww2/philippine/guerrilla-list-2.html) | Chiffres vérifiés dans la transcription. Original à la NARA, non numérisé à ma connaissance. | remplacer (référence d'archive) |
| 26 | WWII (L409) | Alintanahin, ancien président municipal, chef de terrain | https://pepealas.wordpress.com/2023/02/03/liberation-of-binan/ | garder + [Hunters ROTC Association plaque for “Hunter Marcial Alintanahin”, 1971, Biñan City Museum \| Wikimedia Commons](https://commons.wikimedia.org/wiki/File:6361Bi%C3%B1an_City,_Laguna_Museum_Old_Town_Hall_Landmarks_08.jpg) | Le blog dit « Led by Commander Marcial Alintanahin (two-time Municipal President) », sans source. La plaque confirme qu'il était un Hunter. | garder + préciser |
| 27 | WWII (L413) | Los Baños : ~1 500 civils tués ; Calamba : « about one thousand » | https://www.batangashistory.date/2022/08/charge-and-specifications-vs-japanese.html | [Whitney, The Case of General Yamashita: A Memorandum, c. 1950, app. C \| Library of Congress](https://tile.loc.gov/storage-services/service/ll/llmlp/Yamashita/Yamashita.pdf) + [Henderson, Rescue at Los Baños, 2015 \| Internet Archive](https://archive.org/details/rescue-at-los-banos-the-most-daring-prison-camp-raid-of-world-war-2) | Yamashita : ~264 tués dans les barrios de Los Baños ; « More than 2,500… of Calamba… on a single day, 12 February 1945 ». Henderson : « An estimated 1,500 citizens of Calamba ». | corriger |
| 28 | WWII, DYK (L463) | Capinpin passe ses dernières années à Biñan, y meurt en 1958 ; une rue porte son nom | https://en.wikipedia.org/wiki/Mateo_M._Capinpin | [Rizal in Biñan (marker, location) \| NHCP](https://philhistoricsites.nhcp.gov.ph/registry_database/rizal-in-binan/) + [Morton 1953, ch. XXIV \| HyperWar](https://www.ibiblio.org/hyperwar/USA/USA-P-PI/USA-P-PI-24.html) | NHCP : « Location: Gen. M. Capinpin Street, Biñan, Laguna ». Mort à Biñan : seulement Wikipédia (sans référence), Find a Grave, blogs. | nuancer |
| 29 | Post-War (L471) | 1954–55 ; hôtel de ville le 15 mai 1955 sous le maire Jesus M. Garcia | https://experiencebinan.com/histo-cultural/sentrong-pangkultura-ng-biñan/ | [Sentrong Pangkultura ng Biñan \| Experience Biñan](https://experiencebinan.com/histo-cultural/sentrong-pangkultura-ng-binan/) | L'URL avec « ñ » donne 404. La bonne page dit : école sous les Américains, puis hôtel de ville quand l'école part à Santo Domingo ; portique ajouté. Pas de date ni de maire. | corriger l'URL + nuancer |
| 30 | Post-War (L487) | Biñan 1ʳᵉ municipalité du pays par le revenu en 2008 | https://ntucphl.org/?p=28247 | [Most top-earning towns derive income from IRA (NSCB data), BusinessMirror, 2010, reprinted \| NTUC Philippines](https://ntucphl.org/2010/10/most-top-earning-towns-derive-income-from-ira/) + [About Biñan \| Experience Biñan](https://experiencebinan.com/about-binan/) | « Topping the list of the high-income towns was Biñan, Laguna » (chiffres 2008 de la NSCB, chronique de R. Virola). EB : « the richest municipality in the Philippines before its cityhood ». | garder (faute de mieux) |
| 31 | Post-War (L499) | 2024 : 584 479 hab. ; 2020 : 3ᵉ avec 407 437 | https://citypopulation.de/en/philippines/luzon/admin/laguna/043403__bi%C3%B1an/ | garder [Biñan City population \| City Population, after PSA](https://citypopulation.de/en/philippines/luzon/admin/laguna/043403__bi%C3%B1an/) ; idéalement le tableau PSA 2024 | Chiffres vérifiés (Calamba 539 671 et Santa Rosa 414 812 en 2020). PSA bloqué ici. La page Lawphil de la Proclamation 973 n'a pas le tableau. | garder |
| 32 | Post-War, DYK (L533) | Hôtel de ville le 15 mai 1955 ; musée depuis 2016 | https://experiencebinan.com/histo-cultural/sentrong-pangkultura-ng-biñan/ | [Sentrong Pangkultura ng Biñan \| Experience Biñan](https://experiencebinan.com/histo-cultural/sentrong-pangkultura-ng-binan/) | 2016 et « Sentrong Pangkultura ng Biñan » : oui. 15 mai 1955 : non. | corriger l'URL + nuancer |
| 33 | Post-War, DYK (L537) | 2024 : 584 479 contre 575 046 à Calamba ; 1948 : 20 794 | https://citypopulation.de/en/philippines/luzon/admin/0434__laguna/ | garder pour 2020–2024 ; source à trouver pour 1948 | 2024 vérifié. La page ne remonte qu'à 1990 : 20 794 (1948) n'y est pas. | nuancer (1948) |
| 34 | Post-War, DYK (L541), **sans source** | Statue de Rizal (1918), foudre en 2015, réplique 2016, original à l'hôtel de ville en 2017 | — | [Century-old Rizal monument in Laguna ‘beheaded’ by lightning, 2015 \| Philstar](https://www.philstar.com/lifestyle/arts-and-culture/2015/09/24/1503488/century-old-rizal-monument-laguna-beheaded-lightning) + [Rizal Monument \| Experience Biñan](https://experiencebinan.com/heritage-sites/rizal-monument) + [Unveiling Biñan’s Gallant History, 2017 \| City of Biñan](https://www.binan.gov.ph/unveiling-binans-gallant-history/) | Dates et Caedo confirmés. Tolentino : « believed ». « One of the oldest » : non soutenu. | ajouter + nuancer |
| 35 | Getting Around (L621) | SLEX atteint Biñan en 1978 | https://en.wikipedia.org/wiki/South_Luzon_Expressway | [Motorway Development Study in the Kingdom of Thailand, Final Report: Motorways in Various Countries, p. 142 \| JICA](https://openjicareport.jica.go.jp/pdf/10928083_03.pdf) + [Presidential Decree No. 1113, 1977 \| Lawphil](https://lawphil.net/statutes/presdecs/pd1977/pd_1113_1977.html) | JICA : « In 1974, PNCC simultaneously started working on the 30 kilometer SLE Extension from Muntinlupa to Calamba ». PD 1113 (1977) parle de cette extension. 4 mai 1978 : source imprimée seulement. | remplacer + nuancer |
| 36 | Getting Around (L621) | CALAX depuis 2019 à Mamplasan ; sortie Technopark en 2020 | https://en.wikipedia.org/wiki/Cavite%E2%80%93Laguna_Expressway | [CALAX toll collection, 2020 \| CALAX (MPT South)](https://www.calax.com.ph/news-and-events/2020/calax-toll-collection) + [DPWH inspects CALAX Laguna segment interchanges, 2020 \| CALAX (MPT South)](https://www.calax.com.ph/news-and-events/2020/dpwh-inspects-calax-laguna-segment-interchanges-sec-villar-announces-august-18) | « between SLEX Mamplasan Exit and… Santa Rosa-Tagaytay… Partial operations started October 30 ». Échangeur Technopark ouvert « at 12:01AM on August 18, 2020 ». | remplacer |
| 37 | Making a Living (L701) | Revenu ordinaire : 607 M₱ (2009) → 1,53 Md₱ (2016) | https://www.philatlas.com/luzon/r04a/laguna/binan.html | garder (PhilAtlas reprend le BLGF) ; idéalement les données du Bureau of Local Government Finance | PhilAtlas : 606 915 277 ₱ (2009), 1 530 342 881 ₱ (2016), « from the Bureau of Local Government Finance ». Site du BLGF bloqué ici. | garder |

---

## Détails, citation par citation

### Géographie

**1. Superficie et altitude (L9).**
- Phrase : « Biñan occupies a narrow strip of land of about **43.5 square kilometres** on the western shore of **Laguna de Bay**… From the lakeshore, barely 2 metres above sea level, the ground rises gently westward…; the old town centre, the Poblacion, stands at about 12 metres. »
- Constat : PhilAtlas ne dit pas 43,5 km². Il dit : « Area (2013) 40.27 km² » et « Elevation at these coordinates is estimated at 11.8 meters ». Le chiffre de 43,5 km² est celui de la ville (Wikipédia, et Experience Biñan quand il dit que Casile fait « 12 hectares (only 0.27% of the city) »). Pour le rivage, l'ILEC place le lac à « 1.8 m above sea level ».
- Phrase proposée : « Biñan occupies a narrow strip of land of some 40 to 44 square kilometres (43.5 km² in the city's own figures, 40.27 km² in the national statistics of 2013) on the western shore of Laguna de Bay… the old town centre, the Poblacion, stands at about 12 metres. »
- Note : la même superficie de 43,5 km² revient au chapitre Après-guerre (« On its 43.5 square kilometres »). Il faut harmoniser.

**2. Mont Makiling et Taal (L9).**
- Phrase : « **Mount Makiling** (1,090 m) about 27 km to the southeast, **Mount Banahaw** beyond it, and to the southwest the Tagaytay ridge, the rim of the caldera that hides **Taal Volcano**, only 39 km away. »
- Source de Wikipédia : le Global Volcanism Program (GVP). Le site web du GVP était en panne (« Smithsonian site temporarily unavailable »). J'ai lu les données du GVP par son service cartographique.
- Remplacement : [Mount Makiling Forest Reserve | ASEAN Clearing House Mechanism](https://asean.chm-cbd.net/node/538) : « The mountain rises to an elevation of 1,090 m (3,580 ft) above mean sea level ».
- Distances : je les ai recalculées avec les coordonnées de PhilAtlas pour Biñan (14,3373 N ; 121,0827 E), du GVP pour Taal (14,0106 N ; 120,9975 E) et du sommet de Makiling (14,13 N ; 121,20 E). Résultat : Makiling ≈ 26 km au sud-sud-est, Taal ≈ 38 km au sud-sud-ouest. Les chiffres du livre (27 et 39 km) sont proches. On peut écrire « about 26 km » et « only 38 km away ».
- Action : remplacer.

**3. Fosses et faille philippine (L13).**
- Phrase : « Each plate dives beneath the islands along its own deep-sea trench: the **Manila Trench** west of Luzon and the **Philippine Trench** in the east, while the sideways part of the movement is absorbed by the 1,200-kilometre **Philippine Fault**. »
- Sources de Wikipédia : Aurelio 2000 (*Island Arc*), Barrier et al. 1991 (*Geology*). Wiley et l'ADS étaient bloqués.
- Remplacement vérifié :
  - PHIVOLCS : « The 1,200-km-long Philippine fault zone (PFZ) is a major tectonic feature that transects the whole Philippine archipelago from northwestern Luzon to southeastern Mindanao. This arc-parallel, left-lateral strike slip fault… »
  - JICA, MMEIRS (2004) : « The Eurasian Plate (or South China Plate) subducts eastward beneath Luzon Island along the Manila Trench, and the Philippine Sea Plate subducts westward along the East Luzon Trench simultaneously… The PFZ is assumed to release the shear stress caused by the oblique subduction of the ocean plates. »
- Nuance : à la hauteur de Luzon, la fosse orientale s'appelle la fosse (ou fossé) de l'est de Luzon. La fosse des Philippines commence plus au sud.
- Phrase proposée : « Each plate dives beneath the islands along its own deep-sea trench – the Manila Trench west of Luzon, and the East Luzon Trough and the Philippine Trench in the east – while the sideways part of the movement is absorbed by the 1,200-kilometre Philippine Fault. »
- Bonus : la PHIVOLCS liste **24 volcans actifs**. Cela soutient « some two dozen active volcanoes » (même paragraphe).

**4. Couloir de Macolod (L17).**
- Phrase : « Biñan lies at the northern end of the **Macolod Corridor**, a belt of volcanoes about 40 km wide that runs from Taal to Laguna de Bay, where the crust is being stretched and magma rises easily. »
- Sources de Wikipédia : Austria et al. 2023 (*Tectonophysics*) pour les 40 km ; Förster et al. 1990 ; Vogel et al. 2006. Ces articles sont payants (ScienceDirect et Springer bloqués).
- Vérifié :
  - Förster, Oles, Knittel et Defant, « The Macolod Corridor: A rift crossing the Philippine island arc », *Tectonophysics* 183 (1990) : 265–271 (métadonnées Crossref). Le titre soutient l'idée de rift.
  - Sudo et al. 2007 : « Macolod Corridor, southwestern Luzon, Philippines, is an extensive Quaternary volcanic field including Taal caldera and many monogenetic/polygenetic volcanoes… the possible caldera, located at the central lobe of the lake Laguna de Bay, at the north of Macolod Corridor… »
- Non vérifié : la largeur de 40 km.
- Phrase proposée : « Biñan lies at the northern edge of the Macolod Corridor, a belt of young volcanoes that runs from Taal to Laguna de Bay across the volcanic arc of Luzon; geologists see it as a rift, where the crust is being stretched and magma rises easily. »

**5. Caldeira de la Laguna (L17).**
- Phrase : « The central part of the lake is itself a volcanic **caldera**, some 20 by 10 kilometres, between Talim Island and the Jala-Jala peninsula, left by great explosive eruptions whose dates are still debated – somewhere between about two million and thirty thousand years ago. »
- Source de Wikipédia : GVP n° 273080. Résumé géologique actuel du GVP (lu par son service de données le 16/11/2026) : « The elliptical 12 x 24 km Laguna Caldera SE of Manila forms the middle of the three large basins of Laguna de Bay… That [Jalajala] peninsula is the E rim of the caldera… Elongated Talim Island forms the SW rim of the caldera. The caldera may have formed during at least two major explosive eruptions about 1 million and 27,000-29,000 years ago. »
- Sudo et al. 2007 : ponces de la coulée pyroclastique de Teresa, au nord du lac : « 1.85±0.01 Ma and 1.83±0.02 Ma ».
- Phrase proposée : « The central part of the lake is itself a volcanic caldera, some 24 by 12 kilometres, between Talim Island and the Jala-Jala peninsula. It was left by great explosive eruptions whose dates are still debated: about one million and 27,000 to 29,000 years ago according to the Smithsonian's Global Volcanism Program, while volcanic deposits north of the lake have been dated to about 1.8 to 1.9 million years. »

**6. Laguna de Bay en chiffres (L25).**
- Phrase : « **Laguna de Bay** is the largest lake in the Philippines – about 911 km² – but also a very shallow one, only 2.5 to 2.8 metres deep on average. Twenty-one rivers feed it from a basin of some 3,820 km², and it has a single outlet, the **Pasig River**, which carries its water to Manila Bay. »
- Sources de Wikipédia : « LLDA 1995 » (non en ligne), Gonzales 1987 (thèse).
- Vérifié :
  - LLDA (Santos-Borja et Nepomuceno, brief publié par l'ILEC) : « Laguna de Bay is the largest, with a surface area of 900 km2… one of the shallowest, with an average depth of 2.5 m… There is only one outlet, the 27 km Pasig River, which drains to Manila Bay. » Le même texte dit : « More than 100 streams flow into its drainage area, which is divided into 24 sub-basins. »
  - ILEC, World Lake Database : « Surface area ca. 900… Mean depth [m] 2.8… Catchment area [km2] ca. 3,820… Number of outflowing rivers and channels (name): 1 (Pasig R.) ».
  - CBD : « Twenty-one major rivers flow into the 90,000-hectare lake ».
  - Santos-Borja 2020 (ILEC) confirme « the 900 km2 lake » et « total basin area of 3820 km2 ».
- Non trouvé : 911 km².
- Phrase proposée : « Laguna de Bay is the largest lake in the Philippines – about 900 km² – but also a very shallow one, only 2.5 to 2.8 metres deep on average. Twenty-one major rivers and more than a hundred streams feed it from a basin of some 3,820 km², and it has a single outlet, the Pasig River, which carries its water to Manila Bay. »
- Note : la version tagalog (« mga 911 kilometro kuwadrado ») doit suivre.

**7. Les sols de Biñan (L41).**
- Phrase : « Volcanic ash, river silt and flat land made Biñan fertile: its soils, formed on volcanic tuff and ash, are deep and rich in clay, and most of its land is almost level. »
- Constat : l'article Wikipédia « Biñan » donne des chiffres (Carmona 2 577 ha, Guadalupe 660 ha, Lipa 798 ha, Gingua 315 ha ; 85,33 % du territoire en pente de 0 à 2,5 %) **sans aucune référence**. Ces chiffres viennent sans doute du plan d'occupation des sols (CLUP) de la ville, que je n'ai pas trouvé en ligne.
- Source trouvée : JICA, EIS de la route CALA (2006), sur le plateau Cavite–Laguna (dont 4 barangays de Biñan) : les sols Guadalupe sont des argiles « moderately deep-to-deep, well drained… clay » ; le Carmona clay loam provient de « volcanic landscapes derived from pyroclastic materials » ; l'alluvion couvre « the western edge of Laguna de Bay ».
- Phrase proposée : « Volcanic ash, river silt and flat land made Biñan fertile: the soils of this part of the Cavite–Laguna plateau, formed on volcanic tuff, are mostly deep clays and clay loams, and the plain near the lake is almost level. »
- À faire : demander le CLUP au City Planning and Development Office pour des chiffres propres à Biñan.

**8. Encadré « Laguna de Bay is not a "bay" » (L79).**
- Phrase : « The name means "the lake of Bay", after the lakeshore town of Bay, the first capital of the province. In the dry season the lake even drinks from the sea: the Pasig River flows backward and brings in salt water from Manila Bay. »
- Source de Wikipédia : Feleo et Sheniak 2002 (livre de l'Unilever, non en ligne) ; Jocano 1973 pour « capitale de 1581 à 1688 ».
- Vérifié :
  - Blair & Robertson vol. 8 (relevé des encomiendas, 1591) : « Thus the province of La Laguna, which is the alcalde-mayoralty styled by itself Bay, has eleven thousand five hundred whole tributes ».
  - LLDA : « When the lake level is lower than Manila Bay, and when there is sufficient tidal fluctuation that can push saltwater into the lake, Pasig River becomes a tributary. During backflow of the Pasig River, therefore, Laguna de Bay becomes a brackish water lake. »
- Non vérifié par une source primaire : « the first capital ». Experience Biñan (site du BCHATO) dit seulement « Baý served as the capital of the province until 1688 ».
- Phrase proposée : « The name means "the lake of Bay", after the lakeshore town of Bay: in 1591 the whole province was still "the alcalde-mayoralty styled Bay". In the dry season the lake even drinks from the sea: when its level falls below that of Manila Bay, the Pasig River flows backward and brings in salt water. »

**9. Encadré sur le Taal (L83).**
- Phrase : « **One of the world's most active volcanoes is Biñan's neighbour.** Taal Volcano lies only about 39 km from the Poblacion, hidden behind the Tagaytay ridge – barely farther than Manila. The ash of its January 2020 eruption reached the city. »
- Vérifié :
  - PHIVOLCS (résumé de Delos Reyes et al. 2018) : « Taal Volcano in Southern Luzon is the second most frequently erupting volcano of the 24 active volcanoes in the Philippines. »
  - GVP : « Taal is one of the most active volcanoes in the Philippines ».
  - Manila Observatory (2020) : « Taal Volcano's Phreatic Eruption emitted ash that reached Batangas, Cavite, Laguna, Metro Manila, Rizal, and Bulacan… on January 12, 2020 » ; « ashfall covering the towns west and north of Taal Lake reaching as far as Sta. Rosa, Laguna ».
- Nuance : « world's most active » n'est soutenu par aucune de ces sources. Biñan est entre Santa Rosa et Metro Manila : la cendre y est très probable, et l'article de l'Inquirer déjà cité au § Géographie (L17) le dit.
- Phrase proposée : « **One of the most active volcanoes of the Philippines is Biñan's neighbour.** Taal, the second most frequently erupting of the country's 24 active volcanoes, lies only about 38 km from the Poblacion, hidden behind the Tagaytay ridge. The ash of its January 2020 eruption fell across Laguna and Metro Manila, on either side of Biñan. »

### Préhistoire

**10. *Homo erectus* à Java (L91).**
- Phrase : « *Homo erectus* reached the island of Java around 1.5 to 1.3 million years ago. »
- EurekAlert est un service de communiqués de presse. L'original est l'article de *Science* 367 (2020) : 210–214, doi:10.1126/science.aau8556. Résumé (PubMed) : « We conclude that the first appearance datum for the Sangiran hominins is most likely ~1.3 million years ago and less than 1.5 million years ago ».
- Phrase proposée (plus exacte) : « *Homo erectus* reached the island of Java about 1.3 million years ago, and no earlier than 1.5 million years ago. »

**11. Pétroglyphes d'Angono–Binangonan (L111).**
- Phrase : « On a cliff of volcanic tuff in Binangonan, the painter Carlos "Botong" Francisco reported in 1965 more than a hundred figures carved into the rock: the **Angono–Binangonan Petroglyphs**, declared a National Cultural Treasure in 1973 and probably several thousand years old, although their exact age is unknown. »
- Sources de Wikipédia : fiche UNESCO (liste indicative n° 5018, bloquée ici), page ancienne du National Museum, PD n° 260.
- Vérifié :
  - PD n° 260 (1ᵉʳ août 1973) : « the Petroglyphs of the Rockshelter in Angono, Rizal… are hereby declared as National Cultural Treasures ».
  - National Museum (2022) : « zoomorphic and geometric figures engraved on a rock shelter found at the border of the municipalities of Angono and Binangonan… A total of 179 distinct figures… carved on a volcanic ash deposit known as tuff… the age of the petroglyphs itself cannot be established. »
  - Gulf Times, dépêche AFP (23/04/2014) : « The carvings were first documented by acclaimed Philippine artist Carlos Francisco in 1965 while he was leading a Boy Scout troop on a hike » ; « Museum scientists believe the carvings date back to 3000 BC… it is difficult to conclusively determine the age ».
- Phrase proposée : « On a cliff of volcanic tuff at the border of Angono and Binangonan, the painter Carlos "Botong" Francisco documented in 1965 more than a hundred figures carved into the rock – the National Museum now counts 179: the **Angono–Binangonan Petroglyphs**, declared a National Cultural Treasure in 1973. Museum scientists have suggested a date around 3000 BCE, but their age cannot be established. »

### Période précoloniale

**12. Contenu de la plaque de cuivre (L169).**
- Phrase : « Written mostly in Old Malay, in the Kawi script of Java, with Sanskrit and Old Javanese or Old Tagalog words, the **Laguna Copperplate Inscription** is a receipt: on behalf of the chief of **Tundun** (Tondo), it frees the descendants of a man named Namwaran from a debt of one *kati* and eight *suwarna* of gold – about 865 grams – before witnesses from places called Pailah, Puliran and Binuangan. »
- Constat dans Wikipédia : le chiffre « 865 grams » est sourcé par un site personnel (Paul Morrow, mts.net). Postma 1992 (*Philippine Studies*) n'était pas accessible (philippinestudies.net hors ligne, JSTOR bloqué).
- Remplacement lu (scan, pages vues une à une) : Postma 1991, *Bulletin of the Indo-Pacific Prehistory Association* 11 : 160–171.
  - p. 162 : « there are a number of technical Sanskrit words… and some Old Javanese words… However, the main language of the LCI is clearly Old Malay… which contains several words that are identical with or closely related to the Old Tagalog language ». Écriture : « the standard form of the Early Kawi script of the tenth century AD » (p. 161).
  - p. 170 (traduction) : « the honorable Namwran was given, as a special favor, a document of full acquittal by the chief and commander of Tundun… was cleared and acquitted of a salary-related debt of 1 kati and 8 suwarna in the presence of His Honor the Leader of Puliran… the Leader of Pailah… the Leader of Binwangan… all the descendants of the honorable Namwran have been cleared of the whole debt ». Note 7 : « Both words signify certain weights of gold. »
- Phrase proposée : retirer « – about 865 grams – ». Écrire : « …from a debt of one *kati* and eight *suwarna* – weights of gold – before witnesses from places called Pailah, Puliran and Binuangan. »

**13. *Boxer Codex* (L185).**
- Phrase : « Men and women wore gold bracelets and anklets; around 1590 the illustrated *Boxer Codex* shows Tagalog couples in their *baro*, *saya* and *putong*. »
- Remplacement : CCP Encyclopedia of Philippine Art (2020) : « Ca 1590… The Boxer Codex is an illustrated manuscript that contains the first illustrations of the inhabitants of the Philippines at Spanish contact… Philippine illustrations include those of the Tagalog, Visayas, Zambales, Cagayan, and Aeta… Details of costume and jewelry and accessories must have been rendered by a keen observer. » Le manuscrit lui-même est numérisé par la Lilly Library (lien PURL vérifié, il mène à la visionneuse de l'université d'Indiana).
- Remarque : les sentences précédentes du paragraphe (Bathala, Tala, Mapolon, *catalonan*, deuil de quatre jours) viennent de Plasencia. Elles mériteraient la citation de Plasencia (Gutenberg 13701), déjà utilisée au paragraphe d'avant.

**14. Encadré sur la plaque de cuivre (L219).**
- Phrase : « In the year 900, on behalf of the chief of Tondo, a copper plate cleared the descendants of Namwaran of a debt of about 865 grams of gold. It was found by chance, more than a thousand years later, by a worker dredging sand at Lumban, on the other shore of our lake. »
- Postma 1991, p. 164 : « the LCI was not found during a controlled archeological excavation but rather during dredging of sand from the delta of the Lumbang River, which flows into the Laguna de Bay, in the Province of Laguna ». La plaque a été proposée au National Museum en janvier 1990 (p. 160).
- Phrase proposée : « In the year 900, on behalf of the chief of Tondo, a copper plate cleared the descendants of Namwaran of a debt of one *kati* and eight *suwarna* of gold. It was found by chance, more than a thousand years later, while sand was being dredged at the mouth of the Lumbang River, on the other shore of our lake. »

**15. Encadré « Chinese merchants gave white umbrellas to local chiefs » (L223, sans source).**
- Phrase : « According to Zhao Rugua, writing in 1225, traders arriving in Ma-i offered white umbrellas to the chiefs before doing business – and then left their goods on credit until the next season. »
- Source : Hirth et Rockhill 1911, chapitre 40 « Philippine Islands (Ma-i) », p. 159–160 (OCR et pagination vérifiés sur Internet Archive) : « When trading ships enter the anchorage, they stop in front of the officials place, for that is the place for bartering of the country. After a ship has been boarded, the natives mix freely with the ship's folk. The chiefs are in the habit of using white umbrellas, for which reason the traders offer them as gifts. The custom of the trade is for the savage traders to assemble in crowds and carry the goods away with them in baskets… it takes them as much as eight or nine months till they return, when they repay the traders on shipboard with what they have obtained. »
- Date : Hirth et Rockhill (p. 137) plaçaient la rédaction entre 1242 et 1258. La recherche actuelle retient 1224–1225 : le Global Middle Ages Project écrit que Zhao Rugua « composed this text in 1224-1225 ». « Writing in 1225 » est donc correct, avec cette citation.
- Ce qui est faux ou ajouté : « before doing business » n'est pas dans le texte ; « until the next season » simplifie « eight or nine months ».
- Phrase proposée : « According to Zhao Rugua, writing in 1225, the chiefs of Ma-i liked to use white umbrellas, so the Chinese traders offered them as gifts – and the local traders then carried the goods away on credit, paying for them eight or nine months later. »
- Remarque sur le paragraphe principal (L173) : « Chinese annals mention merchants from a country called **Ma-i** as early as 971 ». Hirth et Rockhill (p. 160, note 1) : « The name of Ma-i was first heard of in China in A.D. 982 when some traders from that country brought valuable merchandise to Canton ». Il faut vérifier la date de 971 dans Go 2005, que je n'ai pas pu ouvrir.

### Période espagnole

**16. Étendue de l'Hacienda de Biñán (L239).**
- Phrase : « The **Hacienda de Biñan** stretched from the shore of the lake to the hills of Cavite and from San Pedro Tunasan to the south of Santa Rosa, which made up more than half of it. »
- Le blog « Philippine Villas » est la mise en ligne, par la Pila Historical Society Foundation, de l'article de L. P. R. Santiago, « Pomp, Pageantry and Gold: The Eight Spanish Villas in the Philippines (1565–1887) », *Philippine Quarterly of Culture and Society* 33, n° 1/2 (2005) : 57–75, épilogue « The Last Aspirants: Biñán and Santa Rosa (1898) ». La source fiable est donc l'article de revue. L'identifiant JSTOR 29792568 vient de la notice EDS de la bibliothèque de l'Académie des sciences tchèque (« edsjsr.29792568 ») ; JSTOR lui-même bloque les robots. Je n'ai pas pu ouvrir la page JSTOR.
- Texte de Santiago : « The Dominicans had carved out the Hacienda de Biñán in three phases in the town of Tabuco (now Cabuyao) from 1641 to 1677. Sta. Rosa… formed part of the bigger hacienda, comprising more than half of its total area… the friar lands of Biñán were allotted for the support and maintenance of the Royal and Pontifical University of Santo Tomás. »
- Les limites viennent de Buzeta et Bravo (1850, p. 389) : « el term. que posee este pueblo, ó sea la grande hacienda de Biñan de los padres dominicos, confina por E. con la laguna de Bay… por S. con Santa Rosa… Por O. confina con la prov. de Cavite, y por N. con San Pedro de Tunasan. »
- Nuance : « to the south of Santa Rosa » n'est dans aucune des deux sources. La date de « 1644 » (phrase précédente) n'est pas dans Santiago.
- Phrase proposée : « The Hacienda de Biñan, carved out in three phases between 1641 and 1677, ran from the shore of the lake to the border of Cavite and from San Pedro Tunasan southward; the lands of Santa Rosa made up more than half of it. »

**17. Buzeta et Bravo, 1850 (L251).**
- Phrase : « Around 1850, the geographers Buzeta and Bravo wrote that the Dominicans owned "all the land of Biñan" and that "it is rice that makes its wealth"… Tenants (*inquilinos*) cleared and farmed the land rent-free for three or four years, then paid five *cavans* of rice for each *cabalita* of irrigated land – hardly a tenth of the harvest – and could not be evicted unless they failed to pay for two years running. Many tenants were rich men who worked their land through sharecroppers, the *kasamá*. »
- Wikipédia en espagnol cite Buzeta et Bravo, t. I, p. 388–390 (copie de la Biblioteca Digital de Castilla y León, injoignable ici). J'ai lu l'OCR de l'exemplaire de la Bayerische Staatsbibliothek (pages 388 à 390 = images 414 à 416) :
  - « La casa de la Hacienda de los PP. dominicos, que son dueños de todo el terreno de Biñan »
  - « se aprovechasen de ellas por 3 ó 4 años sin pagar nada… diesen al propietario 5 cabanes de arroz por cada [ca]balita de tierra de regadío »
  - « dan las tierras de esta pobl. un 50 ó 60 por 1… lo que apenas llega al diezmo de sus prod. »
  - « ni se les pueden quitar las tierras, á no ser que dejen de pagar por dos años consecutivos »
  - « La hacienda de Biñan está la mayor parte en poder de ricos que la cultivan por medio de compañeros, á quienes dan el nombre de Casamahanes »
  - « es el arroz la que forma su riqueza »
- Action : remplacer. Remarque : le paragraphe suivant (« By the 19th century Biñan was a town of mestizos… "not a few of stone and tile, a luxury found in few towns"… 10,948 souls… church "of good construction"… joined to it through the choir loft… *calzadas* ») reprend aussi Buzeta (p. 388–390 : « no pocas de piedra y teja, lujo que se halla en escaso núm. de pueblos », « POBL. 10,948 alm. »). Il est cité seulement à Bowring. Il faut ajouter la même citation de Buzeta.

**18. 1896–1898 : bases espagnoles et projet de *villa* (L263).**
- Phrase : « In 1896–97, Biñan and Santa Rosa served as supply bases and field hospitals for the Spanish troops fighting the revolutionaries in Cavite, and on 19 January 1898 the provincial junta even proposed to make them *villas* with the motto "**Siempre fiel y leal**" – "ever faithful and loyal". »
- Santiago 2005 : « In the first phase of the Philippine Revolution in 1896-97, Biñán and Sta. Rosa together with Cabuyao and Calamba, hosted the Spanish counter-insurgency troops in stockpiling and transporting military and food supplies. They also organized provisional hospitals for the Spanish wounded… By January 19, 1898, the provincial board of Laguna, supported by the head of the Dominican order, submitted a formal recommendation to the governor-general to reward the pueblos of Biñán and Sta. Rosa jointly with the lofty status of a villa with the title "siempre fiel y leal" ». Santiago cite les Archives nationales (Sinupang Pambansa) en note xl.
- Petite précision : « the title » plutôt que « the motto ». Action : remplacer.

**19. Photo de l'Alberto Mansion (L297).** C'est un crédit d'image (Bengoeswhere, CC BY-SA 3.0). À garder.

**20. Légende de l'Alberto Mansion (L297).**
- Phrase : « The historic Alberto Mansion in Biñan, originally built in the early 19th century, was the ancestral home of José Rizal's mother, **Teodora Alonso**. »
- Le blog wixsite est un blog personnel.
- Vérifié :
  - GMA News (19/10/2012) : « It is the home where Teodora Alonso, Rizal's mother, grew up… The two-storey Alberto House, owned by the family of the father of Rizal's mother Teodora Alonso… the 200-year-old ancestral home of Rizal's mother ».
  - Experience Biñan (office du tourisme de la ville) : le toit de *tisa* fait penser qu'elle fut bâtie « perhaps two hundred years ago » ; « first owned by Don Lorenzo Alonso de Alberto », père de Teodora.
- Phrase proposée : « The historic Alberto Mansion in Biñan, about two hundred years old, was the house of the family of José Rizal's mother, **Teodora Alonso**, where she grew up. »

**21. Encadré « ever faithful and loyal » (L299).** Même source que le n° 18 pour la *villa*. Pour Bautista, la fiche du NHCP dit : « MAY-AKDA NG "ACTA DE LA PROCLAMACION DE LA INDEPENDENCIA DEL PUEBLO FILIPINO" NA KANYANG BINASA SA HARAP NG MADLA, KAWIT, CAVITE, 12 HUNYO 1898 » et « ISINILANG SA BIÑAN, LAGUNA, 7 DISYEMBRE 1830 ». Action : remplacer.

**22. Encadré « Rizal put Biñan's church in his novel » (L307, sans source).**
- Phrase : « In *Noli me tangere* (1887), the old philosopher Tasio asks why, "in the year 70, when lightning struck in Biñan", it fell precisely on the church tower, destroying the clock and an altar. Rizal was then nine years old and a pupil in the town. »
- Texte espagnol (édition Sempau, Gutenberg n° 47584, chap. XIV « Tasio el loco ó el filósofo ») : « Y decidme, ¿por qué el año 70 cuando cayó un rayo en Biñan, cayó precisamente en la torre y destrozó reloj y un altar? ¿Qué hacía la esquila de Santa Bárbara? »
- Traduction de Derbyshire (Gutenberg n° 6737, chap. XIV « Tasio: Lunatic or Sage ») : « Explain to me why in the year '70, when lightning struck in Biñan, it hit the very church tower and destroyed the clock and altar. What was the bell of St. Barbara doing then? »
- Tout est exact. Le texte espagnol dit « la torre » ; c'est Derbyshire qui précise « church tower ». Rizal, né en juin 1861, avait 9 ans en 1870. Action : ajouter les deux citations.

### Période américaine

**23. Traité de Paris (L315).** ChanRobles est un portail juridique commercial. Le Projet Avalon (Yale) donne le texte : « The United States will pay to Spain the sum of twenty million dollars ($20,000,000) within three months after the exchange of the ratifications of the present treaty. » Action : remplacer.

### Seconde Guerre mondiale

**24. Capinpin à Bataan (L393).**
- Phrase : « On Bataan, Brigadier General **Mateo Capinpin**, a man who would later make Biñan his home, commanded the Philippine Army's 21st Division until the surrender of April 1942 and the Death March that followed. »
- Morton, ch. XXIV : le 5 avril 1942, « It was Easter Sunday » ; « the 21st Division, led by Brig. Gen. Mateo Capinpin » ; vers 16 h 30, le poste de commandement est pris et « General Capinpin, the division commander, was not among those who reached safety; he had become separated from his staff during the disorganized flight and been captured by the Japanese. »
- Erreur : il n'a pas commandé jusqu'à la reddition (9 avril). Sa présence dans la Marche de la mort n'est pas documentée.
- Phrase proposée : « On Bataan, Brigadier General **Mateo Capinpin**, a man who would later make Biñan his home, commanded the Philippine Army's 21st Division until he was captured by the Japanese on 5 April 1942, four days before the surrender. »

**25. Rapport des Hunters-ROTC, décembre 1944 (L405).**
- Phrase : « In December 1944 the Hunters-ROTC placed its 47th Division in Manila, Cavite and Batangas, with its **48th Regiment** in Cavite, just across the hills from Biñan. Its 45th Regiment, in Laguna, counted only 80 armed fighters – but some 5,000 "home guards", civilians who supplied food, messages and information. »
- Le blog transcrit un document d'archive. Sa note donne la cote : « Box 246, Entry 1087, Philippine Archive Collection, Record Group 407, (US) National Archives ». Le rapport est signé par Eleuterio Adevoso (« Terry Magtangol ») ; effectifs « As of 20 Nov 1944 ». J'y ai vérifié : 45ᵉ régiment (Laguna) : « Combat: 80 men under arms. Home Guards: Approx. 5,000 able-bodied » ; 47ᵉ division « Composed of Manila, Cavite and Batangas Troops » ; 48ᵉ régiment « Cavite Province ».
- La NARA ne met pas ce rapport en ligne. Sa liste des dossiers de guérilla (déjà citée au même paragraphe) recense les unités de la 47ᵉ division et du 48th Special Regiment.
- Citation proposée : `[Adevoso, Commander’s report on the Hunters-ROTC, Dec. 1944, RG 407, Entry 1087, Box 246 | U.S. National Archives](https://www.archives.gov/research/military/ww2/philippine/guerrilla-list-2.html)`. On peut garder le lien du blog en complément, comme « transcription ».
- Petite précision : le 45ᵉ régiment faisait partie de la **44ᵉ division** (Rizal, Laguna, Tayabas), pas de la 47ᵉ.

**26. Marcial Alintanahin (L409) – cas spécial.**
- Phrase : « A local account also names a former municipal president, Marcial Alintanahin, as a field commander, but no other source confirms it. »
- Le blog EL FILIPINISMO (Pepe Alas) dit : « Led by Commander Marcial Alintanahin (two-time Municipal President), the Hunters stormed Biñán… » Il ne cite aucune source.
- Nouveau : trois photos de Judgefloro (CC0, 7 mai 2018, Wikimedia Commons, fichiers « 6361Biñan City, Laguna Museum Old Town Hall Landmarks 08 » à « 10 ») montrent une vitrine du musée de l'ancien hôtel de ville. On y voit le livre de Mojica, *Terry's Hunters*, et une plaque « PAGKILALA » de la « Hunters ROTC Association, Inc. ». J'ai agrandi la plaque. On lit : « Hunter MARCIAL ALINTANAHIN » ; « …ipinagkakaloob ngayong ika-25 ng Abril, 1971 sa ika-29 na Taong Pagkakatatag, Pagtitipun-tipon at Kombensiyon ng Hunters ROTC Ass'n, Inc. na ginanap sa […]nino's Resort, San Pedro, Laguna » ; signée « Leonardo L. Cabañero, Pangulo ».
- Ce que cela prouve : Marcial Alintanahin était membre des Hunters-ROTC, honoré par leur association en 1971. Ce que cela ne prouve pas : son commandement le 3 février 1945 et ses deux mandats de président municipal.
- Le caveat actuel est honnête, mais il devient trop fort (« no other source »). Phrase proposée : « A local account names a former municipal president, Marcial Alintanahin, as the commander who led the attack; a 1971 plaque of the Hunters-ROTC Association honouring "Hunter Marcial Alintanahin", now in the city museum, confirms that he belonged to the guerrilla, but no document yet confirms his role on 3 February or his terms as municipal president. »
- Citations : garder le blog et ajouter `[Hunters ROTC Association plaque for “Hunter Marcial Alintanahin”, 1971, Biñan City Museum | Wikimedia Commons](https://commons.wikimedia.org/wiki/File:6361Bi%C3%B1an_City,_Laguna_Museum_Old_Town_Hall_Landmarks_08.jpg)`.
- Pistes : la liste des maires (Commons a une galerie « Mayors of Biñan City, Laguna ») ; un Marcelo Alintanahin apparaît parmi les insurgés de Cabuyao en 1935 (*Philippine Press Clippings*, voir `guerre.md`). Pas de trace de « Alintanahin » dans les métadonnées d'Internet Archive.

**27. Représailles à Los Baños et Calamba (L413).**
- Phrase : « In Los Baños and its barrios some 1,500 civilians were killed, and at Calamba a U.S. military commission later charged Japanese officers with the murder of "about one thousand" unarmed civilians between 11 and 24 February 1945. »
- Le blog transcrit l'acte d'accusation *U.S.A. v. Mikio Taneichi, Yuzo Sakata, Taichi Yamada and Bunji Kanto* (3 mai 1946) : à Calamba, du 11 au 24 février 1945, « Horacio HABANA, Jacinta MAYBITUIN and about one thousand (1,000) other unarmed, non-combatant Filipino civilians ». Le blog dit que l'original est sur Internet Archive, mais je ne l'ai pas retrouvé.
- Sources fiables lues :
  - Whitney, *The Case of General Yamashita: A Memorandum* (GHQ SCAP, vers 1950 ; Bibliothèque du Congrès), annexe C, « Review of Trial Record », p. 64–65 : « On six different occasions from 21 February to 6 March 1945, Japanese officers, soldiers and military police… killed by bayoneting, a total of about 264 men, women and children from different barrios of Los Banos… More than 2,500 men, women and children of Calamba were killed by bayoneting or burning on a single day, 12 February 1945 ». Plus loin : « At Calamba, Laguna, in February 1945, the killings were "by order of the Army" because the people were "anti-Japanese" ».
  - Henderson 2015 (déjà cité dans le chapitre) cite le rapport d'enquête n° 214 : « An estimated 1,500 citizens of Calamba… were massacred between February 12-15, 1945 » ; et des massacres de familles à Maajas (45 morts) et Tadlak (65 morts), barrios de Los Baños.
- Erreur : rien ne soutient « some 1,500 civilians » à Los Baños. Ce chiffre semble concerner Calamba.
- Phrase proposée : « The Japanese took their revenge on those who remained. In the barrios of Los Baños, soldiers killed hundreds of civilians in the weeks after the raid. At Calamba the toll was heavier still: U.S. war crimes investigators estimated that some 1,500 people were massacred between 12 and 15 February 1945, and the evidence heard at General Yamashita's trial put the dead of 12 February alone at more than 2,500. »
- Citations : Whitney (LoC) + Henderson. Si l'on garde l'acte Taneichi, le citer par sa référence d'archive (commission militaire américaine, Manille, 1946) avec la transcription en complément.

**28. Encadré « A Bataan general rests in Biñan » (L463).**
- Phrase : « Brigadier General Mateo Capinpin, who commanded the 21st Division on Bataan in 1942, spent his last years in Biñan, where he died in 1958; a street of the town bears his name. »
- Dans Wikipédia, la phrase « He died from heart attack at his residence in Biñan… on December 28, 1958, and was buried at the Biñan Municipal Cemetery » n'a **pas de référence**. Je n'ai trouvé que Find a Grave, un blog philatélique et des copies de Wikipédia. L'Inquirer USA et Rappler étaient bloqués.
- Vérifié : la rue existe. Fiche NHCP du marqueur « Rizal in Biñan » (1948) : « Location: Gen. M. Capinpin Street, Biñan, Laguna ». Le commandement de la 21ᵉ division à Bataan est dans Morton (n° 24).
- Phrase proposée : « **A Bataan general is remembered in Biñan.** Brigadier General Mateo Capinpin, who commanded the 21st Division on Bataan in 1942, gave his name to a street of the Poblacion, General M. Capinpin Street. » Ajouter « he died in Biñan in 1958 » seulement après vérification (acte de décès, journal de décembre 1958, tombe au cimetière municipal).

### Après-guerre

**29. L'ancienne casa hacienda devient hôtel de ville (L471).**
- Phrase : « In **1954–55** the old Dominican *casa hacienda* on the plaza, which the town had bought in 1910 to house its public school, was repaired, given a portico and opened as the **municipal hall on 15 May 1955** – San Isidro's feast day – under Mayor Jesus M. Garcia, while the school moved to a new building in Santo Domingo. »
- Problème 1 : l'URL citée (avec « biñan ») renvoie **404**. La bonne adresse est `https://experiencebinan.com/histo-cultural/sentrong-pangkultura-ng-binan/`.
- Problème 2 : la bonne page dit seulement : « During the American occupation, the building served as a school until the completion of the school building's construction in Barangay Santo Domingo, immediately after which it served as the new municipal hall… By this time, a portico (porch) had been added… In 2016… Sentrong Pangkultura ng Biñan ». Aucune date, aucun maire.
- Selon `apres-guerre.md`, 1954, le 15 mai 1955 et le maire Jesus M. Garcia viennent d'une notice de la Filipinas Heritage Library (`biblio/17972`). Cette page ne s'affiche pas ici (page vide). Je ne peux donc pas vérifier.
- Bon point : la contradiction avec l'achat de 1910 (chapitre américain) a déjà été corrigée dans cette version.
- Phrase proposée (si la notice FHL n'est pas confirmée) : « In the 1950s the old Dominican *casa hacienda* on the plaza, which the town had bought in 1910 to house its public school, was repaired, given a portico and turned into the municipal hall, once the school had moved to a new building in Santo Domingo. »
- Autre point : « the census of 1948 counted 20,794 people » (même phrase) n'a pas de source vérifiée (voir n° 33).

**30. Biñan, la municipalité la plus riche (L487).**
- Phrase : « Under his successor, Hermis Perez (1998–2007), Biñan became one of the wealthiest towns in the country: government statistics for 2008 ranked it first among all Philippine municipalities by income. »
- ntucphl.org (National Trade Union Center) reproduit un article de presse d'octobre 2010. Texte : « This, according to the National Statistical Coordination Board (NSCB)… Topping the list of the high-income towns was Biñan, Laguna… "Would you believe that nine of the 10 municipalities with the highest total income in 2008 come from Region 4-A or Calabarzon?…" NSCB Secretary-General Dr. Romulo A. Virola said in his online column, Statistically Speaking. »
- L'original (chronique de la NSCB) n'est plus en ligne sur le site de la PSA à ma connaissance. Experience Biñan (site officiel) écrit : « Biñan, the richest municipality in the Philippines before its cityhood ».
- Action : garder, mais renommer la citation pour dire clairement qu'il s'agit d'une reprise de l'article de presse sur les chiffres de la NSCB, et ajouter Experience Biñan. Remarque : en 2008, la maire était Len Alonte (élue en 2007) ; la phrase lie le classement de 2008 au mandat de Perez. On peut écrire « by the end of his terms ».

**31. Recensement de 2024 (L499).** City Population cite la PSA. Chiffres vérifiés : Biñan 407 437 (2020) et 584 479 (2024) ; Calamba 539 671 et 575 046 ; Santa Rosa 414 812 en 2020 (Biñan était bien 3ᵉ). Le site de la PSA (et de son bureau régional) bloque l'accès ici. La page Lawphil de la Proclamation 973 ne contient que le texte, sans le tableau par ville. Action : garder City Population en complément de la proclamation, ou le remplacer par le tableau de la PSA vérifié depuis un navigateur.

**32. Encadré sur l'hôtel de ville (L533).**
- Phrase : « The town bought it in 1910 for its school; on 15 May 1955 it became the municipal hall; and since 2016 it has been the Sentrong Pangkultura ng Biñan, home of the city museum. »
- Corriger l'URL (voir n° 29). « 15 May 1955 » n'est pas sur la page. Phrase proposée : « The town bought it in 1910 for its school; after the war it became the municipal hall; and since 2016 it has been the Sentrong Pangkultura ng Biñan, home of the city museum. »

**33. Encadré « From third to first » (L537).**
- Phrase : « Four years later the census counted 584,479 people, ahead of Calamba (575,046): Biñan had become the most populous place in the province. In 1948 it had just 20,794 inhabitants. »
- 2024 : vérifié. 1948 : la page de City Population commence en 1990. Le chiffre de 20 794 vient de Wikipédia (modèle tiré de Wikidata, d'après la PSA). Le recensement de 1960 est sur Internet Archive (`1960CPHPhilippines`), mais son OCR est inutilisable : il faudrait lire la page de la Laguna à l'œil.
- Action : garder la phrase sur 2024 ; retirer « In 1948 it had just 20,794 inhabitants » ou la citer au recensement de 1948 (ou au tableau historique de la PSA) après vérification.

**34. Encadré « Rizal lost his head – and got it back » (L541, sans source).**
- Phrase : « The monument to José Rizal on Biñan's plaza, put up in 1918 and believed to have been designed by the future National Artist Guillermo Tolentino, is one of the oldest Rizal monuments in the country. On 22 September 2015 lightning struck it and decapitated the statue. A replica by Frederic Caedo took its place on the plaza in 2016, and the restored original was unveiled at City Hall on 18 December 2017. »
- Vérifié :
  - Philstar (24/09/2015) : la maire Len Alonte Naguiat écrit que « the century-old Jose Rizal Monument in Biñan Plaza (believed to be sculpted by National Artist Guillermo Tolentino) was struck by lightning during the heavy rainfall ». Coconuts Manila précise « during a heavy downpour on Tuesday » (le mardi 22 septembre 2015). Attention : dans Philstar, la maire dit « Yesterday » dans un message « on Tuesday », ce qui donnerait le 21. La ville et Coconuts disent le 22.
  - Experience Biñan, « Rizal Monument » : « on September 22, 2015… a lightning decapitated the statue. Its wreckage is kept in the City Museum. In 2016… Frederic "Fred" Caedo created a replica, which found its way to the top of the monument. Caedo also restored the blasted original work, and the restored Rizal statue now stands in its new home across the City Hall of Biñan in Barangay Zapote. » La page « Jose Rizal in Biñan » ajoute : « The monument was finished on May 14, 1918 ».
  - Ville de Biñan (déc. 2017) : « unveiled the monument of our National Hero, Dr. Jose P. Rizal on December 18, 2017 at Biñan City Hall » ; le maire Dimaguila : « in-unveil natin ang orihinal na rebulto ni Dr. Jose Rizal na dating nasa Plaza, ni-repair natin 'yun ».
- Non soutenu : « one of the oldest Rizal monuments in the country ». « Designed » : les sources disent « sculpted ».
- Phrase proposée : « The monument to José Rizal on Biñan's plaza, completed on 14 May 1918 and believed to have been sculpted by the future National Artist Guillermo Tolentino, is more than a century old. On 22 September 2015 lightning struck it during a heavy downpour and decapitated the statue. A replica by Frederic Caedo took its place on the plaza in 2016, and the restored original was unveiled at City Hall on 18 December 2017. »
- Remarque : la page de la ville cite le journal de Rizal avec « I left Biñan on the 17th of December 1871 … It was a Saturday ». Le 17 décembre 1870 était un samedi ; le livre a raison de garder 1870.

### Transports (nouveau chapitre)

**35. SLEX en 1978 (L621).**
- Phrase : « The South Luzon Expressway reached Biñan in 1978, when its extension from Alabang to Calamba was opened. »
- Wikipédia : inauguration par Marcos le 4 mai 1978, d'après le *Fookien Times Philippines Yearbook* 1978 (imprimé) et un rapport de la JICA.
- Le rapport JICA (p. 142, page lue en image) dit : « In 1974, PNCC simultaneously started working on the 30 kilometer SLE Extension from Muntinlupa to Calamba, Laguna ». Il ne donne pas la date d'ouverture. Le PD 1113 (1977) accorde la concession et parle de « the construction of Manila North and Manila South Expressways including the extension of the Manila South Expressway to Calamba, Laguna ».
- Phrase proposée : « The South Luzon Expressway reached Biñan in the late 1970s, when its extension from Alabang to Calamba, begun in 1974, was completed; it was inaugurated in May 1978. » Garder « May 1978 » seulement avec la référence imprimée du *Fookien Times Yearbook* (1978).

**36. CALAX (L621).**
- Phrase : « Since 2019 the **Cavite–Laguna Expressway** (CALAX) has started at the Mamplasan roundabout, at the Greenfield toll plaza; an exit serving the Laguna Technopark opened in Timbao in 2020. »
- Vérifié sur le site officiel de l'exploitant (déjà cité au même paragraphe) :
  - « CALAX toll collection » (2020) : « The first 10.7-km has served about 10,000 vehicles daily between SLEX Mamplasan Exit and the new CALAX Santa Rosa-Tagaytay Interchange… Partial operations started October 30 in time for "Undas." »
  - « DPWH inspects CALAX Laguna segment interchanges » (2020) : « we are ready to open these interchanges to the public at 12:01AM on August 18, 2020… access to and from the Laguna Technopark ».
- Non vérifié : le nom « Greenfield toll plaza » (le site parle de « Mamplasan Toll Barrier »). Action : remplacer. Précision possible : « Since 30 October 2019… on 18 August 2020 ».

### Économie (nouveau chapitre)

**37. Revenu ordinaire de la ville (L701).** PhilAtlas reproduit les chiffres du Bureau of Local Government Finance : « 2009 606,915,277.04 … 2016 1,530,342,881.53 » ; « Data on annual regular income is from the Bureau of Local Government Finance, Department of Finance ». Le site du BLGF bloque l'accès ici (pare-feu Azure). Action : garder, ou remplacer par la page « LGU Fiscal Data » du BLGF après vérification.

---

## Sources gardées sans changement (pour mémoire)

- **Experience Biñan** (`experiencebinan.com`) : c'est le site de l'office municipal de la culture et du tourisme (« Copyright © 2022 Biñan City Culture, History, Arts & Tourism Office »). C'est une source officielle, mais touristique : elle donne souvent la tradition locale (Tubigan et le « patubigan » de Domingo Lamco ; *pospas* ; *puto* ; *maglalatik*). Les liens « fee-binan/maglalatik » redirigent vers « feel-binan/maglalatik » : on peut corriger le lien.
- **OpinYon, Journal News, Daily Tribune, LB Times, DZRH, BusinessMirror, GMA, Philstar** : presse ; à garder.
- **Nouveaux chapitres** : Bulatlat (média alternatif), Dugout.ph (média sportif), PortCalls (presse spécialisée), Top Gear Philippines (magazine) : acceptables pour des faits récents. Pour le stade de football (Dugout.ph), un communiqué de la fédération (PFF) serait meilleur.
- **Commons** (photo de l'Alberto Mansion) : crédit d'image.

## Autres anomalies repérées (hors périmètre)

1. **SEAFDEC, LakeCon 2003 (Géographie, L25)** : le lien `repository.seafdec.org.ph/handle/10862/6120` mène à « The ichthyofauna of Lake Manguao, Taytay, Palawan », pas à un texte sur les poissons de Laguna de Bay. L'erreur est déjà dans `geographie.md`. La recherche dans le dépôt SEAFDEC était bloquée ici.
2. **Faille de la Vallée Ouest (L15)** : le rapport JICA MMEIRS (2004), tableau 2.1.1, donne « West Valley Fault… 7.2… 1658.08.19 ». C'est une meilleure source que DZRH pour M 7,2 et 1658.
3. **Vitales 2013 (Préhistoire)** : `journals.upd.edu.ph` renvoyait 404 le 16/11/2026, même pour la page d'accueil de la revue. C'est peut-être une panne. À revérifier.
4. **Proclamation 973 (Après-guerre, L499)** : la page Lawphil ne contient pas le tableau des chiffres. Elle prouve que les chiffres sont officiels, pas les chiffres eux-mêmes.
5. **Ma-i en 971 (L173)** : voir n° 15 (Hirth et Rockhill disent 982).

## Sites bloqués ou en panne depuis cet environnement (16/11/2026)

JSTOR (défi anti-robot), Inquirer et Inquirer USA, Rappler, Manila Bulletin, UNESCO, PSA et son bureau régional IV-A, BLGF, LLDA, PPP Center, UPLB, SunStar, EurekAlert (403 en fin de journée), Wiley, ScienceDirect, Springer, ADS, HathiTrust, Google Books (quota), Biblioteca Digital de Castilla y León, web.archive.org (connexion coupée), philippinestudies.net (injoignable). Le site web du GVP affichait « temporarily unavailable » (j'ai utilisé son service de données). Project Gutenberg était en maintenance (j'ai lu ses fichiers sur Internet Archive).

## Suivi (16/11/2026, fin de journée)

Toutes les actions « remplacer », « corriger » et « nuancer » de ce tableau ont été appliquées dans `history.md` (anglais et tagalog), dans le compagnon et dans les annexes. Compléments trouvés depuis :
- **SEAFDEC (n° 12 des anomalies)** : la bonne référence est Bagarinao 2001, https://repository.seafdec.org.ph/handle/10862/834. Son résumé confirme les 95 % d'*ayungin*, *biyang puti* et *kanduli* au début des années 1960.
- **1948 (n° 33)** : le chiffre de 20 794 est confirmé par le tableau officiel « Population Enumerated in Various Censuses by City/Municipality: 1903 to 2007 » (NSO, recensement de 2007, fichier Calabarzon), archivé sur Internet Archive : https://archive.org/details/PhilippinesCensusofPopulationLGUs19032007. Le même tableau confirme tous les chiffres de 1903 à 2007 de l'annexe C.
- **Hôtel de ville (n° 29 et 32)** : la galerie des maires du musée confirme que Jesus M. Garcia était maire de 1952 à 1959 (voir `hdp.md`). La date du 15 mai 1955 reste à vérifier sur la notice de la Filipinas Heritage Library (le site renvoie une page vide aux scripts).
- **Herrero 1900** : vérifié sur le texte d'Internet Archive (p. 106–116) pour Biñan en juin 1898 (Pedro de León, « Palañac »).

