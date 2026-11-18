# Note de recherche : l'économie de Biñan aujourd'hui, secteur par secteur (2026)

*Note pour le livre bilingue (anglais/tagalog) sur Biñan City, Laguna. Rédigée le 16/11/2026. Elle prépare une section de chapitre avec cartes. L'auteur demande : « what are the big companies by sector » et « what about agriculture, livestock (I saw pig farms near the lake), fishing ». Je ne refais pas l'histoire de la Laguna Technopark, du LIIP, de Southwoods et de l'industrialisation des années 1990 (voir `apres-guerre.md` § 2.3), ni la question des propriétaires fonciers et des surfaces des grands ensembles (voir `foncier.md`). Les transports et la culture (fêtes, puto, etc.) sont traités par d'autres notes : je ne les cite que lorsqu'ils touchent l'économie (par ex. la route-digue du lac et les pêcheurs).*

**Légende des niveaux de fiabilité** (même grille que les autres notes)
- ✅ **Solide** : source primaire lue (proclamation, décision de la Cour suprême, base de données PSA ou DTI, site officiel, document d'une banque de développement), ou presse nationale concordante.
- 🟡 **À nuancer** : source secondaire sérieuse non lue en entier, annuaire d'entreprises, offre d'emploi, point OpenStreetMap, site de la ville sans document joint, chiffres contradictoires.
- 🔴 **Fragile** : blog, affirmation non sourcée, non vérifiée ou contredite.

**Méthode.** J'ai interrogé par script : la base **CMCI** du DTI (profil de Biñan, années 2014 à 2024) ; l'API **OpenSTAT** de la PSA (recensement agricole 2012 par commune, recensement 2022 par province, inventaire porcin de la Laguna 2010–2026) ; l'API WordPress du site de la ville **binan.gov.ph** (articles sur la pêche, l'agriculture, les malls) ; le site **Experience Biñan** (fiches des 24 barangays). J'ai lu en plein texte sur LawPhil **19 proclamations PEZA** et la décision de la Cour suprême **G.R. No. 200403 (2022)** sur les élevages de porcs. J'ai lu l'étude **PIDS 2008** sur les enclos à poissons, les résolutions **LLDA 540 (2018) et 561 (2019)**, la fiche **AIIB** de la route-digue du lac (LLRN). J'ai géocodé une soixantaine de lieux avec **Nominatim (OpenStreetMap)** et je les ai placés dans les polygones de barangays du dépôt (`book/maps/data/bgysubmuns-municity-403403000.topo.0.1.json`). Licences Commons vérifiées par l'API.
**Bloqués** : Inquirer, Manila Bulletin, PCO (403) ; site du PEZA (503 et erreur TLS) ; dti.gov.ph (hors réseau autorisé) ; bai.gov.ph (erreur TLS) ; API Overpass d'OpenStreetMap (connexion coupée : je n'ai donc pas pu extraire les polygones d'occupation du sol). Les articles bloqués ne sont cités que d'après les extraits des moteurs de recherche (🟡).

---

## Les dix découvertes principales

1. **Biñan est n° 1 du pays pour le « dynamisme économique » parmi les villes de rang « component city ».** ✅ Base CMCI du DTI : 7ᵉ ville « component » la plus compétitive en **2024** (score 47,80), mais **1ʳᵉ en Economic Dynamism** en 2022, 2023 et 2024, et **1ʳᵉ en Innovation** en 2024. Biñan est 1ʳᵉ pour « Employment Generation », « Productivity » et « Getting Business Permits ». Elle est faible en infrastructures (40ᵉ) et en efficacité du gouvernement (38ᵉ). En 2014, elle était 136ᵉ.
2. **Environ la moitié de la Laguna Technopark est dans Biñan.** ✅ (calcul) La somme des **13 proclamations** PEZA que j'ai lues pour la Technopark dans Biñan donne **≈ 239 ha** (`foncier.md` donne ≈ 230 ha avec une liste un peu plus courte). La Technopark fait 471 ha en tout (ALLHC). Les grandes usines du côté de Biñan sont surtout dans les barangays **Biñan (« Biñang »), Loma, Malamig, Timbao et Mamplasan**. Honda Cars et Mitsubishi Motors sont, eux, du côté de **Santa Rosa** (Don Jose).
3. **Les plus gros employeurs sont l'électronique et le stockage informatique.** 🟡 Western Digital/HGST (têtes de lecture ; « over 10,000 » employés selon le gouvernement), Nidec (moteurs de disques durs ; ≈ 8 000 employés en 2020), Toshiba (disques durs et SSD ; usine fondée en 1995, en expansion en 2026), IMI (siège mondial), Amkor, Cirtek, TDK. Puis l'automobile (Isuzu, Honda Parts, Furukawa, Joyson), le médical (Terumo, Unilab/Amherst), l'alimentaire (Gardenia).
4. **Biñan n'a presque plus d'agriculture.** ✅ Recensement agricole PSA 2012 : **118 exploitations, 135 ha** pour toute la ville (≈ 3 % de sa surface). Une étude de 2015 (citée par la presse) dit que la part des terres agricoles passe de **86,9 % en 1979 à 9,13 % en 1999**, et celle du bâti de 10,5 % à 70 %. 🟡
5. **Les « porcheries » de Biñan sont allées jusqu'à la Cour suprême.** ✅ Une ordonnance de 2004 ordonne la fermeture progressive des grands élevages (plus de 10 porcs ou 500 volailles). Deux élevages, **Holiday Hills Stock & Breeding Farm** et **Domino Farms**, l'attaquent. La Cour suprême la déclare « NOT UNCONSTITUTIONAL » le **10 octobre 2022**. En janvier 2025, le maire annonce que les deux élevages n'ont plus de permis. Ils sont à **Langkiwa et Santo Tomas**, **loin du lac** (à l'ouest de la ville).
6. **Les porcs que l'auteur a vus près du lac ne sont documentés nulle part.** 🔴 Je n'ai trouvé aucune source sur des porcheries à Malaban ou De La Paz, ni d'action de la LLDA visant Biñan. Ce sont peut-être de petits élevages de basse-cour (*backyard*), ou des élevages des villes voisines. À vérifier sur place (barangay, City Veterinarian).
7. **La peste porcine africaine (ASF) a fait chuter le cheptel de la Laguna de 69 % en un an.** ✅ PSA : **194 394 porcs au 1ᵉʳ janvier 2020, 60 258 au 1ᵉʳ janvier 2021**. Remontée à 130 206 en 2026. Aucun cas propre à Biñan trouvé dans la presse.
8. **En 2006, Biñan avait le plus d'opérateurs d'enclos et de cages à poissons de la province de Laguna.** ✅ LLDA (via PIDS 2008) : **26 opérateurs d'enclos (*fishpens*) sur 650 ha et 76 opérateurs de cages sur 35 ha**, soit 102 opérateurs et 686 ha (San Pedro : 91 opérateurs, mais 836 ha). Dans les enclos de Biñan, on élève surtout le **bangus** (milkfish) ; dans les cages, presque uniquement la **carpe**.
9. **La rive du lac fait environ 4 km, à Malaban et De La Paz.** ✅ (calcul sur la carte du dépôt) Ce sont les deux barangays de pêcheurs. La ville y organise la « Bangkarera sa Lawa » (course de bateaux) et des nettoyages ; elle a fait prêter serment à 25 gardes-pêche en 2018. La future route-digue du lac (LLRN, tronçon San Pedro–Cabuyao, 12,23 km) touchera des enclos et des cages ; le DPWH cite Malaban, De La Paz, Casile et San Antonio.
10. **Le commerce : un marché, trois malls, 7 500 entreprises.** ✅/🟡 Le marché public de Poblacion (« the largest public market in … CALABARZON », selon la ville, ouvert 24 h/24). Pavilion Mall (1999 🟡), Biñan Central Mall (pré-ouverture le **9 sept. 2011**), Southwoods Mall (**30 sept. 2017**, ₱2 milliards, 58 000 m²). **7 524 entreprises enregistrées en 2021** ; « our top 100 taxpayers contribute 95 percent of all the taxes » (le maire, 2022).

---

## 1. Sources principales

| Abrév. | Référence | URL | Statut |
|---|---|---|---|
| **CMCI** | DTI, *Cities and Municipalities Competitiveness Index*, profil « City of Binan », années 2014–2024 (rangs et scores par pilier) | https://cmci.dti.gov.ph/lgu-profile.php?lgu=Binan&year=2024 (changer `year=` pour les autres années) | ✅ lu par script |
| **OpinYon 2022a** | J. Veloso et C. Go, « Biñan City promotes inclusive economic growth », *OpinYon*, 24 oct. 2022 (7 013 → 7 524 entreprises ; top 100 = 95 % des taxes ; CMCI 2022) | https://opinyon.net/laguna/binan-city-promotes-inclusive-economic-growth | ✅ lu |
| **OpinYon 2025a** | « Significant growth », *OpinYon*, 24 févr. 2025 (BOSS 2025 : 296 nouvelles entreprises, 6 169 renouvellements, ₱588,2 M) | https://opinyon.net/national/significant-growth | ✅ lu |
| **BM 2023** | « Biñan, the City of Life, returns… », *BusinessMirror*, 20 févr. 2023 (budget ₱3,2 milliards ; prêt Landbank ₱3,15 milliards ; hôpital > 180 lits) | https://businessmirror.com.ph/2023/02/20/binan-the-city-of-life-returns-with-a-brighter-future-yet-to-unfold/ | ✅ lu |
| **PhilAtlas** | « Biñan, Province of Laguna » : revenus annuels réguliers 2009–2016 (d'après le BLGF) ; population 2020 par barangay (PSA) | https://www.philatlas.com/luzon/r04a/laguna/binan.html | ✅ lu (reprise de données officielles) |
| **WP Biñan** | Wikipédia, « Biñan », sections *Economy*, *Major industries*, *Shopping centers*, *Investment sites* ; infobox (revenus 2022 et 2024) | https://en.wikipedia.org/wiki/Bi%C3%B1an | 🟡 (liste d'entreprises qui mélange Biñan et Santa Rosa) |
| **EB** | « About Biñan », Experience Biñan (office culturel de la ville) : fiches des 24 barangays (surface, usage du sol, activités) | https://experiencebinan.com/about-binan/ | 🟡 site de la ville, sans sources |
| **Mairie Hist.** | Page « History » de binan.gov.ph : Pavilion Mall, Biñan Central Mall (pré-ouverture 9 sept. 2011), « largest public market » | https://www.binan.gov.ph/history/ | 🟡 auto-description |
| **Mairie 2017** | « Southwoods Mall Grand Opening », binan.gov.ph, 30 sept. 2017 (₱2 milliards, 58 000 m², 3 niveaux, 4 cinémas) | https://www.binan.gov.ph/southwoods-mall-grand-opening/ | ✅ lu |
| **Mairie ELA** | « Executive Legislative Agenda 2018 », binan.gov.ph (point 11 : « Finish the CLUP ») | https://www.binan.gov.ph/executive-legislative-agenda-2018/ | ✅ lu |
| **Procl. PEZA** | 19 proclamations PEZA concernant Biñan, 1997–2025 (liste au § 2.2) | https://lawphil.net/executive/proc/ (URL de chaque texte au § 2.2) | ✅ lues |
| **ALLHC 2020** | « Laguna Technopark celebrates 30 years » (471 ha, > 270 entreprises, > 100 000 emplois) | https://www.ayalalandlogistics.com/all-news/laguna-technopark-celebrates-30-years/ | ✅ (déjà dans `apres-guerre.md`) |
| **ALLHC 17-A** | AyalaLand Logistics, SEC Form 17-A 2025 et rapport intégré 2024 (ALogis Biñan 103 000 m² ; entrepôts frigorifiques Artico Biñan 1 et 2) | voir `foncier.md` § 1 | ✅ (lu pour `foncier.md`) |
| **TIP** | Toshiba Electronic Devices & Storage, « Toshiba Information Equipment (Philippines), Inc. (TIP) » : fondée le 7 juil. 1995, activité depuis le 1ᵉʳ oct. 1996 ; 103 East Main Ave. Ext., Laguna Technopark, Biñan ; 60 036 m² + annexe 21 000 m² | https://toshiba.semicon-storage.com/ap-en/company/about/omc/tip.html | ✅ lu |
| **SC 2022** | *Municipality of Biñan v. Holiday Hills Stock & Breeding Farm Corp. and Domino Farms, Inc.*, G.R. No. 200403, 10 oct. 2022 (ordonnance n° 06 de 2004) | https://lawphil.net/judjuris/juri2022/oct2022/gr_200403_2022.html | ✅ lu |
| **OpinYon 2025b** | « Biñan City LGU to finally close 2 erring hog yards », *OpinYon*, 13 janv. 2025 (Langkiwa et Santo Tomas ; EO n° 19 s. 2018) | https://opinyon.net/national/binan-city-lgu-to-finally-close-2-erring-hog-yards | ✅ lu |
| **PSA CAF 2012** | PSA OpenSTAT, « Number and Area of Holdings/Farms by Region, Province and City/Municipality: Philippines, 2012 » (table 0332E6CNAF1) | https://openstat.psa.gov.ph/PXWeb/api/v1/en/DB/2E/AFC/0332E6CNAF1.px | ✅ lu par script |
| **PSA CAF 2012 b** | Idem, livestock (table 0022E6CLFH0) : porcs de la Laguna en 2012 | https://openstat.psa.gov.ph/PXWeb/api/v1/en/DB/2E/AFC/0022E6CLFH0.px | ✅ lu par script |
| **PSA CAF 2022** | PSA OpenSTAT, tables 0812E6CFHO0 (ménages de pêcheurs), 0802E6CFOW0 (type d'eau), 0832E6CFNA0 (exploitations) : chiffres **par province** (pas par commune) | https://openstat.psa.gov.ph/PXWeb/api/v1/en/DB/2E/AFC/0812E6CFHO0.px | ✅ lu par script |
| **PSA Swine** | PSA OpenSTAT, « Swine: Inventory by Classification, Province, and Quarter, 2010-2026 » (table 0032E4FINS0) | https://openstat.psa.gov.ph/PXWeb/api/v1/en/DB/2E/LP/INV/NEW/0032E4FINS0.px | ✅ lu par script |
| **OpinYon 2022b** | « Land conversion: A threat to Laguna's agri sector? », *OpinYon*, 5 déc. 2022 (cite K. Engay-Gutierrez, Southern Luzon State University, 2015 : 86,9 % → 9,13 %) | https://opinyon.net/laguna/land-conversion-a-threat-to-laguna-s-agri-sector | 🟡 (étude de 2015 non lue) |
| **PIDS 2008** | D. C. Israel, « Fishpen and Fishcage Culture in Laguna de Bay: Status, Economic Importance… », *Philippine Journal of Development* n° 64, 2008 (enquête 2007 ; tableau 1 d'après la LLDA 2006) | https://pidswebs.pids.gov.ph/CDN/PUBLICATIONS/pidspjd08-1fishpen.pdf | ✅ lu |
| **LLDA BR 540** | LLDA, Board Resolution No. 540 s. 2018 (révision du ZOMAP de 1999 ; plafond de 9 200 ha d'enclos et de cages) | https://law.upd.edu.ph/wp-content/uploads/2020/07/LLDA-BR-No-540-Series-of-2018.pdf | ✅ lu |
| **LLDA BR 561** | LLDA, Board Resolution No. 561 s. 2019 (règles d'application ; « ZOMAG of 2019 ») | https://law.upd.edu.ph/wp-content/uploads/2020/08/LLDA-BR-No-561-Series-of-2019.pdf | ✅ lu |
| **AIIB LLRN** | AIIB, *Project Summary Information*, P000725 « Laguna Lakeshore Road Network Phase 1 », 10 déc. 2024 | https://www.aiib.org/en/projects/details/2024/_download/Philippines/AIIB-PSI-P000725-Philippines-LLRN-Laguna-Lakeshore-Road-Network-Phase-1-Project.pdf | ✅ lu |
| **Rappler 2024** | I. Gozum, « Laguna Lakeshore Road Network feared to cause more frequent flooding », *Rappler*, 14 sept. 2024 (témoin de Malaban ; plan de réinstallation de l'ADB : 1 084 familles) | lu sur la reprise https://mindanawonline.blogspot.com/2024/09/laguna-lakeshore-road-network-feared-to.html | 🟡 (lu sur une reprise ; l'original Rappler n'a pas été ouvert) |
| **Mairie pêche** | binan.gov.ph : gardes-pêche (22 nov. 2018) ; marquage SSEN des bateaux (25 juin 2026) ; nettoyage de Malaban (21 nov. 2025) ; Bangkarera sa Lawa (31 juil. 2024 et 2025) ; surimi à Santo Tomas (16 févr. 2024) | https://www.binan.gov.ph/mga-bagong-deputadong-fish-warden-nanumpa/ ; https://www.binan.gov.ph/isinagawa-ang-marking-of-safety-security-and-environmenral-numberingssen-system-para-sa-mga-rehistradong-bangka-sa-lungsod-ng-binan/ ; https://www.binan.gov.ph/4th-quarter-lakeshore-clean-up-drive-ang-binan-city-agriculture-office-kasama-ang-cfarmc-bantay-lawa-at-philippine-coast-guard/ ; https://www.binan.gov.ph/bangkarera-sa-lawa/ ; https://www.binan.gov.ph/bangkarera-sa-lawa-2/ ; https://www.binan.gov.ph/paggawa-ng-surimi-based-products-pinangunahan-ng-gad-office-city-of-binan/ | ✅ lus |
| **Nominatim** | OpenStreetMap (ODbL), géocodage des usines, malls, écoles, hôpitaux | https://nominatim.openstreetmap.org/ | 🟡 (contributeurs bénévoles ; positions bonnes à ± 100 m) |

Les autres sources (entreprises, presse) sont données ligne par ligne dans les tableaux du § 2.

---

## 2. Industrie et grandes entreprises

### 2.1 Où sont les usines ?

- ✅ Trois grands blocs industriels :
  1. **Laguna Technopark (LTI, groupe Ayala)**, au sud-ouest, à cheval sur Biñan et Santa Rosa. Dans Biñan : barangays **Biñan (« Biñang »), Loma, Timbao, Mamplasan** (proclamations) et **Malamig** (adresses d'entreprises ; campus DLSU).
  2. **Laguna International Industrial Park (LIIP)**, au sud, à **Mamplasan et Ganado** (1992 ; Samsung et Solid Corp. ; ≈ 117 ha dont ≈ 30 ha PEZA ; voir `apres-guerre.md` et `foncier.md`, qui signale une accusation de 1994 sur des terres de réforme agraire 🔴).
  3. Des **campus et zones informatiques** : Unilab Pharma Campus (Bungahan/Zapote), Southwoods (San Francisco), One Asia Business Center (San Antonio/Santo Niño), DLSU Innovation Hub (Malamig), Sevina Park (Biñan).
- ⚠️ **Biñan ou Santa Rosa ?** Les adresses disent souvent « Laguna Technopark, Biñan » ou « Laguna Technopark, Sta. Rosa ». J'ai placé chaque usine sur la carte des barangays du dépôt. Résultat : **IMI, Amkor, Toshiba, Western Digital, Nidec, TDK, Isuzu, Honda Parts, Kito, Katolec, Ryonan, Atlas Copco, Cirtek, Furukawa, Joyson, SunPower/Maxeon, Shin-Etsu, Terumo** tombent dans Biñan. **Honda Cars Philippines** et **Mitsubishi Motors Philippines** (Don Jose), **Coca-Cola** (Pulong Santa Cruz) et **Monde Nissin** (Balibago) tombent dans **Santa Rosa**. Wikipédia cite Honda, Toshiba, Diageo, etc. comme « employers » de Biñan sans faire la différence.
- ⚠️ **Les barangays à l'intérieur de la Technopark sont flous.** Les adresses OpenStreetMap disent « Malamig » pour Toshiba et Amkor, « Loma » pour Nidec, alors que la carte du dépôt les place dans le barangay **Biñan**. Les surfaces aussi divergent : Malamig fait 398 ha selon la ville, 52 ha sur la carte (voir `foncier.md` § 4). Pour le livre, écrire « in the Biñan part of Laguna Technopark », et ne donner le barangay que lorsqu'une proclamation ou l'entreprise le cite.

### 2.2 Les zones économiques PEZA de Biñan (proclamations)

| Zone | Procl. | Date | Surface | Barangay | URL | Statut |
|---|---|---|---|---|---|---|
| Laguna Technopark II SEZ | 1052 | 7 août 1997 | 67 ha | Biñan (« Binang ») | https://lawphil.net/executive/proc/proc1997/proc_1052_1997.html | ✅ |
| Laguna Technopark III SEZ | 1147 | 9 janv. 1998 | 96,39 ha | Loma, Timbao | https://lawphil.net/executive/proc/proc1998/proc_1147_1998.html | ✅ |
| Ajout à la Technopark | 51 | 28 mai 2001 | 4,40 ha (4 parcelles) | Biñan | https://lawphil.net/executive/proc/proc2001/proc_51_2001.html | ✅ |
| Ajout au LIIP-SEZ (depuis la zone industrielle générale) | 105 | 1ᵉʳ oct. 2001 | 2,90 ha | Ganado | https://lawphil.net/executive/proc/proc2001/proc_105_2001.html | ✅ |
| Ajout à la Technopark | 109 | 5 oct. 2001 | 3,13 ha | (Biñan, barangay non précisé) | https://lawphil.net/executive/proc/proc2001/proc_109_2001.html | ✅ |
| Ajout à Laguna Technopark I | 1143 | 20 sept. 2006 | 2,44 ha | Biñan | https://lawphil.net/executive/proc/proc2006/proc_1143_2006.html | ✅ |
| Ajout à Laguna Technopark I | 1201 | 28 déc. 2006 | 1,02 ha | Biñan | https://lawphil.net/executive/proc/proc2006/proc_1201_2006.html | ✅ |
| Ajout à Laguna Technopark III | 1399 | 9 oct. 2007 | 2,76 ha | Loma | https://lawphil.net/executive/proc/proc2007/proc_1399_2007.html | ✅ |
| Ajout au LIIP-SEZ | 1504 | 19 mai 2008 | 1,05 ha | Ganado | https://lawphil.net/executive/proc/proc2008/proc_1504_2008.html | ✅ |
| « Laguna Technopark Annex » | 1603 | 27 août 2008 | 29,02 ha | Biñan | https://lawphil.net/executive/proc/proc2008/proc_1603_2008.html | ✅ |
| Ajout à la Technopark | 1704 | 13 janv. 2009 | 4,19 ha | Loma | https://lawphil.net/executive/proc/proc2009/proc_1704_2009.html | ✅ |
| **One Asia Business Center** (IT Park) | 1773 | 14 mai 2009 | 4,80 ha | San Antonio | https://lawphil.net/executive/proc/proc2009/proc_1773_2009.html | ✅ |
| Ajout à la Technopark | 116 | 24 févr. 2011 | 8,08 ha | Mamplasan | https://lawphil.net/executive/proc/proc2011/proc_116_2011.html | ✅ |
| Ajout à la Technopark | 621 | 12 juil. 2013 | 14,24 ha | Mamplasan | https://lawphil.net/executive/proc/proc2013/proc_621_2013.html | ✅ |
| Ajout à la Technopark | 972 | 17 févr. 2015 | 4,53 ha | Mamplasan | https://lawphil.net/executive/proc/proc2015/proc_972_2015.html | ✅ |
| Ajout à la Technopark | 1134 | 1ᵉʳ oct. 2015 | 2,14 ha | Mamplasan | https://lawphil.net/executive/proc/proc2015/proc_1134_2015.html | ✅ |
| **Southwoods Mall** (IT Center ; bâtiment de 128 153,56 m² sur 1,92 ha) | 381 | 13 déc. 2017 | 1,92 ha | « Halang, Soro-Soro, San Francisco » | https://lawphil.net/executive/proc/proc2017/proc_381_2017.html | ✅ |
| **Sevina Park** (IT park, ArthaLand via Cazneau Inc.) | ? | 2024 ? | ? | Biñan ? | MB 2023 (en attente) : https://mb.com.ph/2023/12/27/5-peza-approved-ecozones-pending-proclamation ; liste 2024 : https://business.inquirer.net/497216/27-ecozones-proclaimed-under-marcos-administration | 🟡 (numéro non trouvé) |
| **De La Salle University Innovation Hub** (KIST) | 985 | 30 juil. 2025 | 5 ha | Malamig (campus Leandro V. Locsin) | https://lawphil.net/executive/proc/proc2025/proc_985_2025.html | ✅ |

- ✅ **Total Technopark dans Biñan** (13 proclamations ci-dessus, sans le LIIP) : **≈ 239 ha**. 🟡 La liste n'est peut-être pas complète (la proclamation d'origine de « Laguna Technopark I » n'a pas été trouvée) et je ne sais pas si toutes les parcelles sont bâties. Le site de LTI dit que la phase 1 (1989) faisait **224 ha** sur les deux villes (`foncier.md` 🟡).
- 🟡 **Southwoods Office Towers 1 et 2** : la presse (Philstar, 16 janv. 2018) parle d'une « Presidential Proclamation No. 383 ». Sur LawPhil, la n° 383 concerne SouthPark (Alabang, Muntinlupa) ; c'est la **n° 381** qui vise Southwoods (le bâtiment « Southwoods Mall », relié aux tours). https://philstar.com/business/2018/01/16/1778076/peza-grants-ecozone-starts-megaworlds-southwoods-office-towers. Les deux tours de 12 étages offrent ≈ 46 000 m², pour ≈ 8 000 postes BPO/IT (Inquirer, extrait 🟡 : https://business.inquirer.net/244177/megaworld-gets-peza-accreditation-two-southwoods-office-towers).
- 🟡 **Projet** : Unilab veut faire proclamer son campus de Biñan comme zone pharmaceutique (Inquirer, 10 mai 2025 : https://business.inquirer.net/524770/unilab-plans-to-transform-laguna-campus-into-ecozone).
- 🔴 **Proclamation d'origine du LIIP-SEZ** : non trouvée.

### 2.3 Les grandes entreprises, par secteur

Coordonnées : géocodage Nominatim (OpenStreetMap), précision ± 100 m. « Barangay (carte) » = barangay du polygone du dépôt où tombe le point ; « adresse » = barangay donné par l'entreprise ou OSM.

**A. Électronique, semi-conducteurs, stockage informatique**

| Entreprise | Activité | Lieu | Coord. (lat, lon) | Source | Statut |
|---|---|---|---|---|---|
| **Integrated Micro-Electronics, Inc. (IMI)** (groupe Ayala) | Sous-traitance électronique (EMS) ; **siège mondial** ; installée à la Technopark en 1995 (venue de Muntinlupa) | North Science Ave., Special Export Processing Zone, LTI, Biñan ; carte : Biñan | 14.2674, 121.0539 | Profil IMI 2023 : https://exhibitors.electronica.de/download/1121_11_2_1988_2_1_807/electronica_imi-company-profile-2023-revpdf.pdf ; https://en.wikipedia.org/wiki/Integrated_Micro-Electronics,_Inc. | ✅/🟡 |
| **Western Digital Storage Technologies Philippines** (ex-HGST/Hitachi) | Têtes de lecture (*sliders*) pour disques durs ; usine de ≈ 600 000 pieds² ; « over 10,000 » employés (gouvernement) ; ≈ 9 500 vers 2013 (Inquirer) | LTI, Biñan ; adresse OSM : Malamig ; carte : Biñan | 14.2720, 121.0546 | EMIS : https://www.emis.cn/php/company-profile/PH/Western_Digital_Storage_Technologies_Philippines_Corporation_en_13778914.html ; Inquirer : https://business.inquirer.net/?p=181089 ; RTVM : https://rtvm.gov.ph/?p=38243 | 🟡 |
| **Nidec Philippines Corp.** | Moteurs de broche pour disques durs ; ≈ 8 000 employés (PEZA, 2020, lors d'un foyer de COVID-19) | 136 North Science Ave., SEPZ LTI, Biñan ; adresse OSM : Loma ; carte : Biñan | 14.2770, 121.0616 | Inquirer 2020 : https://newsinfo.inquirer.net/1314648/cluster-outbreaks-hit-laguna-factories | 🟡 (vu en extrait) |
| **Toshiba Information Equipment (Philippines), Inc. (TIP)** | Disques durs, SSD, cartes électroniques ; fondée le 7 juil. 1995 ; premières livraisons de la nouvelle ligne le 2 oct. 2026 ; capacité presque doublée d'ici l'exercice 2027 | 103 East Main Ave. Ext., LTI, Biñan (+ annexe LTI) ; adresse OSM : Malamig ; carte : Biñan | 14.2767, 121.0583 | TIP ✅ ; TechRepublic, 9 oct. 2026 : https://www.techrepublic.com/article/news-toshiba-hdd-production-expansion-apac-philippines/ 🟡 | ✅/🟡 |
| **Amkor Technology Philippines** (P3, P4) | Assemblage et test de semi-conducteurs ; usine P4 achevée en 1998 | 119 North Science Ave., LTI, Biñan ; carte : Biñan | 14.2730, 121.0521 | Offre DOLE : https://philjobnet.gov.ph/job-vacancies/job/executive-director-1222343 ; https://amkor.com/company-history | 🟡 |
| **Cirtek Electronics Corp.** (Cirtek Holdings, coté en bourse) | Assemblage de semi-conducteurs ; projet de motos électriques approuvé par le PEZA en 2024 | 116 East Main Ave., Phase V-SEZ, LTI, Biñan ; carte : Loma | 14.2788, 121.0674 | https://datasheets.com/manufacturer/cirtek-electronics-corporation-37954 ; MB 2024 : https://mb.com.ph/2024/01/05/cirtek-to-venture-into-production-of-electric-motorcycles | 🟡 |
| **TDK Electronics Philippines** | Composants électroniques | LTI ; adresse OSM : Loma ; carte : Biñan | 14.2735, 121.0626 | PhilJobNet : https://philjobnet.gov.ph/job-vacancies/company/tdk-phils-corp-13015 ; OSM | 🟡 |
| **SunPower Philippines / Maxeon** | Cellules solaires ; 1ʳᵉ usine aux Philippines (2003 ou 2004) à la Technopark de Biñan ; « Fab 2 » fermée en 2016 (lieu exact non vérifié) ; actifs non américains repris par TCL | East Main Ave., LTI, Biñan ; carte : Loma | 14.2769, 121.0665 | Cushman & Wakefield : https://www.cushmanwakefield.com/en/philippines/properties/for-lease/office/laguna/binan/e-main-ave-laguna-technopark/maxeon-solar-technologies-l ; PV Tech 2016 : https://www.pv-tech.org/sunpower-closing-700mw-fab-2-with-2500-job-losses/ | 🟡 |
| **Katolec Philippines Corp.** | Électronique (EMS) | LTI ; adresse OSM : Timbao ; carte : Biñan | 14.2787, 121.0583 | Mairie 2018 : https://www.binan.gov.ph/katolec-fulfills-csr-thru-binan-lgu/ ; OSM | ✅ (présence) / 🟡 (lieu) |
| **Ryonan Electric Philippines** | Composants électroniques | LTI ; adresse OSM : Timbao ; carte : Biñan | 14.2788, 121.0592 | OSM ; WP Biñan | 🟡 |
| **Shin-Etsu Philippines** | Matériaux (plusieurs bâtiments, « Plant 7 ») ; aussi client d'ALLHC | East Main Ave., LTI ; adresse OSM : Loma ; carte : Loma | 14.2815, 121.0707 | OSM ; ALLHC (liste de locataires) : https://www.ayalalandlogistics.com/all-news/newbie-station-logistic-inc-signs-new-lease-with-allhc-for-warehouse-expansion/ | 🟡 |
| **BAG Electronics Inc.** | Électronique ; usine de 5 244 m² | Blk 5 Lot 7, LIIP | (LIIP) | Kalibrr : https://www.kalibrr.com/id-ID/c/bag-electronics-inc/jobs | 🟡 |

**B. Automobile et pièces**

| Entreprise | Activité | Lieu | Coord. | Source | Statut |
|---|---|---|---|---|---|
| **Isuzu Philippines Corp.** | Assemblage de camions et véhicules utilitaires ; toit solaire de 1 100 panneaux | 114 Technology Ave., Phase II, LTI, Biñan ; adresse OSM : Malamig ; carte : Biñan | 14.2731, 121.0606 | CAMPI : https://campiauto.org/our-members/isuzu-philippines-corporation/ ; https://www.autoindustriya.com/auto-industry-news/isuzu-ph-powers-up-laguna-plant-with-1-100-solar-panels.html | ✅/🟡 |
| **Isuzu Autoparts Manufacturing Corp.** | Boîtes de vitesses | 114 North Main Ave., Phase II, LTI ; carte : Biñan | 14.2695, 121.0591 | EMIS : https://www.emis.cn/php/company-profile/PH/Isuzu_Autoparts_Manufacturing_Corporation_en_13778916.html | 🟡 |
| **Honda Parts Manufacturing Corp. (HPMC)** | Pièces auto et moto | 101 North Science Ave., Phase II, LTI, Biñan ; carte : Biñan | 14.2640, 121.0553 | EMIS : https://www.emis.cn/php/company-profile/PH/Honda_Parts_Manufacturing_Corporation_en_2624387.html | 🟡 |
| **Furukawa Electric Autoparts Philippines** | Faisceaux électriques | 113 East Main Ave., Phase IV, LTI, **Loma** ; carte : Loma | 14.2795, 121.0651 | EMIS : https://www.emis.cn/php/company-profile/PH/Furukawa_Electric_Autoparts_Philippines_Inc_en_13850278.html | 🟡 |
| **Joyson Safety Systems (Philippines)** (ex-Takata) | Ceintures et airbags | 106 East Main Ave., LTI, adresse « Malamig » ; carte : Loma | 14.2780, 121.0632 | PhilJobNet : https://philjobnet.gov.ph/job-vacancies/company/joyson-safety-systems-philippines-corporation-424109 | 🟡 |
| **Daiwa Seiko Phils. Corp.** | Pièces de transmission automobile ; ≥ 200 employés | Lots 8-9, Block 5, Star Ave., **LIIP, Mamplasan** | 14.2880, 121.0768 | PhilJobNet : https://philjobnet.gov.ph/job-vacancies/company/daiwa-seiko-phils-corporation-83626 | 🟡 |
| **Sazgar Philippines Inc.** | Véhicules (groupe pakistanais ; activité exacte non vérifiée) | Mamplasan | — | Blog : https://philippinesfactory.blogspot.com/2026/10/sazgar-philippines-inc.html | 🔴 |
| *(pour comparaison, côté Santa Rosa)* Honda Cars Philippines ; Mitsubishi Motors Philippines | Assemblage automobile | Don Jose, **Santa Rosa** | 14.2587, 121.0621 ; 14.2631, 121.0624 | OSM | 🟡 — **ne pas attribuer à Biñan** |

**C. Médical, pharmacie, optique**

| Entreprise | Activité | Lieu | Coord. | Source | Statut |
|---|---|---|---|---|---|
| **Terumo Philippines** | Seringues, aiguilles, cathéters IV ; en 2010, plan de 45 M$ ; usine de ₱5,3 milliards, 1,7 milliard de seringues par an | 124 East Main Ave., LTI ; adresse OSM : Loma ; carte : Loma | 14.2844, 121.0728 | Philstar 2010 : https://www.philstar.com/business/2010/07/08/590838/terumo-invest-45-million-five-years-expand-syringe-output ; DHL 2026 : https://www.dhl.com/ph-en/home/press/press-archive/2026/dhl-supply-chain-philippines-expands-dhl-health-logistics-capabilities-with-terumo-partnership.html | 🟡 |
| **Unilab Pharma Campus** : Amherst Laboratories (usine de ₱1,7 milliard inaugurée en mars 2007 sur 35 000 m²), Sekaya (R&D, 2017), First Pioneer Distribution | Médicaments | « Mamplasan » selon la presse ; carte et OSM : Bungahan/Zapote | 14.3088, 121.0794 | Philstar 2007 : https://philstar.com/business/2007/03/15/389572/gma-inaugurates-unilabs-p17-b-amherst-plant-laguna ; Inquirer 2025 (ci-dessus) | 🟡 |
| **Transitions Optical Philippines** | Verres photochromiques | Star Ave., **LIIP** ; carte : Mamplasan | 14.2862, 121.0738 | OSM ; WP Biñan | 🟡 |
| **Optodev (groupe Essilor)** | Verres optiques | LIIP ; adresse OSM : Ganado ; carte : Mamplasan | 14.2850, 121.0743 | OSM ; WP Biñan | 🟡 |

**D. Alimentation et boissons**

| Entreprise | Activité | Lieu | Coord. | Source | Statut |
|---|---|---|---|---|---|
| **Gardenia Bakeries Philippines** (QAF, Singapour) | Pain industriel ; à Biñan depuis 1997–1998 ; 6ᵉ usine (₱1 milliard, 150 000 pains/jour) ouverte en janv. 2016 ; visites d'usine | **LIIP, Mamplasan** | 14.2898, 121.0799 | Philstar 2016 : https://www.philstar.com/business/2016/01/21/1544938/gardenia-opens-new-p1-b-bread-plant-laguna/amp/ ; https://www.gardenia.com.ph/?p=3456 | ✅/🟡 |
| **The Laguna Creamery Inc.** (glaces Carmen's Best) | Laiterie | Brentville, Mamplasan | — | Blog : https://philippinesfactory.blogspot.com/2025/07/the-laguna-creamery-inc.html | 🔴 |
| Producteurs de **puto**, *pinipig*, *sorbetes* | Petites entreprises familiales (San Vicente, Canlalay, Santo Domingo) | centre-ville | voir § 7 | EB | 🟡 (voir la note culture) |
| *(côté Santa Rosa)* Coca-Cola ; Monde Nissin | Boissons ; biscuits et nouilles | Santa Rosa | 14.2828, 121.0869 ; 14.2930, 121.1001 | OSM | 🟡 — **pas Biñan** |

**E. Machines, métal, autres industries**

| Entreprise | Activité | Lieu | Coord. | Source | Statut |
|---|---|---|---|---|---|
| **Atlas Copco Philippines** | Compresseurs, outils | LTI ; carte : Biñan | 14.2699, 121.0568 | OSM ; WP Biñan | 🟡 |
| **Kito Philippines** | Palans et appareils de levage | North Science Ave., LTI ; carte : Biñan | 14.2756, 121.0587 | OSM ; WP Biñan | 🟡 |
| **Maccaferri Philippines Manufacturing** (groupe italien) | Gabions, grillages | Blk 3 Lots 3-4, Star Ave., LIIP | (LIIP) | Kalibrr : https://www.kalibrr.id/c/maccaferri-philippines-manufacturing-inc/jobs | 🟡 |
| **Nuvali Steel Processing Center (NSPC)** | Traitement de l'acier ; visite du président Marcos le 26 sept. 2026 | Biñan (lieu exact non donné) | — | Tribune : https://tribune.net.ph/2026/09/26/marcos-inspects-steel-processing-center-highlights-industry-40 | 🟡 |
| **MCC Labels** | Étiquettes ; panneaux solaires inaugurés avec le BPLO | Biñan | — | Titre d'un article de la mairie (non lu) : https://www.binan.gov.ph/bplo-binan-head-atty-edward-vange-arriba-and-senior-vp-mr-lars-scheutzenmeister-spearheaded-the-inauguration-of-mcc-labels-solar-panels-for-sustainable-energy/ | 🟡 |
| **Kaserbisyo Garments Factory** (projet de la ville) | Uniformes ; 120 machines à coudre ; client Gingersnaps | Malamig, ouvert le 21 juin 2017 | — | https://www.binan.gov.ph/kaserbisyo-garments-factory-opens-expect-binaneses/ | ✅ |

**F. Logistique et entrepôts**

| Entreprise | Activité | Lieu | Source | Statut |
|---|---|---|---|---|
| **ALogis Biñan** (AyalaLand Logistics) | Entrepôts prêts à louer (103 000 m² selon le rapport 2024) ; Newbie Station loue 13 000 m² (2021) puis 11 600 m² (2023) à « ALogis Biñan 4 » ; Lemcon s'y installe | LTI | https://www.ayalalandlogistics.com/all-news/newbie-station-logistic-inc-signs-new-lease-with-allhc-for-warehouse-expansion/ ; https://www.ayalalandlogistics.com/all-news/lemcon-philippines-moves-to-new-alogis-binan-warehouse-facility/ | ✅/🟡 |
| **ALogis Artico Biñan 1 et 2** | Entrepôts frigorifiques (rachat de Technofreeze en 2021, terrain de 11 800 m²) | Biñan | MB 2021 : https://mb.com.ph/2021/04/14/allhc-acquires-laguna-cold-storage-facility/ ; `foncier.md` | 🟡 |
| **NX Logistics Philippines** (Nippon Express) | Logistique | LTI | https://philjobnet.gov.ph/job-vacancies/company/nx-logistics-philippines-inc-103497 | 🟡 |

**G. Services informatiques (BPO) et innovation**

- 🟡 **Southwoods Office Towers 1 et 2** (Megaworld) : seules tours de bureaux « BPO » documentées à Biñan ; aucun locataire nommé trouvé.
- ✅ **One Asia Business Center** (IT Park, 2009), San Antonio, près de Pavilion Mall (14.3246, 121.0870).
- ✅ **DLSU Innovation Hub** (KIST, 2025), campus DLSU Laguna (14.2609, 121.0428).
- 🟡 Le chapitre actuel parle de « business process outsourcing offices » : c'est vrai pour Southwoods, mais sans chiffre d'emplois. (Voir aussi `apres-guerre.md`, tableau des corrections.)

**H. Industries locales anciennes** (🟡 EB) : chaussures et *bakya* (Casile, San Jose, De La Paz, Malaban), casquettes et chapeaux (Platero, où l'on fait aussi des jeepneys et du métal), *padyak* ou cyclo-pousses (Soro-soro), camions et calèches (Canlalay). La ville dit que l'industrie de la chaussure de De La Paz « has disappeared over the years ».

---

## 3. Commerce, services et finances de la ville

### 3.1 Malls et marché

| Lieu | Ouverture | Barangay | Coord. | Source | Statut |
|---|---|---|---|---|---|
| **Marché public de Biñan** (marché couvert + marché humide) | ancien ; aire de restauration à l'étage en 2021 | Poblacion (rue J. Gonzales) | 14.3380, 121.0847 | Mairie Hist. ; EB (« 24/7 ») ; WP (pic de 3 h à 7 h) ; https://www.binan.gov.ph/mayor-arman-vm-gel-lead-project-tour-2021/ | 🟡 (« le plus grand de CALABARZON » = auto-description) |
| **Pavilion Mall** (Ayala Malls) | **oct. 1999** | San Antonio / Santo Niño (route nationale) | 14.3288, 121.0896 | Wikipédia « Ayala Malls » : https://en.wikipedia.org/wiki/Ayala_Malls | 🟡 |
| **Biñan Central Mall** | pré-ouverture **9 sept. 2011** | San Vicente | 14.3323, 121.0821 | Mairie Hist. | 🟡 |
| **Southwoods Mall** (Megaworld) | **30 sept. 2017** ; ₱2 milliards ; 58 000 m² ; 4 cinémas | San Francisco (Southwoods City) | 14.3315, 121.0505 | Mairie 2017 | ✅ |
| SM Savemore | ? | San Vicente | 14.3321, 121.0822 | OSM | 🟡 |
| C. Morales Mall ; Umbria Commercial Center | ? | ? | — | WP Biñan | 🔴 (non vérifiés) |
| « Robinsons Townville Biñan » | projet | ? | — | Liste « planned malls » d'un miroir de Wikipédia : https://wiki.privacytools.io/wiki/Robinsons_Malls | 🔴 |

- 🟡 Les grandes enseignes nouvelles s'installent **à côté** : SM Center San Pedro (fin 2023), SM Nuvali à Santa Rosa (prévu fin 2026). Avant, les habitants de San Pedro « would have to go to Biñan » (Philstar 2023 : https://philstar.com/business/2023/12/29/2321982/sms-san-pedro-mall-spurs-local-economy).

### 3.2 Santé, écoles, universités

| Lieu | Barangay | Coord. | Source | Statut |
|---|---|---|---|---|
| **Biñan City Hospital** (Ospital ng Biñan), hôpital public ; nouveau bâtiment de 3 étages, « more than 180 beds » visé fin 2023 | Canlalay | 14.3371, 121.0743 | BM 2023 ; MB 2022 : https://mb.com.ph/2022/2/11/binan-lgu-seeks-to-boost-healthcare-services-with-bigger-public-hospital | 🟡 |
| Biñan Doctors Hospital (privé) | Platero | 14.3260, 121.0921 | OSM | 🟡 |
| **University of Perpetual Help System Laguna** (UPHSL), fondée en 1976 (ex-Perpetual Help College of Laguna) | Santo Niño | 14.3298, 121.0848 | https://en.wikipedia.org/wiki/University_of_Perpetual_Help_System | 🟡 |
| **PUP Biñan** (Polytechnic University of the Philippines), vers 2010 | Zapote | 14.3136, 121.0782 | https://en.wikipedia.org/wiki/Polytechnic_University_of_the_Philippines_Bi%C3%B1an ; EB | 🟡 |
| **DLSU Laguna Campus** (Leandro V. Locsin Campus), 50 ha ; intégré à DLSU en 2012 | Malamig (procl. 985) | 14.2609, 121.0428 | Procl. 985 ✅ ; https://en.wikipedia.org/wiki/De_La_Salle_Canlubang 🟡 | ✅/🟡 |
| Colegio San Agustin–Biñan | San Francisco (Southwoods) | 14.3305, 121.0535 | OSM | 🟡 |

- 🔴 **Banques** : la ville parle de « numerous banking institutions » ; aucun chiffre trouvé (piste : liste des agences bancaires par ville du BSP). Le CMCI classe Biñan 15ᵉ pour « Financial Deepening » (2024).

### 3.3 Promoteurs immobiliers

(Surfaces et propriétaires : voir `foncier.md` § 2–3. Je ne donne ici que le rôle économique.)
- ✅/🟡 **Ayala** : Laguna Technopark et ALogis (ALLHC) ; Pavilion Mall (Ayala Malls) ; **Broadfield** (Alveo, lancé en oct. 2019 ; 120 ha selon Philstar 2019, 63 ha selon MB 2025 : https://philstar.com/business/2019/10/18/1961069/alveo-unveils-120-hectare-laguna-township ; https://mb.com.ph/2025/5/1/ayala-land-s-southern-estates-showcase-blueprint-for-regional-expansion).
- 🟡 **Megaworld / GERI** : **Southwoods City**, 561 ha à la limite de Carmona (Cavite) et de Biñan ; golf, mall, tours de bureaux (https://www.megaworldcorp.com/townships/southwoods-city).
- 🟡 **ArthaLand** : **Sevina Park**, ≈ 8,1 ha, lancé en juin 2019, présenté comme le premier quartier certifié « LEED Platinum » du pays, près de l'échangeur Laguna Boulevard de la CALAX (https://businessmirror.com.ph/2019/06/24/arthaland-launches-green-pioneering-township-project/). **Ce n'est pas un projet d'Ayala** (erreur fréquente).
- 🟡 Autres : Filinvest (Brentville), Federal Land (Meadowcrest), lotissements Jubilation (Zapote) et One Asia (`foncier.md`).

### 3.4 Finances, entreprises, classement

- 🟡 **Classe** : ville « component city » de **1ʳᵉ classe** (Wikipédia ; non lu dans un texte du DOF).
- ✅ **Revenus annuels réguliers** (BLGF, via PhilAtlas) : ₱606,9 M (2009) ; ₱764,3 M (2010) ; ₱965,1 M (2011) ; ₱1 024,2 M (2012) ; ₱1 095,6 M (2013) ; ₱1 297,2 M (2014) ; ₱1 349,5 M (2015) ; **₱1 530,3 M (2016)**.
- 🟡 **Revenus récents** (Wikipédia, d'après Wikidata) : ₱2 698 M (2020), ₱3 555 M (2022), **₱3 890 M (2024)**. Budget voté pour 2023 : **₱3,2 milliards** (BM 2023 ✅).
- ✅ **Entreprises** : **7 013 enregistrées en 2020, 7 524 en 2021** (+531) ; « more than 20,000 employment opportunities » proposés par le PESO en 2021 ; « our top 100 taxpayers contribute 95 percent of all the taxes here in Biñan City » (Dimaguila, OpinYon 2022a). Guichet unique de janvier 2025 : 296 nouvelles entreprises, 6 169 renouvellements, **₱588,2 M** encaissés (+4,36 %) (OpinYon 2025a).
- ✅ **CMCI (DTI), catégorie « component city »** :

| Année | Rang général | Score | Economic Dynamism | Remarque |
|---|---|---|---|---|
| 2014 | 136ᵉ | 15,88 | 119ᵉ | |
| 2015 | 35ᵉ | 38,24 | 7ᵉ | |
| 2016 | 32ᵉ | 36,89 | 3ᵉ | |
| 2017 | 7ᵉ | 48,29 | 5ᵉ | |
| 2018 | 6ᵉ | 50,98 | 5ᵉ | |
| 2019 | 13ᵉ | 47,72 | 17ᵉ | |
| 2020 | 11ᵉ | 50,68 | 12ᵉ | |
| 2021 | 9ᵉ | 44,20 | 7ᵉ | |
| 2022 | 6ᵉ | 45,62 | **1ᵉʳ** | Innovation 2ᵉ |
| 2023 | 9ᵉ | 46,14 | **1ᵉʳ** | Innovation 3ᵉ |
| 2024 | **7ᵉ** | **47,80** | **1ᵉʳ** (9,73) | **Innovation 1ᵉʳ** ; Government Efficiency 38ᵉ ; Infrastructure 40ᵉ ; Resiliency 35ᵉ |

  Détail 2024 (✅ CMCI) : Employment Generation 1ᵉʳ ; Productivity 1ᵉʳ ; Getting Business Permits 1ᵉʳ ; Local Economy Size 2ᵉ ; Cost of Doing Business 60ᵉ ; Capacity of School Services 84ᵉ. Le rang 7ᵉ de 2024 est confirmé par Philstar (https://philstar.com/nation/2024/08/26/2380604/quezon-city-named-most-competitive-city-fourth-straight-year) 🟡. Les pondérations ont changé (3 piliers en 2014–2016, 4 piliers à 25 % en 2017–2021, 5 piliers à 20 % depuis 2022) : comparer les rangs, pas les scores.

---

## 4. Agriculture

- ✅ **Ce qu'il reste : très peu.** Recensement agricole PSA 2012 : **118 exploitations, 135 ha** (1,1 ha en moyenne) à Biñan. À comparer : San Pedro 360 exploitations / 532 ha ; Santa Rosa 300 / 223 ha ; Calamba 2 333 / 1 721 ha ; province de Laguna 47 331 / 50 296 ha (PSA CAF 2012). 135 ha ≈ **3 %** des ≈ 4 350 ha de la ville.
- 🟡 **La conversion.** Étude de K. Engay-Gutierrez (2015, citée par OpinYon 2022b) : terres agricoles **86,9 % (1979) → 9,13 % (1999)** ; bâti **10,5 % → 70,0 %**. C'est l'effet de la Technopark (1990), du LIIP (1992), de la SLEX et des lotissements (`apres-guerre.md`, `foncier.md`).
- ✅/🟡 **Où sont les dernières terres agricoles ?** La ville classe comme ayant un usage « agricultural » : **Bungahan, Langkiwa, Malaban, Platero, Santo Tomas, Timbao** (EB). Le barrage de **Timbao** (époque espagnole) « is still supplying irrigation water to the farmlands » (EB 🟡) ; OSM le place à 14.2780, 121.0449. Des rizières irriguées sont documentées en justice à **Langkiwa** (Alberto, 1992–2001 ; Puyat, 1998–2012) et à **Zapote** (années 1960) : voir `foncier.md` § 2.
- ✅ **Le PLU (CLUP) n'existe pas en ligne.** En 2018, l'agenda de la mairie dit encore « Finish the CLUP » (Mairie ELA). Dans l'affaire Puyat, les cultivateurs disaient que Biñan n'avait « no approved plan/zoning ordinance » (HLURB) au moment du classement industriel de 1997 (`foncier.md`). Le CMCI 2024 classe pourtant Biñan 7ᵉ pour « Land Use Plan » : un plan existe donc au moins sous forme de document soumis. **Carte d'occupation du sol officielle : non trouvée.**
- ✅ **Agriculture urbaine.** « Organik Farm » à De La Paz (formation de 25 personnes à l'agriculture biologique, déc. 2017 ; plantation d'arbres avec Katolec en 2018) (https://www.binan.gov.ph/25-binanenses-graduate-organic-farming/).
- 🟡 Les noms gardent la mémoire du riz : **pinipig** (riz gluant aplati) de Canlalay ; Bungahan « used to be a rice field » ; San Vicente « used to be a rice field owned mostly by the Samaniegos and Reyeses » (EB).
- ✅ **Contexte provincial** (PSA CAF 2022) : Laguna, **78 345 exploitations sur 41 499 ha** (0,53 ha en moyenne). Les tableaux 2022 par commune ne sont pas publiés.

---

## 5. Élevage : porcs et volailles

- ✅ **L'ordonnance de 2004.** Le 24 nov. 2004, le conseil municipal adopte l'**ordonnance n° 06 (2004)**, « Urban Control Zones Regulation and Gradual Phase Out of Large Livestock Farms in Biñan, Laguna ». Les grands élevages (« more than ten (10) swine heads or more than five hundred birds ») ont trois ans pour réduire leurs bêtes ; aucun nouveau grand élevage n'est permis. La province l'approuve le 6 avril 2005 (SC 2022).
- ✅ **Le procès.** Holiday Hills Stock & Breeding Farm Corp. et Domino Farms, Inc. attaquent l'ordonnance (7 févr. 2006). Le tribunal de Biñan la valide (2008), la Cour d'appel l'annule (2011 ; une porcherie serait une nuisance *per accidens*). La **Cour suprême** casse l'arrêt le **10 oct. 2022** : « Municipal Ordinance No. 06 (2004) of the Municipality of Biñan, Laguna is declared NOT UNCONSTITUTIONAL ». Les habitants avaient signé une plainte collective contre ces élevages « stench-emitting », situés « near residential subdivisions » (SC 2022). Une première affaire entre les mêmes parties existe (G.R. No. 172700 ; https://jur.ph/jurisprudence/summary/municipality-biaan-v-holiday-hills-farm) 🟡.
- ✅ **La fermeture (2025).** Selon OpinYon (13 janv. 2025), les deux élevages sont à **Langkiwa et Santo Tomas** ; la ville avait ordonné leur arrêt par l'**Executive Order n° 19 s. 2018**. Le maire Dimaguila : « Wala na pong permit ang Holiday Hills at Domino Farms kaya sila po ay titigil na sa pag-aalaga ng mga baboy. » Le capitaine de Langkiwa parle d'une odeur « amoy baboy ». EB cite aussi, à Langkiwa, « the Holiday Hills Stock and Breeding Farm, which is engaged in the piggery industry ».
- 🔴 **Porcs près du lac.** Je n'ai trouvé **aucune source** sur des élevages à Malaban, De La Paz ou San Antonio, ni de mesure de la LLDA contre une porcherie de Biñan. La LLDA applique depuis 2016 le DAO 2016-08 et exige un permis de rejet même pour les petits éleveurs (LBTimes 🟡 : https://lbtimes.ph/whole-of-community-approach-stricter-policy-implementation-sought-water-quality-of-laguna-lakes-tributaries-continue-to-deteriorate-%e2%88%92-llda-data/). Les porcheries de basse-cour sont une source connue de pollution des rivières du lac (GMA 🟡 : https://www.gmanetwork.com/news/scitech/content/233349/why-laguna-lake-water-isn-t-safe-for-people-or-fish/story). **Formulation conseillée** : « Small backyard piggeries can still be seen near the lakeshore » seulement si l'auteur l'a vu lui-même et le présente comme une observation ; sinon ne rien affirmer.
- ✅ **Porcs de la Laguna** (PSA, au 1ᵉʳ janvier ; total / dont élevages commerciaux) :

| Année | Total | Commercial | Remarque |
|---|---|---|---|
| 2010 | 279 204 | 202 454 | |
| 2015 | 180 556 | 130 923 | |
| 2019 | 146 081 | 105 925 | |
| 2020 | 194 394 | 140 957 | avant l'ASF |
| **2021** | **60 258** | 43 694 | **−69 %** (ASF) |
| 2022 | 64 565 | 49 377 | |
| 2023 | 92 965 | 69 138 | |
| 2024 | 103 207 | 72 845 | |
| 2026 | 130 206 | 109 307 | la répartition 2025 est incohérente : ne pas l'utiliser |

  Recensement 2012 : 6 757 élevages porcins et 155 258 porcs dans la province (PSA CAF 2012 b). **Aucun chiffre pour Biñan seule.**
- 🟡 **ASF dans la Laguna.** Province indemne jusqu'en mars 2020 ; premier foyer à **Calamba** ; foyers de juillet 2020 surtout dans des élevages de basse-cour de Calamba, Los Baños, Victoria et Pakil, ≈ 1 000 porcs abattus (Inquirer 2020 : https://newsinfo.inquirer.net/1304745/swine-fevers-return-in-laguna-farms-jacks-up-pork-prices). Août 2024 : San Pablo, Calamba, Nagcarlan (GMA ✅ : https://www.gmanetwork.com/regionaltv/news/103476/asf-spreads-to-two-cities-and-one-town-in-laguna/story/). **Biñan n'apparaît dans aucune source lue.** La chute de 2020–2021 vient aussi des interdictions de transport et de l'abattage préventif dans toute la région.
- 🔴 Volailles : aucune donnée locale.

---

## 6. Pêche et aquaculture

- ✅ **Deux barangays au bord du lac.** Sur la carte du dépôt, la rive de Biñan fait **≈ 4,1 km**, de 14.3605, 121.0774 (limite de San Pedro) à 14.3370, 121.0991 (limite de Santa Rosa), le long de **De La Paz** (≈ 1 km) et surtout de **Malaban** (≈ 3 km). EB : à Malaban, « the people's main sources of livelihood … are fishing and footwear production » ; De La Paz est « known for shoemaking, kalesa making, and fishing ».
- ✅ **Enclos et cages en 2006** (LLDA via PIDS 2008, tableau 1) :

| Commune (zone ZOMAP) | Enclos : opérateurs | Enclos : ha | Cages : opérateurs | Cages : ha | Total ha |
|---|---|---|---|---|---|
| Muntinlupa (A) | 107 | 2 179 | 218 | 168 | 2 347 |
| San Pedro (A) | 26 | 778 | 65 | 58 | 836 |
| **Biñan (B)** | **26** | **650** | **76** | **35** | **686** |
| Santa Rosa (B) | 2 | 100 | 8 | 3 | 103 |
| Calamba (B) | 8 | 150 | 43 | 25 | 174 |
| Tout le lac | 455 | 12 117 | 1 599 | 998 | 13 115 |

  Biñan a **le plus grand nombre d'opérateurs** de la province de Laguna (102, contre 91 à San Pedro) et la **2ᵉ surface** après San Pedro (686 ha contre 836 ha), soit ≈ 5 % du total du lac. Avec Binangonan et Muntinlupa, elle a été choisie pour l'enquête de 2007 « because they had the largest number of registered fishpen and fishcage operations ».
- ✅ **Les poissons** (enquête 2007, 20 enclos et 40 cages de Biñan) : **16 enclos sur 20 élèvent seulement du bangus** (milkfish), 2 de la carpe ; **39 cages sur 40 élèvent seulement de la carpe** (en élevage extensif). Les alevins de bangus viennent surtout de Bulacan ; le tilapia vient des écloseries du bord du lac (PIDS 2008). La carpe à grosse tête (*bighead carp*) est l'espèce des cages (voir l'image Commons du port de Cardona).
- 🟡 **Pêche libre** : tilapia, *kanduli* (poisson-chat), *ayungin*, *dalag*, crevettes sont les espèces classiques du lac ; **aucune donnée de capture propre à Biñan** trouvée. Sur tout le lac, la pêche municipale passe de 64 116 t (2018) à 32 535 t (2022) (Bulatlat 2025 : https://www.bulatlat.com/2025/07/18/fisherfolk-face-displacement-from-laguna-lake/) 🟡.
- ✅ **Les règles aujourd'hui.** La résolution LLDA n° 540 (2018) révise le zonage de 1999 et plafonne enclos et cages à **9 200 ha** pour tout le lac. 60 % de cette surface va aux particuliers (cage de 200 m² à 1 ha ; enclos de 1 à 5 ha), 40 % aux sociétés et coopératives (enclos par blocs de 10 ou 20 ha). La priorité va aux pêcheurs enregistrés de la région. La résolution n° 561 (2019) donne un numéro permanent à chaque structure (LLDA BR 540, 561).
- 🔴 **Nombre de pêcheurs de Biñan** : non trouvé (ni FishR du BFAR, ni chiffre de la ville). Repères : province de Laguna, **5 609** opérateurs de pêche « household », dont 5 463 en eau douce (PSA CAF 2022 ✅) ; tout le lac, « around 13,000 fisherfolk » (LLDA, citée par Rappler 2024 🟡).
- ✅ **Ce que fait la ville** : 25 gardes-pêche députés par le BFAR et la LLDA prêtent serment le 22 nov. 2018 ; marquage SSEN des bateaux enregistrés avec la **Philippine Coast Guard Sub-Station Biñan** (25 juin 2026) ; nettoyage de la rive à **Malaban** avec le CFARMC et Bantay Lawa (Journée mondiale de la pêche, 21 nov. 2025) ; **Bangkarera sa Lawa** à Malaban (CENRO, 31 juil. 2024 et 2025 : course de bateaux de pêcheurs pour la gestion des eaux usées ; en 2025, les prix aident « mga mangingisdang labis na naapektuhan ng mga pagbaha ») ; formation aux produits de surimi à Santo Tomas (2024) (Mairie pêche).
- 🟡 **Pollution et jacinthe d'eau.** La rivière de Biñan (station LLDA n° 7) échoue régulièrement au critère d'oxygène dissous (LBTimes). La baie sud (« San Pedro, Biñan, Sta. Rosa, Cabuyao and Calamba ») favorise la prolifération de la jacinthe d'eau (PNA : https://alpha.pna.gov.ph/articles/1074643). Une mortalité de poissons a touché les cages de Biñan et Pila (tilapia, bangus ; date à préciser) (https://upd.edu.ph/category/research-highlights/page/8/). Le typhon Rolly (nov. 2020) a endommagé des structures de pêche à Biñan (https://newsinfo.inquirer.net/1356160/fish-pen-owners-in-laguna-de-bay-lose-p115m-to-rolly/amp).
- ✅/🟡 **La route-digue du lac (LLRN).** Phase 1 : **12,23 km** entre San Pedro et Cabuyao (7,50 km de digue, 4,73 km de viaducs), avec « one interchange at San Pedro/Binan ». « There will be economic impacts to commercial businesses, fish pens and fish cages » à San Pedro, Biñan, Santa Rosa et Cabuyao (AIIB LLRN ✅). Plan de réinstallation provisoire de l'ADB : **1 084 familles** touchées (681 déplacées physiquement, 182 économiquement, 273 avec des structures de pêche ou comme locataires, employés ou cultivateurs) ; le DPWH (2020) cite **Malaban, De La Paz, Casile et San Antonio** comme indirectement touchés ; une habitante de Malaban (installée en 1986) : « The fisherfolk would have a harder time » ; les inondations de Malaban seraient passées d'une tous les deux ans à « about every six months » (Rappler 2024 🟡). Pour le reste du projet, voir la note transports.

---

## 7. Usage du sol par barangay (pour la carte)

Population 2020 (PSA via PhilAtlas ✅). Usage du sol : fiches de la ville (EB 🟡). Catégorie : ma synthèse. **Ne pas imprimer les surfaces** (la carte du dépôt et la ville ne concordent pas, voir `foncier.md`).

| Barangay | Pop. 2020 | Usage du sol (EB) | Catégorie proposée pour la carte | Repère économique |
|---|---|---|---|---|
| Biñan | 7 103 | résidentiel, commercial, mixte, industriel | **Industriel** (Technopark) | IMI, Toshiba, WD, Isuzu, Amkor ; Sevina Park, Broadfield, DLSU |
| Bungahan | 3 812 | résidentiel, industriel, **agricole**, institutionnel | Mixte, reste agricole | Unilab Pharma Campus (limite Zapote) |
| Canlalay | 20 354 | résidentiel, commercial, industriel, institutionnel | Résidentiel/commercial | Hôpital de la ville ; *pinipig* |
| Casile | 4 240 | résidentiel, commercial, institutionnel | Centre (petit) | *bakya* |
| De La Paz | 32 841 | résidentiel, commercial, mixte | **Rive du lac, pêche** | Organik Farm |
| Ganado | 6 463 | résidentiel, industriel, commercial | **Industriel** (LIIP) | Optodev, Laguna BelAir |
| Langkiwa | 39 526 | résidentiel, commercial, industriel, **agricole** | Résidentiel, rizières résiduelles | **Holiday Hills** (porcs, fermé 2025) ; NHA Southville |
| Loma | 13 068 | industriel, résidentiel, mixte | **Industriel** (Technopark III) | Cirtek, Furukawa, Joyson, Terumo, Shin-Etsu, SunPower |
| Malaban | 35 989 | résidentiel, **agricole** | **Rive du lac, pêche** | Bangkarera, nettoyages, enclos au large |
| Malamig | 6 243 | industriel, résidentiel | **Industriel** + nature | Technopark (partie) ; DLSU Innovation Hub ; chutes de Tibagan ; Kaserbisyo |
| Mampalasan (Mamplasan) | 13 586 | industriel, résidentiel, commercial… | **Industriel** (LIIP + Technopark) | Gardenia, Daiwa Seiko, Transitions |
| Platero | 14 803 | industriel, résidentiel, commercial, **agricole**, institutionnel | Mixte | casquettes, jeepneys ; Biñan Doctors Hospital |
| Poblacion | 5 012 | résidentiel, commercial, mixte, patrimoine | **Centre commerçant** | **Marché public** |
| San Antonio | 36 163 | résidentiel, commercial, mixte | Résidentiel/commercial | Pavilion Mall ; One Asia IT Park |
| San Francisco | 39 271 | résidentiel, mixte, commercial | Résidentiel, grands projets | **Southwoods** (mall, tours) ; Splash Island (limite Soro-soro) |
| San Jose | 6 599 | résidentiel, commercial, mixte | Centre | chaussures, *bakya* |
| San Vicente | 9 236 | résidentiel, parcs, commercial, mixte | Centre commerçant | Biñan Central Mall ; puto |
| Santo Domingo | 6 159 | résidentiel, commercial, industriel | Centre | *sorbetes*, bougies |
| Santo Niño | 7 736 | résidentiel, commercial, institutionnel, mixte | Commercial/écoles | UPHSL |
| Santo Tomas | 55 649 | résidentiel, mixte, commercial, industriel, **agricole** | Résidentiel | **Domino Farms** (porcs, fermé 2025) |
| Soro-soro | 7 189 | résidentiel, commercial, industriel, mixte | Résidentiel | *padyak* |
| Timbao | 15 464 | résidentiel, industriel, **agricole**, mixte | Mixte | barrage d'irrigation ; Technopark III |
| Tubigan | 6 247 | résidentiel, commercial, industriel, mixte | Résidentiel | |
| Zapote | 14 684 | résidentiel, institutionnel, industriel, commercial | Résidentiel/administratif | hôtel de ville ; PUP ; Jubilation |

**Synthèse** : industriel = Biñan, Loma, Malamig, Mamplasan, Ganado (+ Timbao en partie) ; centre commerçant = Poblacion, San Vicente, Santo Domingo, San Jose, Casile ; rive du lac et pêche = Malaban, De La Paz ; restes agricoles = Bungahan, Langkiwa, Platero, Santo Tomas, Timbao (et Malaban selon EB) ; grands élevages documentés = Langkiwa, Santo Tomas (fermés en 2025).

---

## 8. Données pour les cartes

Coordonnées en degrés décimaux (WGS84). Sauf mention, **géocodage Nominatim/OpenStreetMap (± 100 m)**, vérifié dans les polygones du dépôt.

**Parcs et zones**

| Objet | Point (lat, lon) | Emprise approximative | Statut |
|---|---|---|---|
| Laguna Technopark, partie Biñan | nœud OSM 14.2761, 121.0521 | ≈ 14.257–14.290 N, 121.042–121.075 E (enveloppe des usines géocodées ; limite sud = limite de Santa Rosa sur la carte du dépôt) | 🟡 approximatif |
| LIIP | 14.2880, 121.0796 | ≈ 14.282–14.292 N, 121.072–121.082 E (Mamplasan/Ganado) | 🟡 approximatif |
| Southwoods City (Mall + tours) | 14.3315, 121.0505 | la ville nouvelle déborde sur Carmona (ouest) | 🟡 |
| One Asia Business Center (IT Park) | 14.3246, 121.0870 | 4,8 ha | 🟡 |
| Unilab Pharma Campus | 14.3088, 121.0794 | 3,5 ha | 🟡 |
| DLSU Laguna / Innovation Hub | 14.2609, 121.0428 | campus 50 ha ; hub 5 ha | 🟡 |
| Sevina Park | 14.2582, 121.0450 | ≈ 8 ha | 🟡 |
| Broadfield | 14.2618, 121.0490 | 63 à 120 ha | 🟡 |

**Commerce et services** : marché public 14.3380, 121.0847 ; Pavilion Mall 14.3288, 121.0896 ; Biñan Central Mall 14.3323, 121.0821 ; Southwoods Mall 14.3315, 121.0505 ; hôtel de ville 14.3150, 121.0798 ; Biñan City Hospital 14.3371, 121.0743 ; UPHSL 14.3298, 121.0848 ; PUP Biñan 14.3136, 121.0782.

**Usines (pastilles par secteur)** : voir les coordonnées du § 2.3.

**Lac, pêche, élevage, agriculture**

| Objet | Coordonnées | Source | Statut |
|---|---|---|---|
| Rive du lac (polyligne) | 14.3605, 121.0774 → 14.3575, 121.0882 → 14.3539, 121.0904 → 14.3500, 121.0925 → 14.3433, 121.0927 → 14.3407, 121.0958 → 14.3370, 121.0991 | limite de la ville moins les limites des communes voisines, carte du dépôt (simplifiée) | ✅ (calcul) / 🟡 (précision) |
| Quartier de pêcheurs de Malaban | 14.3464, 121.0889 (centre OSM) | OSM ; EB | 🟡 |
| Quartier de pêcheurs de De La Paz | 14.3490, 121.0849 (centre OSM) | OSM ; EB | 🟡 |
| Zone d'enclos et de cages au large de Biñan | **à dessiner** dans une bande à l'est de la rive (environ 121.10–121.13 E, 14.33–14.37 N) | aucune carte ZOMAP lue | 🔴 **à vérifier** sur une image Sentinel-2 ou Google Earth récente, ou sur la carte de zonage de la LLDA |
| Élevages de porcs Holiday Hills et Domino Farms (fermés 2025) | barangays Langkiwa (centre ≈ 14.2972, 121.0610) et Santo Tomas (≈ 14.3142, 121.0723) ; **emplacement exact inconnu** | OpinYon 2025b | 🟡 (barangay) / 🔴 (point) |
| Barrage d'irrigation de Timbao | 14.2780, 121.0449 | OSM | 🟡 |
| Restes agricoles | barangays Bungahan, Langkiwa, Platero, Santo Tomas, Timbao, Malaban (hachures par barangay, pas de parcelles) | EB ; PSA CAF 2012 (135 ha en tout) | 🟡 |

**Pour une vraie carte d'occupation du sol** (le CLUP n'étant pas en ligne) : extraire les polygones `landuse=industrial/residential/farmland` d'OpenStreetMap (fichier Geofabrik « philippines », l'API Overpass était bloquée ici), ou utiliser la couverture ESA WorldCover 10 m (2021) ; demander la carte du CLUP au City Planning and Development Office.

---

## 9. Plan proposé de la section (≈ 4 paragraphes, en anglais)

1. **A factory city on old cane fields.** Biñan's economy today rests on export manufacturing. About half of the 471-hectare Laguna Technopark lies inside the city (roughly 240 hectares proclaimed as economic zones between 1997 and 2015), in barangays Biñan, Loma, Malamig, Timbao and Mamplasan; the Laguna International Industrial Park fills Mamplasan and Ganado. The biggest employers make parts of the digital world: hard-disk sliders (Western Digital), disk-drive motors (Nidec), drives and SSDs (Toshiba), electronics assembly (IMI, Amkor, Cirtek). Car parts (Isuzu, Honda Parts, Furukawa, Joyson), medical devices (Terumo, Unilab) and bread (Gardenia) follow. Note for the reader: Honda Cars and Mitsubishi are on the Santa Rosa side of the Technopark.
2. **Market town, mall town.** The old role survives in Poblacion, where the public market — which the city calls the largest in CALABARZON — trades around the clock. Malls arrived with the subdivisions: Pavilion (1999), Biñan Central Mall (2011), Southwoods Mall (2017), with office towers for outsourcing firms. Some 7,500 businesses were registered in 2021, but the top 100 taxpayers pay 95 percent of local taxes. The national competitiveness index has ranked Biñan first among component cities for economic dynamism every year since 2022.
3. **The farms that disappeared — and the pigs that went to court.** In 1979 most of Biñan was farmland; by 1999 less than a tenth was, and the 2012 farm census counted only 118 farms on 135 hectares. Pockets of farmland survive in Langkiwa, Timbao and Bungahan, some still watered by a Spanish-era irrigation dam in Timbao. Large hog farms were phased out by a 2004 ordinance that the Supreme Court upheld in 2022; the last two lost their permits in 2025. African swine fever cut Laguna's pig herd by two-thirds in 2020.
4. **Life on the lakeshore.** Only two barangays, Malaban and De La Paz, touch Laguna de Bay, along about four kilometers of shore. In 2006 Biñan had more registered fish-pen and fish-cage operators than any other town in Laguna province — 102 of them, on some 686 hectares, mostly milkfish in the pens and carp in the cages. Fishers now face polluted rivers, water hyacinth, floods and the planned lakeshore road-dike, which will cross the fish-pen belt between San Pedro and Cabuyao.

---

## 10. Trois idées d'encadrés « Did you know? »

1. **Pigs before the Supreme Court.** In 2004 Biñan's town council voted to phase out every farm with more than ten pigs or five hundred birds. Two big hog farms sued. The case went all the way to the Supreme Court, which ruled on 10 October 2022 that the ordinance was "NOT UNCONSTITUTIONAL". In January 2025 the mayor announced that the two farms, in Langkiwa and Santo Tomas, no longer had permits. Sources : G.R. No. 200403 (2022) ; OpinYon, 13 Jan. 2025.
2. **Milkfish and carp off Malaban.** In 2006, 102 registered fish-pen and fish-cage operators worked 686 hectares of Laguna de Bay off Biñan — more operators than any other town in Laguna province (San Pedro had 91, on a larger area). Sixteen of twenty pens surveyed in 2007 raised only *bangus*; thirty-nine of forty cages raised only carp. Source : LLDA data in Israel, *Philippine Journal of Development* (2008).
3. **From 87 percent farmland to 3 percent.** In 1979, 86.9 percent of Biñan was farmland. Twenty years later it was 9 percent. The 2012 farm census found 118 farms on just 135 hectares — about 3 percent of the city's land. Sources : Engay-Gutierrez (2015) via OpinYon (2022) ; PSA, Census of Agriculture and Fisheries 2012.

(Variantes : **Half a Technopark** : ≈ 239 ha de proclamations PEZA dans Biñan, sur 471 ha ; **Number one for dynamism** : Biñan 1ʳᵉ « component city » du pays en Economic Dynamism (CMCI 2022–2024) ; **The hard drives of the world** : Western Digital, Nidec et Toshiba fabriquent à Biñan des pièces de disques durs.)

---

## 11. Illustrations libres de droits (licences vérifiées par l'API Commons le 16/11/2026 ; rien n'a été téléchargé dans le dépôt)

| # | Sujet | URL | Licence / auteur | Remarque |
|---|---|---|---|---|
| 1 | Southwoods Mall, juillet 2024 | https://commons.wikimedia.org/wiki/File:Southwoods_Mall,_Bi%C3%B1an,_Laguna,_Jul_2024_(1).jpg | CC BY-SA 4.0, Ralff Nestor Nacor ✅ | Série : intérieur, terminal de transport, « Market Manila » |
| 2 | Tours de bureaux et mall de Southwoods | https://commons.wikimedia.org/wiki/File:Southwoods_Office.jpg | CC0, B18IDH (16 avr. 2023) ✅ | Pour le BPO |
| 3 | Façade de Southwoods Mall (avant l'ouverture) | https://commons.wikimedia.org/wiki/File:Southwoods-Mall_facade.jpg | CC BY-SA 4.0, Kileyco17 (1ᵉʳ sept. 2017) ✅ | |
| 4 | Entrée de Pavilion Mall (2013) | https://commons.wikimedia.org/wiki/File:Bi%C3%B1an,Lagunajf9153_14.JPG | CC BY-SA 3.0, Ramon F. Velasquez (8 sept. 2013) ✅ | Sujet vérifié sur la vignette |
| 5 | Étals de fruits (durians, oranges), marché de Biñan (2013) | https://commons.wikimedia.org/wiki/File:Bi%C3%B1an,Lagunajf9153_20.JPG et https://commons.wikimedia.org/wiki/File:Bi%C3%B1an,Lagunajf9153_21.JPG | CC BY-SA 3.0, Ramon F. Velasquez ✅ | Sujet vu sur la vignette ; lieu exact (marché public ?) 🟡 |
| 6 | Rue commerçante près de la Plaza Rizal (2013) | https://commons.wikimedia.org/wiki/File:Bi%C3%B1an,Lagunajf9186_24.JPG | CC BY-SA 3.0, Ramon F. Velasquez ✅ | Ambiance du centre commerçant |
| 7 | Sortie « Laguna Technopark » de la CALAX | https://commons.wikimedia.org/wiki/File:CALAX_Laguna_Technopark_Nuvali_2024-06-09.jpg | CC BY-SA 4.0, LMP 2001 ✅ | Déjà proposée dans `apres-guerre.md` |
| 8 | Technopark et Greenfield, côté Santa Rosa | https://commons.wikimedia.org/wiki/File:686Barangays_Santa_Rosa_Laguna_Landmarks_06.jpg | CC0, Judgefloro (16 mai 2018) ✅ | Légende : « Santa Rosa side » |
| 9 | Enclos à poissons du lac la veille du typhon Basyang (Cardona, 2010) | https://commons.wikimedia.org/wiki/File:The_day_before_Basyang_(Cardona,_Rizal)_(4819809592).jpg | CC BY 2.0, Shubert Ciencia ✅ | Pas Biñan : à légender « fish pens on Laguna de Bay » |
| 10 | Carpe à grosse tête au port de Cardona (2013) | https://commons.wikimedia.org/wiki/File:Cardona,RizalFishPortjf5152_08.JPG | CC BY-SA 3.0, Ramon F. Velasquez ✅ | L'espèce des cages de Biñan |
| 11 | Laguna de Bay vue de l'ISS (22 janv. 2020) | https://commons.wikimedia.org/wiki/File:Laguna_de_Bay_2020_(cropped).jpg | Domaine public (NASA) ✅ | Cendres du Taal ; vérifier si les enclos sont visibles près de la rive ouest |
| 12 | Port de pêche de Cardona | https://commons.wikimedia.org/wiki/File:Cardona,Rizaljf5197_02.JPG | CC BY-SA 3.0, Ramon F. Velasquez ✅ | Pas Biñan |

**Remarques.** Aucune photo libre de la rive de Malaban, des usines de la Technopark côté Biñan, du LIIP, d'une rizière ou d'un élevage de Biñan n'a été trouvée sur Commons. Pistes : photos de l'auteur (Malaban, De La Paz, rizières de Langkiwa) ; demander aux entreprises (Gardenia fait visiter son usine) ; NASA Earth Observatory pour le lac. Les fichiers CC BY-SA exigent l'attribution et le partage dans les mêmes conditions.

---

## 12. Lacunes et prochaines étapes

1. **CLUP et zonage** : demander le CLUP et l'ordonnance de zonage au City Planning and Development Office (le CMCI montre qu'un plan a été soumis). C'est la seule source pour une vraie carte d'occupation du sol.
2. **Liste officielle des entreprises PEZA de Biñan** (nombre, emplois, exportations par zone) : site du PEZA (inaccessible ici), ou BPLO de la ville. Le chiffre d'emplois de chaque usine reste 🟡.
3. **Proclamation d'origine du LIIP-SEZ** et de « Laguna Technopark I » ; numéro de la proclamation de Sevina Park (2024).
4. **Pêcheurs de Biñan** : registre FishR du BFAR (Région IV-A) ou City Agriculture Office ; nombre actuel d'enclos et de cages au large de Biñan (LLDA) ; carte ZOMAG 2019.
5. **Porcs** : chiffres du City Veterinarian (élevages de basse-cour, cas d'ASF éventuels) ; vérifier sur place l'observation de l'auteur (porcheries près du lac : barangay, taille, dates).
6. **Riz** : surfaces récentes par barangay (City Agriculture Office ; RSBSA du DA) ; l'étude Engay-Gutierrez (2015) à lire directement.
7. **Malls** : date d'ouverture de Pavilion Mall (source Ayala), statut de « Robinsons Townville Biñan », C. Morales Mall.
8. **Banques et BPO** : nombre d'agences bancaires (BSP) ; locataires des tours de Southwoods.
9. **Revenus 2017–2025** : tableaux du BLGF (« Statement of Receipts and Expenditures ») au lieu de Wikipédia.
10. **Sites bloqués à reprendre** : Inquirer, Manila Bulletin, PCO (403) ; dti.gov.ph (réseau) ; bai.gov.ph et peza.gov.ph (TLS, 503) ; API Overpass.
