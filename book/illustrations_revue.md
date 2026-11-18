# Illustrations IA : fiche de relecture et cahier des charges pour un illustrateur

Le livre contient 15 illustrations créées avec un générateur d'images (Figma, novembre 2026). Elles sont marquées « Illustration » dans les légendes et créditées « Illustration created for this book (AI-generated, Figma). Artist's impression ». Ce ne sont pas des documents historiques.

**Règle suivie.** On garde une illustration seulement si aucun document ne peut montrer la scène : préhistoire, époque précoloniale, scènes de la vie courante que personne n'a photographiées. Une illustration ne doit jamais montrer un fait que le texte ne prouve pas.

## Retirées lors de la relecture du 16 novembre 2026

| Fichier | Raison |
|---|---|
| `liberation-1945.jpg` (drapeau hissé sur la place, 1945) | La scène du drapeau n'est attestée par aucune source ; le texte l'a d'ailleurs retirée. Les guérilleros portent des casques de type américain. |
| `town-hall-1955.jpg` (inauguration de l'hôtel de ville, 1955) | Le bâtiment réel est connu par des photos (Sentrong Pangkultura) : il ne ressemble pas à l'image. L'église du fond a un dôme, ce qui est faux. |
| `ondoy-flood-2009.jpg` (barangay inondé, 2009) | Événement récent, abondamment photographié : une image inventée pourrait passer pour un document. Mieux vaut une vraie photo, avec autorisation. |

## À faire vérifier par quelqu'un de Biñan

Pour chaque image, la question est la même : un habitant ou un historien local y voit-il une erreur (bâtiment, vêtement, objet, paysage) ?

| Ch. | Fichier | Ce que montre l'image | Ce qui est sûr | Ce qui est inventé, à vérifier | Priorité pour un dessin réel |
|---|---|---|---|---|---|
| I | `salt-lagoon-6000-years.jpg` | Rivage d'eau salée, mangroves, ramasseurs de coquillages | Bras de mer il y a 6 600–4 700 ans (sources du chapitre) | Les personnages et le paysage exact | basse |
| II | `homo-luzonensis.jpg` | *Homo luzonensis* à l'entrée d'une grotte | Espèce connue par des dents et des os | Toute l'apparence (la légende le dit) | basse |
| II | `austronesian-arrival.jpg` | Arrivée de navigateurs austronésiens | Pirogues à balancier, poteries rouges, cochons, chiens | Visages, vêtements | basse |
| III | `lakeshore-village.jpg` | Village sur pilotis au bord du lac, vers l'an 1000 | Maisons sur pilotis, pêche, taro | Détails des maisons | basse |
| III | `chinese-trade.jpg` | Commerce avec des marchands chinois | Récit de Zhao Rugua (1225) : ombrelle blanche, porcelaine | Jonque, costumes | basse |
| III | `catalonan-feast.jpg` | Rituel aux *anito* dirigé par une *catalonan* | Descriptions de 1572 et 1589 | Vêtements simplifiés (dit dans la légende) | moyenne |
| III | `salcedo-galley-1571.jpg` | Galère de Salcedo arrivant dans un village du lac | Expédition de 1571 | Forme de la galère, armures | moyenne |
| IV | `dominican-hacienda.jpg` | L'hacienda dominicaine : rizières, barrage, moulin à canne | Barrages de Timbao, riz, canne | Bâtiments, disposition | haute : un illustrateur local pourrait s'appuyer sur le barrage réel |
| IV | `binan-plaza-1870.jpg` | Une place de la Laguna vers 1870 | Église de pierre, bahay-na-bato, marché | L'église n'est pas celle de Biñan (dit dans la légende) | haute : un dessin d'après des photos anciennes de Biñan serait bien meilleur |
| IV | `cascos-on-the-lake.jpg` | Les *cascos* (barges couvertes) sur le lac | Barges de transport vers Manille | Proportions, voiles | moyenne : il existe des gravures du XIXᵉ siècle (domaine public) qui pourraient la remplacer |
| V | `us-troops-laguna-1900.jpg` | Colonne américaine sur une route de la Laguna | Campagne de janvier 1900 | Uniformes (chemises bleues), fumée d'un village | moyenne : il existe des photos de 1899–1900 dans le domaine public |
| V | `thomasite-classroom.jpg` | Classe vers 1905, institutrice américaine | Écoles américaines dans la Laguna dès 1904 | Tout le détail | moyenne : des photos d'époque existent (domaine public) |
| V | `binan-station-1908.jpg` | Arrivée du train à Biñan, 1908, vendeurs de puto | Ligne ouverte le 2 septembre 1908 | Gare, locomotive, quai | haute : une photo ancienne de la gare de Biñan, si elle existe, serait précieuse |
| VI | `occupation-rice-1943.jpg` | Réquisition du riz par l'armée japonaise | Crise du riz, BIBA (sources du chapitre) | Camion, affiches, foule | moyenne |
| VI | `guerrillas-1944.jpg` | Réunion nocturne de guérilleros | Réseaux de guérilla, « home guards » | Toute la scène | basse |

## Cahier des charges pour un illustrateur local

Si un artiste de Biñan accepte de redessiner certaines scènes (priorité « haute »), voici ce qu'il faudrait lui donner :

1. **La page du livre** où l'image apparaît, avec le texte du paragraphe et la légende.
2. **Les documents réels** : photos du barrage de Timbao, de l'église (et des photos d'avant l'incendie de 1968 si la paroisse en a), de l'ancienne casa hacienda (Sentrong Pangkultura), cartes anciennes du chapitre IV, photo de la route de 1934 (chapitre V).
3. **Les contraintes** : format paysage, environ 3:2 ; un style cohérent pour tout le livre (dessin à l'encre et lavis, ou aquarelle) ; aucune inscription dans l'image ; vêtements et objets de l'époque, vérifiés avec le musée de la ville (BCHATO).
4. **Les droits** : l'artiste cède le droit de reproduction pour l'édition électronique gratuite et l'édition papier à prix coûtant, avec son nom dans les crédits.

Remplacer une image IA par un dessin réel ne demande qu'un changement dans `book/build_pdf.py` (champ `src` et `credit` de la figure).
