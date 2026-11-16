# Avant publication : liste de contrôle

Objectif : une édition électronique gratuite (PDF, éventuellement EPUB) et une édition papier vendue à prix coûtant. État au 16 novembre 2026 : livre de 200 pages environ et compagnon de 86 pages, en brouillon.

## 1. Le texte

- [ ] **Relecture du tagalog par un locuteur natif, de préférence de Biñan.** Mon tagalog est correct mais pas natif ; la relecture précédente (`relecture_tagalog.md`) avait trouvé des fautes. Points à surveiller : registre (tagalog courant ou filipino soutenu), termes techniques (géologie, droit foncier), cohérence des titres de section avec le compagnon.
- [ ] **Relecture de l'anglais** par un lecteur extérieur, pour les phrases trop longues et les répétitions.
- [ ] **Faits encore fragiles** (marqués 🟡 ou 🔴 dans `book/research/*.md`) : l'issue de l'affaire Alonte après 1998 ; la liste complète des maires avant 1980 ; la date d'ouverture de la SLEX à Biñan ; Marcial Alintanahin ; les porcheries près du lac ; le nombre de pêcheurs ; les tarifs et couleurs des tricycles ; la Semaine sainte ; la loi martiale.
- [ ] **Sources faibles restantes** : quelques citations de Wikipédia ou de blogs (liste dans `book/research/sources_remplacement.md`).
- [ ] **Entretiens** : intégrer les témoignages recueillis avec `guide_entretiens.md`.
- [ ] **Licence du texte** : à décider. Le projet est public sur GitHub ; une licence libre (par exemple CC BY-SA 4.0, compatible avec les photos et les cartes sous cette licence) permettrait aux écoles et aux bibliothèques de le copier. Sinon, écrire « Tous droits réservés » dans le colophon.

## 2. Les images

- [ ] **Droits à obtenir (6 images)** — voir `droits_a_obtenir.md` :
  - 5 photos d'Experience Biñan (BCHATO) : Puto Biñan, pospas, maglalatik, Puto Latik Festival, Araw ng Biñan la nuit ;
  - la photo de la route du Sud en 1934 (Robert L. Pendleton, American Geographical Society Library).
  Sans accord, `python3 book/build_pdf.py --final` les retire automatiquement.
- [ ] **Illustrations IA** : faire relire les 15 images restantes par quelqu'un de Biñan, et si possible faire redessiner les plus importantes par un illustrateur local (`illustrations_revue.md`).
- [ ] **Photos manquantes** : un tricycle de Biñan, la rive de Malaban, une rizière, un atelier, le marché, le festival. Les plus simples : les prendre soi-même (et les publier sous licence libre).
- [ ] **Crédits** : vérifier que chaque photo CC BY ou CC BY-SA est créditée (page « Credits ») ; les cartes qui contiennent des failles GEM ou des données OpenStreetMap portent la mention de leur licence.

## 3. L'édition électronique (gratuite)

- [ ] Version finale : `python3 book/build_pdf.py --final`, puis relire le PDF sur un téléphone (les notes de bas de page sont des liens cliquables).
- [ ] Optionnel : une version EPUB, plus lisible sur téléphone, peut être produite à partir des mêmes fichiers Markdown avec pandoc.
- [ ] Diffusion : page GitHub du projet (Releases), Internet Archive, bibliothèque de la ville et musée (BCHATO), écoles de Biñan.

## 4. L'édition papier (à prix coûtant)

- [ ] **ISBN.** Aux Philippines, l'agence ISBN est la Bibliothèque nationale (National Library of the Philippines) ; en France, c'est l'AFNIL. Les deux attribuent les ISBN aux éditeurs de leur pays : à vérifier selon le lieu de publication. Prévoir un ISBN différent pour le livre, le compagnon et, s'il y en a un, l'EPUB.
- [ ] **Dépôt légal** : aux Philippines, auprès de la Bibliothèque nationale ; en France, à la BnF. À vérifier selon le pays d'édition.
- [ ] **Format** : 6 × 9 pouces (152 × 229 mm), déjà réglé. Intérieur en couleur : c'est le poste le plus cher en impression à la demande ; comparer avec une version intérieure en noir et blanc (les cartes restent lisibles) et un cahier couleur.
- [ ] **Fichier d'impression** : la plupart des imprimeurs demandent un PDF/X avec images en CMJN et 300 dpi. WeasyPrint produit un PDF standard : prévoir une conversion (par exemple avec Ghostscript) et un essai d'impression.
- [ ] **Couverture** : le PDF actuel contient une couverture avant et arrière au format page. Pour l'impression, il faut un fichier de couverture séparé avec le dos (largeur selon le nombre de pages et le papier) et 3 mm de fond perdu.
- [ ] **Imprimeur** : comparer l'impression à la demande (plateformes en ligne) et un imprimeur de Manille ou de la Laguna pour un tirage court, distribué sur place.
