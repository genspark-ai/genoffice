# Docs : traitement de texte

Docs est le processeur de texte analogue à Word : il lit et écrit de vrais .docx, avec une pagination en WYSIWYG fidèle.

## Le ruban

Onglets : **Accueil / Insertion / Mise en page / Création / Références / Révision / Affichage**, plus des onglets contextuels pour l’objet sélectionné (Création de tableau, images).

- **Accueil** : presse-papiers ; police (y compris tailles et signes d’emphase est-asiatiques) ; paragraphe (alignement/retrait/interligne/listes) ; styles (Titre 1-6/Normal/Citation, modifiables).
- **Insertion** : sauts de page et de section, tableaux (y compris dessinés et rapides), images, formes, hyperliens, en-tête et pied de page, numéro de page, date, zones de texte.
- **Mise en page** : marges, orientation et format du papier, colonnes, retraits et espacements de paragraphe.
- **Création** : thèmes, jeux de couleurs, filigrane, bordures de page.
- **Références** : table des matières (actualisable), notes de bas de page et notes de fin, légendes, renvois.

  ![L’onglet Références](img/docs-references.png)

- **Révision** : vérification orthographique, commentaires, suivi des modifications (vues Toutes les marques / Marques simples), comptage des mots.
- **Affichage** : règle, quadrillage, volet de navigation, zoom, et la boîte de dialogue de **raccourcis clavier** consultable.

## Le volet de navigation

**Affichage ▸ volet de navigation** ouvre un volet latéral qui contient le plan des titres du document, un champ de recherche sur l’ensemble du document et une vignette par page. S’il est ouvert ou non est mémorisé entre les lancements : un document que vous parcourez par son plan le reste navigable.

**Le plan** est l’arborescence des titres. Un clic droit sur un titre du plan sert à le replier ou à le restructurer, pas seulement à naviguer :

- **Réduire / Développer** sur un titre replie toute sa sous-arborescence — le chapitre disparaît, son texte reste dans le document.
- **Tout réduire / Tout développer** replie ou déploie l’ensemble d’un seul coup. Sur un long rapport, c’est la différence entre un plan lisible et un mur de texte.
- **Afficher les niveaux de titre** filtre l’arborescence sur les profondeurs qui vous intéressent, si bien que *Afficher le titre 1* vous laisse une table des matières que vous pouvez réellement parcourir.
- **Promouvoir / Abaisser** changent le niveau du titre, et avec lui le niveau qu’héritent tous les titres en dessous — comme un chapitre qui devient une section.
- **Nouveau titre avant / après** en insère un à l’emplacement du curseur, sans quitter le volet.
- **Supprimer** retire le titre *et tout ce qu’il contient* : c’est celui-là auquel il faut faire attention, il supprime une sous-arborescence, pas une ligne.
- **Sélectionner le titre et le contenu** sélectionne du titre jusqu’à la fin de sa sous-arborescence, prêt pour une édition de section entière.

## Menu contextuel

Clic droit n’importe où dans le corps — le menu correspond à ce que vous avez cliqué. Principaux groupes :

- **Presse-papiers** : couper / copier / coller / **collier en texte brut**.
- **Police, paragraphe** : changer la police et la taille, gras/italique/souligné, alignement/retrait/interligne sans passer par le ruban.
- **Synonymes** : liste des synonymes du mot sélectionné ; cliquez sur l’un pour remplacer.
- **Traduire** (IA) : traduit la sélection vers la langue cible (anglais, chinois simplifié, japonais, coréen, français, allemand, espagnol, …) via le panneau IA.
- **Nouveau commentaire** : attache un commentaire à la sélection.
- **Orthographe** (sur un mot mal orthographié) : remplacements suggérés, tout ignorer, ajouter au dictionnaire, définir la langue de correction.
- **Lien hypertexte** : ouvrir / modifier / copier le lien / supprimer le lien hypertexte.
- **Image** : afficher l’image, enregistrer l’image sous, **Habillage** (dans le texte / carré gauche et droite / haut et bas / derrière le texte / devant le texte), ordre de disposition.
- **Champs** (table des matières, numéros de page) : mettre à jour le champ / afficher les codes de champ / modifier le champ.
- **Numérotation** (dans une liste) : recommencer la numérotation / poursuivre la numérotation / modifier le niveau de liste / définir la valeur de numérotation.
- **Tableau** (curseur dans un tableau) : insérer lignes/colonnes, fusionner / fractionner les cellules, fractionner le tableau, ajuster automatiquement, alignement des cellules, répartir lignes/colonnes, propriétés du tableau, menu de suppression, sélection.

## Édition

- Rechercher et remplacer (ctrl+F / ctrl+H) : respect de la casse, mot entier, expressions régulières.
- Copier la mise en forme ; annuler/rétablir étendu ; options de collage.
- Tableaux : fusionner/fractionner les cellules, opérations sur les lignes et colonnes, bordures et trames, tri, formules.
- Images : habillage du texte, rognage, compression ; canevas de dessin.

## Typographie est-asiatique

- La compression des signes de ponctuation et les règles de césure (kinsoku) sont alignées sur Word ; conversion pleine largeur / demi-largeur.
- Les polices candidates couvrent les noms de familles est-asiatiques courants de Windows et de macOS.

## IA

- Bouton IA du ruban et panneau latéral : réécriture, développement, traduction, résumé, insertion de tableaux prédéfinis, plus des instructions libres.
- Chaque tour de l’IA enregistre d’abord un instantané ; vous pouvez revenir en arrière depuis la liste des versions, et ce retour en arrière est lui-même annulable.

## Enregistrement et export

- Enregistre du .docx en ne réécrivant que les paragraphes modifiés — le contenu intact reste identique octet pour octet.
- Exporte en PDF (selon la pagination en cours) et en images page par page.

## Impression

ctrl+P via la boîte de dialogue du système, pages en WYSIWYG.
