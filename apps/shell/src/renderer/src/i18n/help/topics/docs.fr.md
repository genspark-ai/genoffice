# Docs : traitement de texte

Docs est le processeur de texte analogue à Word : il lit et écrit de vrais .docx, avec une pagination en WYSIWYG fidèle.

## Le ruban

Onglets : **Accueil / Insertion / Dessin / Mise en page / Création / Références / Révision / Affichage**, plus des onglets contextuels pour l’objet sélectionné (Création de tableau et mise en page, Format de l’image, en-tête et pied de page).

- **Accueil** : presse-papiers ; police (y compris tailles et signes d’emphase est-asiatiques) avec **Effacer toute la mise en forme** et une case à cocher pour **Afficher/masquer les marques de mise en forme** ; paragraphe (alignement/retrait/interligne/listes) plus **Définir une nouvelle puce / Définir un nouveau format de numérotation / Liste à plusieurs niveaux**, qui enregistrent vos propres styles de liste dans le document ; styles (Titre 1-6/Normal/Citation, modifiables) avec un **Volet Styles** pour la liste complète.
- **Insertion** : sauts de page et de section ; tableaux (une grille lignes × colonnes, ou **Insérer un tableau…** pour une taille exacte) ; images, formes, zones de texte ; **Page de garde** et **Page vierge** issues d’une galerie prédéfinie ; **Graphique** ; **Lettrine** ; **WordArt** ; champs (date, heure, numéro de page, nombre total de pages, nom de fichier) ; hyperliens, **Signet** et renvois ; commentaires ; en-tête et pied de page et numéros de page ; symboles et équations.
  - **Graphique** insère un véritable objet graphique avec ses propres données — en barres, en courbes ou en secteurs — et non une image. _Modifier les données_ de Word ouvre les chiffres qui sont derrière.
- **Dessin** : encre sur la page, dans le groupe **Outils de dessin** — **Sélectionner** revient à l’édition du texte, puis **Stylet**, **Surligneur** et **Gomme** (un clic ou un balayage supprime tout le trait). À côté, **Style de stylet** / **Style de surligneur** est une seule commande qui contient des pastilles de couleur et une rangée d’épaisseurs ; son libellé suit l’outil actif. L’encre est enregistrée dans le document comme une annotation flottant au-dessus du texte, elle survit donc à l’enregistrement et à la réouverture, et **Tout effacer** dans le groupe suivant la supprime entièrement.
- **Mise en page** : marges, orientation et format du papier, colonnes, retraits et espacements de paragraphe.
- **Création** : thèmes, jeux de couleurs, filigrane, bordures de page.
- **Références** : table des matières (actualisable), notes de bas de page et notes de fin, légendes, renvois.

  ![L’onglet Références](img/docs-references.png)

- **Révision** : **Éditeur** relit tout le document — orthographe, grammaire et ponctuation ; **Traduire** ; vérification orthographique ; commentaires (**Commentaires par IA** traite ceux qui sont ouverts) ; suivi des modifications avec les vues Toutes les marques / Marques simples, accepter/refuser, et **Résumé IA des révisions** ; comptage des mots ; **Comparer** avec un autre fichier ; **Protéger le document**.
- **Affichage** : cinq façons de regarder le fichier — **Page**, **Web**, **Plan**, **Mode Lecture** et **Aperçu des pages** ; zoom arrière/avant/100 %/largeur de la page/une page ; **Volet IA** ; **Mode sombre** ; règle, quadrillage et volet de navigation ; **Nouvel onglet**, **Fractionner** et **Changer d'onglet** ; la boîte de dialogue de **raccourcis clavier** consultable.
  - **Mode sombre** assombrit la page et le canevas autour d’elle, jamais le ruban — la séparation de Word entre une surface d’édition sombre et une fenêtre sombre. Le choix est mémorisé et l’emporte sur le thème de l’application dans tous les cas.
  - **Fractionner** ouvre un second volet en dessous qui défile indépendamment et reflète le premier ; fermez-le avec le × sur son bord.

## Le volet de navigation

**Affichage ▸ volet de navigation** ouvre un volet latéral qui contient le plan des titres du document, un champ de recherche sur l’ensemble du document et une vignette par page. S’il est ouvert ou non est mémorisé entre les lancements : un document que vous parcourez par son plan le reste navigable.

**Le plan** est l’arborescence des titres. Un clic droit sur un titre du plan sert à le replier ou à le restructurer, pas seulement à naviguer :

- **Réduire / Développer** sur un titre replie toute sa sous-arborescence — le chapitre disparaît, son texte reste dans le document.
- **Tout réduire / Tout développer** replie ou déploie l’ensemble d’un seul coup. Sur un long rapport, c’est la différence entre un plan lisible et un mur de texte.
- **Afficher les niveaux de titre** filtre l’arborescence sur les profondeurs qui vous intéressent, si bien que _Afficher le titre 1_ vous laisse une table des matières que vous pouvez réellement parcourir.
- **Promouvoir / Abaisser** changent le niveau du titre, et avec lui le niveau qu’héritent tous les titres en dessous — comme un chapitre qui devient une section.
- **Nouveau titre avant / après** en insère un à l’emplacement du curseur, sans quitter le volet.
- **Supprimer** retire le titre _et tout ce qu’il contient_ : c’est celui-là auquel il faut faire attention, il supprime une sous-arborescence, pas une ligne.
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
