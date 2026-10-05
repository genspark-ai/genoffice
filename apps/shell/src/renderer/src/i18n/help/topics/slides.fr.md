# Slides : présentations

Slides est l’éditeur analogue à PowerPoint : il lit et écrit de vrais .pptx.

## L’interface

- **Ruban** : l’onglet Accueil porte l’insertion et la mise en forme (zone de texte/forme/image/tableau/graphique), la police et le paragraphe, l’alignement et la disposition (ordre de superposition, alignement, répartition).
- **Rail de miniatures** (gauche) : cliquez pour changer, faites glisser pour réordonner, clic droit pour nouveau/dupliquer/supprimer.
- **Canevas** : édition en WYSIWYG ; glissement, poignées de redimensionnement, repères.
- **Notes** : notes du présentateur par diapositive, conservées lors des exports.

## Onglets du ruban

La barre d’onglets (macOS commence à Accueil ; Windows ajoute un onglet Fichier) :

- **Accueil** : insertion et mise en forme — zone de texte, formes, images, tableaux, graphiques ; police et paragraphe ; alignement et disposition (ordre/alignement/répartition) ; disposition de page.
- **Insertion** : zone de texte, tableau (lignes et colonnes au choix), images, numéro de page, un bouton de saut (cliquez dessus pendant la présentation pour aller à une diapositive) et plus encore.
- **Dessin** : le **stylet** (dessiner à la main sur la diapositive, enregistré comme encre sur la page) et le **surligneur** (traits translucides et plus épais), avec l’épaisseur du trait ; cliquez de nouveau sur l’outil pour annuler.
- **Création** : thèmes, jeu de couleurs et arrière-plan ; masques et dispositions.
- **Transitions** : choisissez une transition pour la diapositive courante (elle prend effet dans le mode Présentateur de PowerPoint), avec application à toutes ; Aucune la supprime.

  ![L’onglet Transitions](img/slides-transitions.png)

- **Animations** : effets d’entrée et d’accentuation pour la forme sélectionnée, **Trajectoires** (se déplacer le long d’un chemin) ; **Aperçu** joue les animations de la diapositive sur le canevas ; Aucune les supprime.

  ![L’onglet Animations](img/slides-animations.png)

  Essayez : sélectionnez la zone de texte du titre ▸ onglet Animations ▸ choisissez un effet d’entrée ▸ **Aperçu** le joue sur le canevas.

- **Diaporama** : présenter depuis le début ou depuis la diapositive courante, plus les paramètres de présentation.
- **Révision** : **nouveau commentaire** sur la diapositive courante (écrit dans le pptx, visible dans PowerPoint).
- **Affichage** : **Normal** (miniatures + canevas), **Mode Plan** (parcourir et sauter par le texte), **Trieuse de diapositives** (vue d’ensemble en grille, double-clic pour éditer), **Mode Lecture** (plein écran, diapositive par diapositive ; Échap pour quitter).

## Menus contextuels

Le menu dépend de l’endroit du clic droit :

- **Sur le canevas vide** : nouvelle diapositive (disposition au choix), couper/copier/coller la diapositive (aussi **coller comme image** et **coller en conservant la mise en forme source**), dupliquer/supprimer/masquer la diapositive, **ajouter une section** (avant), renommer/déplacer vers le haut ou le bas/supprimer/replier tout/déplier tout la section, **mise en forme de l’arrière-plan** / **changer l’image d’arrière-plan**, réinitialiser la disposition.
- **Sur un élément sélectionné** : modifier le texte, lien hypertexte, taille et position, **aligner** (gauche/centrage horizontal/droite/haut/centrage vertical/bas) et répartir horizontalement et verticalement, mettre au premier plan / envoyer à l’arrière-plan, grouper/dégrouper/regrouper, miroir horizontal/vertical, rogner l’image / remplacer l’image / enregistrer comme image, changer la forme, **définir comme forme par défaut**, modifier les points ; les éléments de texte ont aussi gras/italique/souligné/puces/numérotation.
- **Sur un tableau sélectionné** : insérer lignes/colonnes (au-dessus/en dessous/à gauche/à droite), supprimer lignes/colonnes, fusionner / fusionner à droite / fusionner vers le bas, fractionner une cellule, **trame de fond de cellule**, ancrage du contenu de la cellule (haut/milieu/bas).

## Contenu

- Zones de texte, formes (remplissage/contour/ombre), images (rogner/remplacer), tableaux, graphiques (colonnes/barres/lignes/secteurs…, données modifiables).
- Masques et dispositions : unifier polices, espaces réservés et arrière-plans ; les nouvelles diapositives héritent de la disposition choisie.
- Les couleurs et polices du thème suivent le thème.

## Génération par IA

- La carte AI Slides de l’accueil : donnez un sujet ou un plan et l’IA construit le diaporama ; si la génération dans le cloud échoue, elle bascule sur la génération locale.
- Continuez à ajuster ensuite avec le panneau IA (l’IA remanie la mise en forme via un script isolé et contrôlé — le même mécanisme que dans [Sheets](help://sheets)).

## Présentation et export

- **Exporter en PDF** : rastérisation page par page dans une fenêtre masquée, avec progression et garde-fou de délai pour les gros diaporamas.
- Export d’images : un PNG par page.
- Présentation en plein écran là où la version le permet.

## Enregistrement

Aller-retour complet en .pptx : les pages non modifiées restent identiques octet pour octet ; les notes, les masques et les annotations sont conservés.
