# L’éditeur HTML

L’éditeur HTML ouvre les .html / .htm avec deux modes : **Aperçu** (la page rendue) et **Source**.

- **Aperçu** : rendu réel ; les feuilles de style et images relatives se chargent à côté du fichier.
- **Inspecteur d’aperçu** : cliquez pour sélectionner un élément, double-cliquez pour modifier son texte sur place, supprimez via la barre d’outils, et demandez à l’IA au sujet de la sélection.
- **Mode source** : modifiez le HTML ; ctrl+F pour rechercher, Remplacer tout enregistre le balisage réécrit.
- **Enregistrement** : fidèle octet pour octet (BOM/CRLF/saut de ligne final préservés) ; un enregistrement sans modification ne réécrit rien.
- **Zoom** : ctrl+molette / pincement met le résultat à l’échelle de l’aperçu ; ctrl+Z dans l’aperçu annule la dernière modification.

## La barre d’outils

Cliquez sur n’importe quel élément dans l’aperçu et une barre d’outils flotte au-dessus :

![La barre d’outils flottante sur un élément sélectionné](img/html-toolbar.png)

- **Fichier et historique** : Enregistrer, Enregistrer sous, Annuler, Rétablir, Rechercher ; l’interrupteur **Enregistrement automatique** écrit les modifications à intervalles réguliers.
- Bascule **Aperçu / Source** ; **Plein écran** affiche la page en plein écran.
- **Mise en forme** : gras, italique, agrandir/réduire la taille de police ; le **panneau de style** de l’élément sélectionné (couleurs et plus encore).
- **Insertion** : titre, paragraphe, tableau, image (par lien), et plus.
- **Actions sur l’image** (avec une image sélectionnée) : rogner, **Supprimer l’arrière-plan**, remplacer, verrouiller les proportions.
- **Actions sur l’élément** (avec un élément sélectionné dans l’inspecteur d’aperçu) : supprimer, dupliquer, monter/descendre.
- **Bouton IA** : ouvre le panneau IA ; demandez directement au sujet de l’élément sélectionné.
