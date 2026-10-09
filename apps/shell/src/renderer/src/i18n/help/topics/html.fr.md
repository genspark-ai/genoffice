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

## Exportation

Menu Fichier, tout est local et tout demande où mettre le résultat :

- **Exporter en Word…** et **Exporter en PDF…** écrivent un vrai .docx ou .pdf.
- **Exporter en HTML (fichier unique)…** écrit un seul .html avec les images incorporées. Il n’écrasera pas le fichier que vous avez ouvert et il indique combien d’images il n’a pas pu incorporer.

## Insérer un squelette

Pour une page vide, **Insertion ▸ Insérer un squelette** écrit un document minimal en mode standards :

```html
<!DOCTYPE html>
<html lang="en">
  <head>
    <meta charset="UTF-8" />
    <title></title>
  </head>
  <body></body>
</html>
```

Chaque partie est là pour une raison, ce qui explique que ce soit une commande plutôt que quelque chose à taper :

- le **doctype**, sinon l’aperçu tourne en mode quirks, où le dimensionnement des boîtes et la mise en page des tableaux suivent d’autres règles que celles que vous attendez ;
- le **`lang`**, sinon un lecteur d’écran n’a pas de langue pour lire la page et le navigateur choisit une police et un correcteur orthographique pour la mauvaise ;
- le **charset**, sinon une page en écriture non latine peut se décoder en caractères déformés.

La balise meta de viewport est volontairement absente : ceci se rend dans un volet de bureau, sans viewport mobile auquel elle s’appliquerait.

Le `lang` suit la langue d’interface de l’application : le squelette que vous insérez est donc celui pour lequel votre outillage est déjà configuré. Modifiez-le librement ensuite.

L’élément n’apparaît qu’en mode édition, et seulement tant que le document est vide — dès qu’il y a du contenu, il n’y a plus rien _dans quoi_ insérer un squelette.
