# L’éditeur Markdown

L’éditeur Markdown ouvre les .md / .markdown avec une expérience source + aperçu rendu.

- **Ouvrir** : depuis l’accueil ou Fichier ▸ Ouvrir ; la ligne de commande fonctionne aussi.
- **Modifier** : édition en texte brut ; les extensions GFM (tableaux, listes de tâches, barré, liens automatiques) sont rendues dans l’aperçu.
- **Aperçu** : en direct ; les ressources relatives comme les images se résolvent à côté du document.
- **Enregistrement** : fidèle octet pour octet — le BOM, les CRLF et la présence ou non d’un saut de ligne final sont préservés ; un enregistrement sans modification ne réécrit pas le fichier.
- **Rechercher et remplacer** : ctrl+F cherche dans la source ; Remplacer tout réécrit.
- **IA** : des boutons prédéfinis permettent à l’assistant de réécrire, de développer ou de traduire le document.

## La barre d’outils

Une rangée de boutons au-dessus de l’éditeur (survolez pour les infobulles) :

![La barre d’outils Markdown](img/md-toolbar.png)

- **Fichier et historique** : Enregistrer, Enregistrer sous, Annuler, Rétablir, Rechercher ; l’interrupteur **Enregistrement automatique** à droite écrit les modifications sur le disque à intervalles réguliers.
- **Bouton IA** : ouvre le panneau IA ; à côté se trouvent les prédéfinis réécriture / développement / traduction.
- **Style de paragraphe** (menu déroulant) : passer du corps de texte aux niveaux de titre.
- **Mise en forme en ligne** : **Gras**, _Italique_, ~~Barré~~, `Code en ligne`, lien.
- **Listes** : à puces, numérotée, liste de tâches.
- **Insertion** : tableau, image, ligne de séparation.
- **Propriétés** : insère le bloc YAML front matter en tête de fichier ou y saute.
- **Plan** : navigation par hiérarchie de titres.
- **Orthographe** : active ou désactive la vérification orthographique pour ce document.

Trois exemples rapides :

- **Titre** : placez le curseur sur la ligne ▸ menu style de paragraphe ▸ « Titre 1 ».
- **Tableau** : cliquez sur **Insérer un tableau** ▸ faites glisser le nombre de lignes et de colonnes ▸ saisissez dans les cellules ; l’aperçu le rend immédiatement.
- **Liste de tâches** : sélectionnez quelques lignes ▸ cliquez sur **Liste de tâches** ▸ chaque ligne devient `- [ ]`, rendue comme des cases à cocher dans l’aperçu.
