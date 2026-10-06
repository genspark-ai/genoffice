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

## Exportation

Menu Fichier, tout est local et tout demande où mettre le résultat :

- **Exporter en Word…** et **Exporter en PDF…** écrivent un vrai .docx ou .pdf.
- **Exporter en images…** écrit un PNG par page dans un dossier que vous choisissez.
- **Convertir et ouvrir dans Docs** convertit en .docx et l’ouvre dans l’onglet Docs intégré ici dans l’application — ce n’est pas un transfert vers quoi que ce soit dans le cloud, et la copie convertie vit dans un dossier de cache nettoyé au bout d’environ une semaine.

## La vue source

La barre d’outils porte un interrupteur **Source** (localisé avec l’application). Activez-le et l’éditeur cède la place au Markdown brut : exactement le texte qu’écrit un enregistrement, rien d’habillé, rien de normalisé sous vos pieds.

- **L’édition est fidèle octet pour octet.** Un enregistrement depuis la vue source produit les mêmes octets qu’un enregistrement depuis l’éditeur — le BOM, les CRLF et la présence d’un saut de ligne final survivent tous.
- **C’est le même document.** Basculez librement dans les deux sens ; la source est le texte de l’éditeur lui-même, pas une copie qu’il faudrait fusionner.
- **La barre de mise en forme est indisponible** tant que la vue est ouverte, parce que la plupart de ces boutons insèrent des constructions d’éditeur qui n’ont de sens que du côté rendu. Elle revient quand vous fermez la vue.
- **Les fichiers JSON et autres fichiers en mode source** s’ouvrent ici directement : il n’y a rien à rendre, donc la source _est_ le document.
