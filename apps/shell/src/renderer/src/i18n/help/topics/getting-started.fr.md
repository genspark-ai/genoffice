# Démarrage rapide : l’interface et les bases

GenOffice est une suite bureautique qui fonctionne entièrement sur votre machine : une fenêtre, une rangée d’onglets, six éditeurs — Docs (traitement de texte), Sheets (tableurs), Slides (présentations), PDF, Markdown et HTML. Les fichiers sont de vrais .docx / .xlsx / .pptx / .pdf, totalement interchangeables avec Word, Excel et PowerPoint. Aucun réseau requis.

## Aperçu de l’interface

![L’écran d’accueil](img/home-screen.png)

La fenêtre comporte trois parties :

- **Barre d’onglets (en haut)** : chaque fichier ouvert correspond à un onglet. L’onglet Accueil le plus à gauche est toujours présent et ne peut pas être fermé ; les autres onglets sont vos documents. Un double-clic sur un onglet permet de renommer le fichier sur place.
- **Zone de contenu** : l’éditeur (ou l’accueil) correspondant à l’onglet actif.
- **Menu** : dans la barre de menus système sur macOS, en haut de la fenêtre sur Windows/Linux. Les menus Fichier/Édition/Affichage s’adaptent à l’éditeur actif.

## Créer un document

Chacun de ces moyens convient :

- Cliquez sur une carte de création rapide de la section [Démarrage rapide](help://getting-started) de **Accueil** (AI Docs, AI Sheets, AI Slides, …).
- Menu **Fichier ▸ Nouveau** : Docs (⌘N/ctrl+N), Sheets, Slides, Markdown, HTML ou PDF.
- Faites glisser un fichier sur la fenêtre, ou double-cliquez dessus dans votre gestionnaire de fichiers (si GenOffice est l’application par défaut).

Un nouveau document s’ouvre sans titre ; le fichier sur le disque n’est créé qu’à la première sauvegarde.

## Ouvrir des fichiers

- Le menu **Fichier ▸ Ouvrir** (⌘O/ctrl+O) ouvre le sélecteur du système : .docx / .doc / .xlsx / .xls / .csv / .tsv / .pptx / .ppt / .pdf / .md / .html.
- Cliquez sur n’importe quel élément de la liste **Récents** de l’accueil.
- `genoffice <file>` depuis un terminal ouvre aussi les fichiers.

## Le modèle d’enregistrement

- **Enregistrement manuel** : ⌘S/ctrl+S, ou Fichier ▸ Enregistrer / Enregistrer sous. Le premier enregistrement demande l’emplacement et le nom.
- **L’enregistrement automatique** ne s’active qu’après avoir enregistré le fichier manuellement au moins une fois — un PDF que vous ne faites que lire n’est jamais réécrit en silence. Il se déclenche peu après une modification du contenu.
- Fermer un onglet avec des modifications non enregistrées demande d’abord Enregistrer / Supprimer / Annuler.
- Chaque écriture est atomique (fichier temporaire puis renommage) : une coupure de courant ne peut pas laisser un fichier à moitié écrit.

## Raccourcis courants

| Action           | macOS | Windows / Linux |
| ---------------- | ----- | --------------- |
| Nouveau document | ⌘N    | ctrl+N          |
| Ouvrir           | ⌘O    | ctrl+O          |
| Enregistrer      | ⌘S    | ctrl+S          |
| Fermer l’onglet  | ⌘W    | ctrl+W          |
| Ouvrir ce manuel | F1    | F1              |
| Réduire le ruban | ⌥⌘R   | Ctrl+F1         |

**La réduction du ruban** fonctionne dans tous les éditeurs. La rangée d'onglets reste en place et la bande de commandes qui se trouve en dessous disparaît ; l'onglet sélectionné sert aussi de commande de réduction : tant que le ruban est réduit, aucun onglet n'est sélectionné, et appuyer sur n'importe quel onglet fait revenir la bande. Un double-clic sur un onglet fait de même. La façon dont vous l'avez laissée est mémorisée éditeur par éditeur.

Les raccourcis propres à chaque éditeur (copier la mise en forme, rechercher et remplacer, opérations sur les tableaux, …) sont dans leurs chapitres ; Docs fournit en plus une boîte de dialogue de raccourcis clavier consultable (**⌘/**) (voir son chapitre).

## Les raccourcis Option+Commande

Option+Commande est la couche que Word réserve aux sauts structurés, et GenOffice la remplit de la même façon. Docs en prend l'essentiel, Sheets en prend deux à lui pour la parité avec Excel ; un raccourci fonctionne partout.

**Docs**

| Raccourci (macOS) | Effet                           | Windows / Linux    |
| ----------------- | ------------------------------- | ------------------ |
| ⌥⌘1 / ⌥⌘2 / ⌥⌘3   | Titre 1 / 2 / 3                 | Ctrl+Alt+1 / 2 / 3 |
| ⌥⌘0               | Normal                          | Ctrl+Alt+0         |
| ⌥⌘M               | Paragraphe                      | Ctrl+Alt+M         |
| ⌥⌘A               | Nouveau commentaire             | Ctrl+Alt+A         |
| ⌥⌘F               | Insérer une note de bas de page | Ctrl+Alt+F         |
| ⌥⌘E               | Insérer une note de fin         | Ctrl+Alt+D         |
| ⌥⌘G               | Atteindre                       | Ctrl+G             |

Deux d'entre eux changent sur Windows, pour la même raison qui fait que Word les sépare. **macOS possède ⌥⌘D** — il affiche et masque le Dock —, si bien que la note de fin est ⌥⌘E sur le Mac et Ctrl+Alt+D partout ailleurs. Et **Atteindre** laisse tomber l'Alt : Ctrl+G, là où le raccourci du Mac le conserve.

**Sheets**, quand la grille a le focus

| Raccourci (macOS) | Effet                | Windows / Linux |
| ----------------- | -------------------- | --------------- |
| ⌥⌘0               | Bordures extérieures | Ctrl+Shift+7    |
| ⌥⌘−               | Aucune bordure       | Ctrl+Shift+−    |

Windows n'est pas une réécriture de la paire Mac. Excel pour Mac donne à Sheets **les deux** : ⌘⇧7 et ⌥⌘0 sont deux touches pour les mêmes bordures extérieures, si bien que sur Windows la commande garde la case Ctrl+Shift qu'elle avait déjà, et la couche Option est simplement absente.

Remarquez que **⌥⌘0 signifie Normal dans Docs et Bordures extérieures dans Sheets**. Les deux n'apparaissent jamais dans le même éditeur, donc rien ne se percute à l'usage, mais ⌥⌘0 est déjà pris et n'est pas disponible comme raccourci global.

**Tous les éditeurs** : **⌥⌘R / Ctrl+F1** réduit le ruban, comme décrit plus haut.

Cela laisse ⌥⌘D libre pour que GenOffice l'utilise sur macOS si une future commande en a besoin.

## Pour aller plus loin

- Où se trouvent les fichiers : [L’écran d’accueil](help://home-screen).
- Gérer beaucoup de fichiers ouverts : [Onglets et gestion des fenêtres](help://tabs-and-windows).
- Confier le travail à l’IA : [Le panneau d’assistant IA](help://ai-panel).
- Langue, thème, applications par défaut : [Paramètres, langue, thème et intégrations MCP](help://settings-integrations).
