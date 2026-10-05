# Opérations sur les fichiers : renommer, supprimer, exporter

Ce chapitre couvre les opérations sur les fichiers communes à tous les éditeurs ; les options d’export propres à chaque éditeur sont dans son chapitre.

Le **menu ⋯** de la ligne (survolez une ligne de fichier) rassemble ces actions :

![Le menu ⋯ d’une ligne de fichier](img/file-ops.png)

## Renommer

Deux points d’entrée, un même ensemble de contrôles :

- **⋯ ▸ Renommer** sur une ligne de l’accueil.
- **Double-cliquez sur un onglet de fichier** pour renommer sur place (voir [Onglets et gestion des fenêtres](help://tabs-and-windows)).

Règles : l’extension est conservée automatiquement ; les caractères interdits, les points finais et les noms réservés (CON/NUL et consorts) sont refusés avec un message ; un conflit de nom dans le même dossier est également bloqué. Le vrai fichier sur le disque est renommé, et les récents et les favoris suivent.

## Supprimer

- **⋯ ▸ Supprimer** sur l’accueil : déplace le fichier vers la **corbeille du système**, restaurable depuis l’OS.
- Une notification de suppression avec annulation apparaît pendant quelques secondes — annuler remet le fichier à sa place.

## Dupliquer

**⋯ ▸ Dupliquer** crée une copie nommée « <name> copy » dans le même dossier et l’ouvre dans un nouvel onglet ; les collisions reçoivent automatiquement un compteur.

## Enregistrer et Enregistrer sous

- **⌘S / ctrl+S** enregistre le fichier courant ; un fichier sans titre demande d’abord l’emplacement et le nom.
- **Enregistrer sous** écrit un nouveau fichier et laisse l’original intact ; les modifications suivantes ciblent le nouveau fichier.
- Chaque enregistrement est atomique (fichier temporaire puis renommage) ; quitter en cours d’écriture ne peut pas corrompre le fichier.
- L’enregistrement automatique n’intervient qu’après la première sauvegarde manuelle (voir [Démarrage rapide](help://getting-started)).

## Export en PDF

- **Docs** : Fichier ▸ Exporter en PDF (ou le bouton du ruban), paginé selon la mise en page courante.
- **Slides** : l’export rastérise page par page, avec progression pour les gros diaporamas.
- **Sheets** : l’export suit la pagination d’impression.
- Les exports sont rendus dans une fenêtre masquée et arrivent là où vous le choisissez.

## Export en Word / images

- **PDF ▸ PDF en Word** : transforme le PDF en .docx (conversion locale ; les mises en page complexes sont au mieux préservées).
- **Docs** peut exporter les pages en images (un PNG par page).

## Impression

Fichier ▸ Imprimer dans chaque éditeur (⌘P/ctrl+P) ouvre la boîte de dialogue d’impression du système ; les PDF s’impriment avec l’ordre de pages et les rotations courants.

## Où vivent les fichiers sans titre

L’emplacement que vous choisissez à la première sauvegarde est son domicile ; avant cela, le document n’existe qu’en mémoire. L’enregistrement automatique ne prend le relais qu’après cette première sauvegarde.
