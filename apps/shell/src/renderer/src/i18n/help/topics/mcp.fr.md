# Connecter un agent de codage

GenOffice parle le Model Context Protocol : un agent de programmation peut donc lire, écrire et rendre vos documents avec les mêmes moteurs que l’application. L’agent ne devine pas un format de fichier : il reçoit les schémas typés des ops à partir des définitions mêmes contre lesquelles l’executor valide.

## L’enregistrer depuis l’application

C’est dans **Paramètres ▸ Intégrations** qu’on fait cela. Le panneau a deux moitiés, et vous pouvez utiliser l’une, l’autre, ou les deux.

**La compétence.** Une ligne par agent de programmation trouvé sur cette machine — Claude Code, Codex, Cursor, Gemini CLI, Copilot CLI, OpenCode, Windsurf —, chacune avec **Installer**, **Mettre à jour** et **Désinstaller**, plus **Installer dans un autre dossier…**, **Télécharger le skill (zip)** et **Copier le chemin**. Si votre assistant n’est pas dans la liste, indiquez à GenOffice un dossier depuis lequel il lit `SKILL.md`, ou enregistrez le zip et laissez l’assistant l’installer. La compétence et le MCP peuvent coexister : l’assistant choisit l’un des deux, et ils font exactement les mêmes choses.

**MCP.** Deux voies sont proposées : **Lancé par l’assistant (recommandé)**, où vous ajoutez la configuration affichée à votre client et où l’assistant lance le serveur lui-même, et le **Serveur HTTP local**, que l’application fait tourner pour vous. Dans les deux cas, l’assistant finit par parler à GenOffice et vous ne tapez jamais une commande.

## L’enregistrer depuis la ligne de commande

La même chose depuis un terminal — c’est la voie avancée, celle qu’il faut prendre quand l’agent se trouve là où le panneau ne peut pas le trouver :

```sh
genoffice mcp install all
```

Il trouve les agents de programmation présents sur cette machine et écrit l’entrée du serveur stdio dans la configuration de chacun, en laissant le reste du fichier tel qu’il l’a trouvé.

```sh
genoffice mcp list             # where each agent stands
genoffice mcp install cursor    # just one
genoffice mcp uninstall cursor  # take it back out
```

Un agent installé ailleurs demande `--dir <path>` ; `--force` réécrit une entrée déjà présente.

Tout ce que le serveur accepte tient sur un seul écran — les formes d’enregistrement, de retrait et de liste, le service en HTTP, et les deux options de schéma :

![La vraie sortie de genoffice mcp --help : les formes install, uninstall et list, avec les options --http, --host, --token, --compact-schemas, --dir et --force](img/mcp.png)

## L’exécuter sans assistant

Pour un client sur une autre machine, servez-le plutôt en HTTP :

```sh
genoffice mcp --http 8765 --token <secret>
```

`--host <addr>` change le point d’écoute. Passez le même jeton au client.

En HTTP, les fichiers circulent aussi : `PUT /files/<name>` en envoie un, chaque outil accepte une URL `http(s)` à la place d’un chemin, et les sorties reviennent sous forme d’URL de téléchargement — et, si elles sont assez petites, de ressources intégrées. Les ops, les specs et le Markdown sont transmis en ligne dans les deux cas.

## Ce que l’agent obtient

Chaque commande est un outil. Les plus intéressantes :

- **`docs`, `sheet`, `slides`** — lire et modifier un fichier par le chemin d’écriture de l’application, un **op** à la fois. Un nouveau diaporama suit `deck_start`, `deck_page`, `deck_build`.
- **`render`** — un PNG par page, mis en page par le moteur de rendu de l’application, pour que l’agent regarde une diapositive au lieu de la deviner.
- **`pdf`** — la couche de texte d’un PDF, page par page, sans lancer l’application.
- **`info`** — métadonnées et résumé de la structure : c’est généralement l’appel le moins cher sur un fichier inconnu.
- **`search`, `image`, `media`** — les fournisseurs configurés dans l’application, donc l’agent n’a pas besoin de ses propres clés.
- **`merge`** — remplir un modèle `{{key}}`.

## Schémas et budget plus serré

`apply` et `create` annoncent leurs paramètres `ops`, `cells` et `data` avec le schéma typé de chaque op, généré à partir de `genoffice guide <domain> --json`. C’est précis, et ce n’est pas petit. Un client dont la fenêtre de contexte est serrée peut demander de simples tableaux à la place :

```sh
genoffice mcp --compact-schemas
GENOFFICE_MCP_COMPACT_SCHEMAS=1 genoffice mcp
```

## Pourquoi la compétence

Un agent qui ne connaît pas le vocabulaire des ops devine. La compétence embarque la référence et les guides de conception — le même contenu que celui qu’imprime `genoffice guide` —, si bien que l’assistant écrit des ops dont il a réellement lu la spécification. Installez-la depuis le panneau ci-dessus, ou avec `genoffice skill` depuis un terminal.

## Ce qu’il atteint dans l’application

Le serveur ne se limite pas aux fichiers sur le disque. Tant que GenOffice tourne, l’agent peut aussi travailler par la fenêtre :

- **`open_in_genoffice`** ouvre un fichier dans un onglet et le met au premier plan.
- **`open_documents`** liste tous les documents que vous avez ouverts — id, type, chemin, et s’ils ont des modifications non enregistrées —, puis lit le contenu actuel de l’un d’eux ou le ferme, en enregistrant d’abord, sauf si vous lui demandez de jeter.
- **Les outils de contenu** prennent cet id (ou le chemin) comme argument `document`, si bien qu’une modification atterrit dans l’onglet que vous avez déjà ouvert et que la fenêtre bascule pour l’afficher.

Deux choses restent hors de portée : il n’y a pas de panneau IA, et la boîte de dialogue de mise à jour de l’application ne s’applique pas.

## Le panneau Serveur HTTP local

Sous **Serveur HTTP local**, l'application fait tourner le serveur elle-même au lieu de le laisser à l'assistant : un interrupteur d'activation, un champ **Port**, un indicateur **En cours / À l'arrêt** et la **Génération en arrière-plan** (écrire les documents directement dans un chemin sans ouvrir l'interface) et l'**Exemple de configuration client** à copier. Ouvrir **Avancé** ajoute les deux URL de connexion — Streamable HTTP et celle, plus ancienne, de SSE —, une URL de **Test de santé** et un interrupteur de **Journalisation** qui enregistre l'activité du serveur et des outils dans un fichier local que vous pouvez **Ouvrir**, **Actualiser** ou **Effacer** depuis cet endroit. Il n'écoute que sur localhost.
