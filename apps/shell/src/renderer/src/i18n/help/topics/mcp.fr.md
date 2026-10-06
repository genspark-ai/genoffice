# Connecter un agent de codage

GenOffice parle le Model Context Protocol : un agent de programmation peut donc lire, écrire et rendre vos documents avec les mêmes moteurs que l’application. L’agent ne devine pas un format de fichier : il reçoit les schémas typés des ops à partir des définitions mêmes contre lesquelles l’executor valide.

## L’enregistrer

Le cas courant tient en une commande :

```sh
genoffice mcp install all
```

Il trouve les agents de programmation présents sur cette machine — Claude Code, Codex, Cursor, Gemini CLI, Copilot CLI, OpenCode, Windsurf — et écrit l’entrée du serveur stdio dans la configuration de chacun, en laissant le reste du fichier tel qu’il l’a trouvé.

```sh
genoffice mcp list             # where each agent stands
genoffice mcp install cursor    # just one
genoffice mcp uninstall cursor  # take it back out
```

Un agent installé ailleurs demande `--dir <path>` ; `--force` réécrit une entrée déjà présente.

## L’exécuter vous-même

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

## La compétence

Un agent qui ne connaît pas le vocabulaire des ops devine. `genoffice skill` installe une compétence GenOffice dans les agents qu’il trouve, en emportant la référence et les guides de conception — le même contenu que celui qu’imprime `genoffice guide`.

## Ce que ce n’est pas

Le serveur MCP lit et écrit des fichiers. Ce n’est pas la fenêtre : il n’y a pas de panneau IA, et la boîte de dialogue de mise à jour de l’application ne s’applique pas. Si une étape a besoin de la fenêtre, ouvrez le fichier.
