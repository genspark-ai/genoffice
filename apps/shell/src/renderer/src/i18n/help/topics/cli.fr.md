# Ligne de commande et agents

Chaque installation livre une commande `genoffice` qui pilote les mêmes moteurs que la fenêtre : les mêmes analyseurs, le même rédacteur, le même moteur de rendu. Un fichier enregistré par l’application et un fichier écrit par la commande sont le même fichier, et un contrôle que le panneau IA de l’application valide, la commande le valide aussi.

```sh
genoffice --help          # every command
genoffice guide slides    # the op reference, for writing your own
```

Voilà toute la surface sur un seul écran : chaque commande avec une ligne qui dit ce qu'elle fait, puis les options globales et les codes de sortie :

![La vraie sortie de genoffice --help : la bannière de version, chaque commande avec une description d'une ligne, ainsi que les options globales et les codes de sortie](img/cli.png)

## Obtenir la commande

macOS et Windows la livrent dans le bundle de l’application. Pour l’appeler par son nom, exécutez `genoffice install-cli` une fois : il crée un lien symbolique vers le binaire fourni dans `/usr/local/bin`, ou dans le `PATH` utilisateur sous Windows.

## Les commandes à connaître

| Commande          | Ce qu’elle fait                                                                                                                                                                                                                                  |
| ----------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| `open`            | Ouvrir un document dans l’application ; démarre l’application si elle ne tourne pas.                                                                                                                                                             |
| `convert`         | Convertir entre formats avec les moteurs de l’application.                                                                                                                                                                                       |
| `create`          | Créer un document à partir de contenu structuré.                                                                                                                                                                                                 |
| `render`          | Un PNG par page, tel que le moteur de rendu le met en page.                                                                                                                                                                                      |
| `pdf`             | Lire la couche de texte d’un PDF page par page, sans lancer l’application.                                                                                                                                                                       |
| `info`            | Métadonnées et résumé de la structure d’un document.                                                                                                                                                                                             |
| `search`          | Recherche web ou d’images via le fournisseur configuré dans l’application.                                                                                                                                                                       |
| `image` / `media` | Générer une image, ou décrire un fichier image, vidéo ou audio et lui poser des questions.                                                                                                                                                       |
| `merge`           | Remplir les espaces réservés `{{key}}` d’un modèle `.docx`, `.pptx` ou `.xlsx`.                                                                                                                                                                  |
| `capabilities`    | Indiquer quelles fonctions cloud sont configurées sur cette machine.                                                                                                                                                                             |
| `guide`           | La référence des ops et les guides de conception, générés à partir des définitions mêmes contre lesquelles l’executor valide : ils ne peuvent donc pas diverger de ce que `apply` accepte. `--json` renvoie le tout avec le schéma de chaque op. |
| `install-cli`     | Placer `genoffice` sur le `PATH`.                                                                                                                                                                                                                |
| `skill`           | Lister les agents de programmation trouvés sur cette machine et y installer ou mettre à jour la compétence GenOffice.                                                                                                                            |
| `mcp`             | Exposer chaque commande comme un outil Model Context Protocol. Voir **Connecter un agent de codage**.                                                                                                                                            |

## Édition : documents, tableurs, présentations

`genoffice docs`, `genoffice sheet` et `genoffice slides` lisent et modifient par le même chemin d’écriture que l’application, et partagent un même vocabulaire : un **op** est une seule modification, et une **spec** est une liste d’ops appliquées dans l’ordre.

```sh
genoffice sheet read report.xlsx
genoffice slides apply deck.pptx --ops ops.json --dry-run
genoffice docs apply brief.docx --ops ops.json
```

`--dry-run` rapporte ce qu’un lot ferait et n’écrit rien : c’est le moyen bon marché de vérifier une spec avant qu’elle ne s’applique. Le panneau IA de l’application s’exécute exactement sur ces ops, donc tout ce que vous pouvez lui demander se script.

## Model Context Protocol

`genoffice mcp` expose chaque commande comme un outil MCP, et `genoffice mcp install <agent|all>` l’enregistre dans la configuration d’un agent de programmation. Pour ce côté-là, voir **Connecter un agent de codage**.

## Ce que la commande ne fait pas

Elle lit et écrit le fichier. Ce n’est pas l’application : il n’y a pas de fenêtre, et la boîte de dialogue de mise à jour de l’application ne s’applique pas. Tout ce qui a besoin de la fenêtre — le panneau IA, le contrôle QC d’une diapositive rendue — attend que vous ouvriez le fichier. `genoffice render` vous donne les pixels sans elle, et `genoffice slides` audite la mise en page d’un diaporama tout seul.
