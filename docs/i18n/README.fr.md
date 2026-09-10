<p align="center">
  <a href="https://genoffice.ai/">
    <picture>
      <source srcset="../assets/readme/hero-dark.webp" media="(prefers-color-scheme: dark)">
      <img src="../assets/readme/hero.webp" alt="GenOffice — la suite bureautique IA open source : Docs, Sheets, Slides, PDF, Markdown et HTML avec un panneau IA intégré" width="100%">
    </picture>
  </a>
</p>

<h1 align="center">GenOffice</h1>

<p align="center"><b>La première suite bureautique IA open source complète au monde.</b><br>
Fichiers Word, Excel, PowerPoint et PDF, modifiés par vous et votre IA, enregistrés dans les formats natifs.</p>

<p align="center">
  <a href="../../LICENSE"><img src="https://img.shields.io/github/license/genspark-ai/genoffice" alt="License: Apache-2.0"></a>
  <a href="https://github.com/genspark-ai/genoffice/releases/latest"><img src="https://img.shields.io/github/v/release/genspark-ai/genoffice" alt="Latest release"></a>
  <a href="https://github.com/genspark-ai/genoffice/releases"><img src="https://img.shields.io/github/downloads/genspark-ai/genoffice/total" alt="Downloads"></a>
  <a href="https://github.com/genspark-ai/genoffice/stargazers"><img src="https://img.shields.io/github/stars/genspark-ai/genoffice?style=flat" alt="GitHub stars"></a>
</p>

<p align="center"><a href="../../README.md">English</a> · <a href="README.es.md">Español</a> · <a href="README.pt-BR.md">Português (Brasil)</a> · <a href="README.de.md">Deutsch</a> · <b>Français</b> · <a href="README.zh-CN.md">简体中文</a> · <a href="README.zh-TW.md">繁體中文</a> · <a href="README.ko.md">한국어</a> · <a href="README.ja.md">日本語</a> · <a href="README.ar.md">العربية</a> · <a href="README.ru.md">Русский</a> · <a href="README.it.md">Italiano</a> · <a href="README.nl.md">Nederlands</a> · <a href="README.pl.md">Polski</a> · <a href="README.cs.md">Čeština</a> · <a href="README.id.md">Bahasa Indonesia</a> · <a href="README.ms.md">Bahasa Melayu</a> · <a href="README.th.md">ไทย</a> · <a href="README.hi.md">हिन्दी</a> · <a href="README.he.md">עברית</a></p>

<p align="center">
  <a href="#download"><b>Télécharger</b></a> ·
  <a href="https://genoffice.ai/"><b>Site web</b></a> ·
  <a href="https://genoffice.ai/join"><b>Communauté</b></a> ·
  <a href="../../PRIVACY.md"><b>Confidentialité</b></a>
</p>

GenOffice est une alternative gratuite et open source à Microsoft Office pour
macOS, Windows et Linux. Elle ouvre et enregistre des fichiers natifs
`.docx`, `.xlsx` et `.pptx`, permet d'éditer des PDF, du Markdown et du HTML,
et place un agent IA au cœur de chaque document — pas une simple fenêtre de
chat greffée sur le côté, mais un éditeur qui lit le fichier, effectue la
modification, et vous montre exactement ce qu'il a touché.

- **Formats natifs, préservation au bit près.** Seul ce que vous modifiez est
  réécrit. Le reste du fichier est conservé à l'identique, si bien que vos
  documents continuent de fonctionner dans Word, Excel et PowerPoint.
- **Une IA que vous pouvez vérifier.** Les modifications arrivent sous forme
  de suivi des modifications et de diffs, avec une annulation en un clic. Les
  feuilles de calcul reçoivent de vraies formules, pas des chiffres collés.
  Les diapositives et les pages sont générées directement sur le canevas et
  restent entièrement modifiables.
- **Local par conception.** L'ouverture, l'édition, l'enregistrement et la
  conversion des fichiers se font sur votre machine. Les conversions
  PDF → Word / Excel / PowerPoint, Markdown → Word et HTML → Word s'exécutent
  toutes en local. Seuls les appels à l'IA quittent la machine, vers le
  fournisseur de votre choix.
- **Vos clés ou aucune.** Connectez-vous avec Genspark pour vous dispenser de
  clé, ou utilisez votre propre clé pour Claude, OpenAI, Gemini, DeepSeek,
  Kimi, GLM, Qwen, Doubao, MiniMax, Grok, Mistral, OpenRouter, ou tout point
  de terminaison compatible OpenAI, y compris les serveurs locaux.

**À télécharger :** [macOS](https://github.com/genspark-ai/genoffice/releases/latest) (Apple Silicon et Intel) ·
[Windows](https://github.com/genspark-ai/genoffice/releases/latest) (x64 et Arm) ·
[Linux](https://github.com/genspark-ai/genoffice/releases/latest) (deb, rpm, AppImage) —
détails et prérequis dans la section [Télécharger](#download).

## Démonstration

Six applications, un seul panneau IA. Chaque capture d'écran montre
l'application réelle sur macOS, avec l'IA pilotée par le prompt que vous
pouvez lire dans le panneau.

### 1 · Docs — ouvrez et modifiez des `.docx` avec une IA que vous pouvez vérifier

<table>
<tr>
<td width="50%"><img src="../assets/readme/docs-report.webp" alt="GenOffice Docs affichant une page de rapport annuel à deux colonnes avec une image de couverture pleine largeur, un tableau de KPI ombré, un en-tête et un pied de page, à un zoom de 80 % avec le panneau IA réduit"></td>
<td width="50%"><img src="../assets/readme/docs-ai.webp" alt="GenOffice Docs : une présentation d'entreprise avec une image de bannière ; l'IA a resserré la section Overview et inséré une nouvelle section à puces, et le panneau propose une annulation en un clic"></td>
</tr>
<tr>
<td><b>Ouvre le fichier comme Word le met en page</b> — sections à deux colonnes, images pleine page, tableaux ombrés, en-têtes et pieds de page, pagination fondée sur les métriques de ligne de Word. Styles, commentaires, suivi des modifications, équations et annotations manuscrites font l'aller-retour sans être altérés.</td>
<td><b>Demandez la modification</b> — l'IA lit les blocs dont elle a besoin, réécrit la section Overview et insère une nouvelle section à puces. Chaque intervention de l'IA est un instantané que vous pouvez annuler ; avec le <b>suivi des modifications</b> activé, les changements arrivent sous forme de révisions au style de Word.</td>
</tr>
</table>

### 2 · Sheets — des `.xlsx` avec des formules et des graphiques vivants, pas des chiffres collés

<table>
<tr>
<td width="50%"><img src="../assets/readme/sheets-ai.webp" alt="GenOffice Sheets : l'IA a ajouté une feuille Summary avec le chiffre d'affaires par région et par catégorie à l'aide de formules SUMIF, ainsi qu'un graphique en colonnes, et indique 43 modifications appliquées avec un bouton Annuler"></td>
<td width="50%"><img src="../assets/readme/sheets-qa.webp" alt="GenOffice Sheets : interrogée sur la région ayant généré le plus de chiffre d'affaires au T2, l'IA répond Europe avec la répartition par catégorie et cite les cellules utilisées sous forme de liens, à côté de la feuille Orders"></td>
</tr>
<tr>
<td><b>Construisez-la</b> — à partir d'une seule phrase, l'agent ajoute une feuille Summary avec de vraies formules <code>SUMIF</code> par région et par catégorie, insère un graphique en colonnes, et applique les 43 modifications sous forme d'un seul lot annulable.</td>
<td><b>Interrogez-la</b> — les questions sur le classeur reviennent avec le raisonnement et les cellules exactes utilisées, présentées comme des citations cliquables. En arrière-plan : un moteur `.xlsx` Rust maison, des tableaux croisés dynamiques, des segments, une mise en forme conditionnelle et le traçage des formules.</td>
</tr>
</table>

### 3 · Slides — d'un prompt à un jeu de diapositives `.pptx`

<img src="../assets/readme/slides-generate.webp" alt="Time-lapse de GenOffice Slides générant le deck investisseurs Aurora Home : l'IA planifie la trame narrative dans le panneau, les diapositives apparaissent une à une sur le canevas, et le deck terminé se conclut sur la demande finale" width="100%">

<table>
<tr>
<td width="50%"><img src="../assets/readme/slides-cover.webp" alt="GenOffice Slides : la diapositive de couverture d'un jeu de diapositives Aurora Home pour investisseurs généré par l'IA sur le canevas, avec le prompt original d'une ligne et le résumé de l'IA de ce qu'elle a construit dans le panneau"></td>
<td width="50%"><img src="../assets/readme/slides-ai.webp" alt="GenOffice Slides : la diapositive de clôture, designée, du même jeu de 11 diapositives, avec le bandeau de miniatures à gauche et le panneau IA résumant la trame narrative"></td>
</tr>
<tr>
<td><b>Une ligne en entrée</b> — « Créer un jeu de 10 diapositives pour un pitch investisseur pour Aurora Home… ». GenOffice planifie la trame narrative, recherche les chiffres, et rédige chaque diapositive directement sur le canevas sous forme d'un véritable `.pptx`.</td>
<td><b>Un jeu de diapositives fini en sortie</b> — onze diapositives designées avec une typographie et une iconographie cohérentes, et un appel à l'action final ; continuez à modifier avec les masques, les mises en page, les guides intelligents et le recadrage non destructif, ou demandez au panneau de restyliser, réécrire et réorganiser.</td>
</tr>
</table>

### 4 · PDF — modifiez le texte d'un PDF sur place, convertissez un PDF en Word en local

<table>
<tr>
<td width="50%"><img src="../assets/readme/pdf-edit.webp" alt="GenOffice PDF : le mode Modifier le texte délimite chaque bloc de texte de la page pour une modification sur place tandis que le panneau IA répond à une question sur le rapport avec des citations de page"></td>
<td width="50%"><img src="../assets/readme/pdf-convert.webp" alt="GenOffice Docs affichant un document Word converti localement à partir du PDF de bilan trimestriel Helios, ouvert dans un second onglet à côté du PDF d'origine"></td>
</tr>
<tr>
<td><b>Modifiez à l'intérieur de la page</b> — le mode Modifier le texte délimite chaque bloc de texte pour une nouvelle saisie sur place ; le flux de contenu est réécrit via PDFium avec les polices d'origine, et non par une annotation de recouvrement. Interrogez l'IA sur un long rapport et obtenez des réponses avec citations de page.</td>
<td><b>Convertissez en local</b> — <b>Convertisseur PDF → PDF vers Word</b> produit un `.docx` modifiable qui s'ouvre dans Docs à côté de la source, avec titres, lignes de statistiques et paragraphes intacts. Les cibles Excel et PowerPoint fonctionnent de la même façon ; les pages scannées passent par l'OCR du système.</td>
</tr>
</table>

### 5 · HTML — un générateur IA de pages et d'interfaces, avec un brief de design d'abord

Indiquez à quoi sert la page et à qui elle s'adresse. L'IA propose d'abord un
**brief de design** — accroche, palette, typographie et orientations
stylistiques — puis construit un fichier `.html` autonome à partir de ces
jetons.

<img src="../assets/readme/html-restyle-motion.webp" alt="Time-lapse de GenOffice HTML relookant la page d'accueil Lumen : une seule requête Restyle dans le panneau transforme la page sombre Midnight Studio en version chaleureuse Solar Daybreak, tandis que chaque section et tout le contenu restent en place" width="100%">

<table>
<tr>
<td width="50%"><img src="../assets/readme/html-ai.webp" alt="GenOffice HTML : une page de destination générée pour une lampe de bureau solaire dans l'orientation sombre Midnight Studio, affichée dans l'aperçu en direct avec le panneau IA résumant la page qu'elle vient de construire"></td>
<td width="50%"><img src="../assets/readme/html-restyle.webp" alt="La même page de destination Lumen restylisée par l'IA dans l'orientation chaude Solar Daybreak : fond papier, titres en serif et une couleur d'accent orange, avec toutes les sections et tout le texte conservés"></td>
</tr>
<tr>
<td><b>Générée à partir d'un seul prompt</b> — une accroche audacieuse, des cartes de fonctionnalités, une grille tarifaire et un formulaire de liste d'attente pour Lumen, construits dans l'orientation Midnight Studio. Cliquez sur n'importe quel élément pour le restyliser, double-cliquez pour modifier le texte, ou passez à la vue source CodeMirror.</td>
<td><b>Même design, nouvelle orientation</b> — une seule demande de <b>restylisation</b> remplace les jetons du brief et la page suit : papier chaud, serif éditorial, accent orange soleil, rien n'est réécrit. Présentez en plein écran, ou exportez en PDF ou en document Word natif et modifiable.</td>
</tr>
</table>
<table>
<tr>
<td width="50%"><img src="../assets/readme/html-dashboard.webp" alt="GenOffice HTML : un tableau de bord personnel généré pour un designer freelance dans un style lin chaleureux, avec un rail latéral gauche, une salutation en serif et quatre cartes de métriques"></td>
<td width="50%"><img src="../assets/readme/html-report.webp" alt="GenOffice HTML : un rapport de données sur le marché des véhicules électriques généré dans un style broadsheet, avec un bandeau de titre en serif, un chiffre principal de 17,3 millions et une ligne de statistiques"></td>
</tr>
<tr>
<td><b>Maquettes d'interface</b> — le modèle de démarrage « tableau de bord personnel » transforme un persona en une mise en page fonctionnelle : rail latéral, salutation, sparkline d'heures facturables, cartes de factures et de taux d'occupation, tout en vrai HTML que vous pouvez confier à un développeur.</td>
<td><b>Récits de données</b> — le modèle de démarrage « rapport de données » construit un broadsheet éditorial : bandeau de titre en serif, un chiffre principal, une ligne de statistiques séparée par une règle, des graphiques SVG intégrés et une note méthodologique.</td>
</tr>
</table>

### 6 · Markdown — un éditeur par blocs sur du `.md` brut, avec Ask AI

<table>
<tr>
<td width="50%"><img src="../assets/readme/markdown-ai.webp" alt="GenOffice Markdown : un paragraphe sélectionné affiche une bulle Ask AI avec une instruction saisie et des puces de suggestion telles que Polish, Make more concise, Expand et Fix grammar, ainsi que des boutons Send now et Add to queue"></td>
<td width="50%"><img src="../assets/readme/markdown-render.webp" alt="GenOffice Markdown affichant un document de notes de lancement avec un tableau, un diagramme de flux Mermaid et une liste de tâches, avec les prompts de démarrage du panneau IA à gauche"></td>
</tr>
<tr>
<td><b>Demandez à l'IA à propos d'une sélection</b> — sélectionnez n'importe quel passage et une puce <b>Ask AI</b> apparaît : saisissez une instruction ou choisissez une suggestion, envoyez-la immédiatement, ou mettez en file plusieurs modifications ancrées et exécutez-les en une seule passe. La même entrée existe dans chaque application.</td>
<td><b>Rendu, enregistré en Markdown brut</b> — titres, listes, tableaux, images, blocs de code et diagrammes Mermaid dans un éditeur par blocs Tiptap, réécrits en `.md` brut, avec un export <b>Markdown → Word</b> entièrement local.</td>
</tr>
</table>

## Pourquoi GenOffice

- **Open source**, licence Apache-2.0, développé au grand jour sur GitHub.
- **À vous de l'exécuter.** Applications natives pour macOS, Windows et
  Linux ; les fichiers restent sur votre disque et chaque modification,
  enregistrement et conversion s'effectue sur votre machine.
- **De vrais fichiers Office.** `.docx`, `.xlsx` et `.pptx` natifs, préservés
  au bit près : les parties du fichier non modifiées sont copiées telles
  quelles.
- **Une IA qui modifie le document lui-même.** Suivi des modifications dans
  Docs, formules et graphiques vivants dans Sheets, diapositives dessinées
  directement sur le canevas, chaque étape de l'IA laissant un instantané que
  vous pouvez restaurer.
- **Votre modèle, votre clé.** Connectez-vous avec Genspark, ou apportez une
  clé pour Claude, OpenAI, Gemini, DeepSeek et d'autres, serveurs locaux et
  tout point de terminaison compatible OpenAI inclus.
- **Le PDF fait sérieusement.** Modifiez le texte directement dans la page,
  et convertissez le PDF en Word, Excel ou PowerPoint en local, avec OCR
  système pour les scans.
- **Markdown et HTML aussi**, avec le même panneau IA et un export local vers
  Word.
- **Gratuit**, pour les particuliers comme pour les équipes.

## Moteurs IA

**Connectez-vous avec Genspark** et il n'y a rien à configurer : les appels
aux modèles passent par le proxy Genspark (familles Claude, GPT et Gemini) et
les agents ont accès à la recherche web et d'images, à la génération
d'images, et à l'analyse d'images, d'audio et de vidéo.

**Ou utilisez votre propre clé.** Paramètres → IA propose Claude, OpenAI,
Gemini, DeepSeek, Kimi, GLM, Qwen, Doubao, MiniMax, Grok, Mistral, OpenRouter
et OpenCode Zen/Go, ainsi qu'un emplacement personnalisé pour tout point de
terminaison compatible OpenAI (URL de base + clé), y compris les serveurs de
modèles locaux. La recherche et les médias disposent de leurs propres
fournisseurs par fonctionnalité sous **IA Médias et Recherche** : Serper ou
Tavily pour la recherche web, et OpenAI, Gemini, Doubao/Seedream, GLM, Grok,
Qwen, MiniMax ou tout point de terminaison d'images compatible OpenAI pour la
génération d'images et l'analyse d'images/vidéos.

Toute la suite propose les thèmes clair, sombre et système. Les thèmes ne
changent que ce qui est affiché à l'écran : les exports, les impressions et
les fichiers enregistrés conservent toujours les couleurs propres au
document.

<a id="download"></a>

## Télécharger

| Plateforme                           | Prérequis                                             | Téléchargement                                                                                |
| ------------------------------------ | ----------------------------------------------------- | --------------------------------------------------------------------------------------------- |
| **macOS** — Apple Silicon (arm64)    | macOS 11+                                             | [Dernier `.dmg` (arm64)](https://github.com/genspark-ai/genoffice/releases/latest)            |
| **macOS** — Intel (x64)              | macOS 11+                                             | [Dernier `.dmg` (x64)](https://github.com/genspark-ai/genoffice/releases/latest)              |
| **Windows** (x64, la plupart des PC) | Windows 10+, Intel/AMD                                | [Dernier installateur `-x64.exe`](https://github.com/genspark-ai/genoffice/releases/latest)   |
| **Windows** sur Arm (ARM64)          | Windows 11 sur Arm (Snapdragon X et similaires)       | [Dernier installateur `-arm64.exe`](https://github.com/genspark-ai/genoffice/releases/latest) |
| **Linux** — Debian / Ubuntu          | x86_64, glibc 2.34+ (Ubuntu 22.04 ou plus récent)     | [Dernier `.deb`](https://github.com/genspark-ai/genoffice/releases/latest)                    |
| **Linux** — Fedora / RHEL / openSUSE | x86_64, glibc 2.34+ (Fedora 35+, RHEL 9+, Leap 15.6+) | [Dernier `.rpm`](https://github.com/genspark-ai/genoffice/releases/latest)                    |
| **Linux** — autres distributions     | x86_64, glibc 2.34+, FUSE 2                           | [Dernier `.AppImage`](https://github.com/genspark-ai/genoffice/releases/latest)               |

Toutes les versions proviennent de `main` ; les installateurs macOS et
Windows sont signés. Les versions plus anciennes se trouvent sur la page des
[Releases](https://github.com/genspark-ai/genoffice/releases).

<details>
<summary><b>Installation sur Linux</b></summary>

Le paquet deb s'installe avec apt — il récupère les dépendances et ajoute
GenOffice au menu des applications :

```bash
sudo apt install ./genoffice_<version>_amd64.deb
```

Sur Fedora / la famille RHEL / openSUSE, installez plutôt le paquet rpm :

```bash
sudo dnf install ./genoffice-<version>.x86_64.rpm     # Fedora / famille RHEL
sudo zypper install ./genoffice-<version>.x86_64.rpm  # openSUSE
```

L'AppImage s'exécute sur place : installez le runtime FUSE 2
(`sudo apt install libfuse2` ; sur Ubuntu 24.04 le paquet se nomme
`libfuse2t64`), rendez le fichier exécutable, puis lancez-le :

```bash
chmod +x GenOffice-<version>.AppImage
./GenOffice-<version>.AppImage
```

</details>

## Fonctionnement

Sept applications Electron — Docs, Sheets, Slides, PDF, Markdown, HTML et la
coquille à onglets — partagent une même couche moteur composée de paquets
TypeScript purs, ainsi qu'un sidecar Rust pour le `.xlsx`. Le fichier
d'origine est toujours la source de vérité : les modifications sont
appliquées sous forme de correctifs ciblés, et tout ce que l'éditeur n'a pas
touché survit à l'aller-retour sans altération.

```
open docx ─► archive l'original par hachage (jamais touché)
          ─► analyse word/document.xml en un arbre de blocs, chaque bloc ancré à son XML d'origine
          ─► éditeur Tiptap (édition manuelle + IA, suivi des modifications non enregistrées)
save      ─► blocs modifiés → fragments OOXML (référençant uniquement les styles existants)
          ─► insertion dans le document.xml d'origine ; les blocs non touchés conservent leurs octets
          ─► reconstruction du zip ; chaque autre entrée est copiée octet pour octet
```

La visite guidée paquet par paquet (moteurs docx/pptx, `pdf2docx`,
`html2docx`, le cœur de l'agent et les fournisseurs) se trouve dans
[CONTRIBUTING.md](../../CONTRIBUTING.md#engine-packages).

## Développement

```bash
npm install
npm run fixtures     # génère les fixtures .docx de test
npm test             # tests unitaires moteur + applications (docs/sheets/slides n'ont pas besoin d'affichage)
npm run typecheck    # tsc --noEmit sur tout le workspace
npm run dev          # les six éditeurs + la coquille avec les serveurs de dev Vite
npm run dev:docs     # une seule application (le même schéma fonctionne par workspace)
npm run dist:mac     # empaquette le dmg macOS (régénère les mentions tierces)
npm run dist:win     # empaquette l'installateur nsis Windows
npm run dist:linux   # empaquette l'AppImage + deb + rpm Linux
```

L'application sheets a en outre besoin d'une chaîne d'outils Rust pour son
sidecar xlsx (`cargo` dans le PATH) ; `npm run build -w @genoffice/sheets` le
compile automatiquement. Consultez
[CONTRIBUTING.md](../../CONTRIBUTING.md) pour les vérifications que chaque
modification doit passer et la façon dont les pull requests sont intégrées.

## Communauté

GenOffice est en développement actif et vos retours contribuent à le
façonner.

- **Signalez un bug ou proposez une fonctionnalité** sur
  [GitHub Issues](https://github.com/genspark-ai/genoffice/issues).
- **Rejoignez le salon de discussion GenOffice** sur
  [GenTeam](https://genoffice.ai/join) pour échanger avec l'équipe et les
  autres utilisateurs.
- **Mettez une étoile au dépôt** si GenOffice vous est utile — c'est la
  meilleure façon de soutenir le projet.

## FAQ

<details>
<summary><b>GenOffice est-il gratuit ?</b></summary>

Oui. GenOffice est gratuit et open source sous licence Apache-2.0 — pas
d'essai, pas de palier payant pour les applications elles-mêmes.

</details>

<details>
<summary><b>GenOffice peut-il ouvrir des fichiers Microsoft Word, Excel et PowerPoint ?</b></summary>

Oui. GenOffice ouvre et enregistre des fichiers natifs `.docx`, `.xlsx` et
`.pptx`. L'enregistrement préserve les octets : les parties du fichier que
vous n'avez pas modifiées sont réécrites telles quelles, si bien que vos
documents continuent de fonctionner dans Microsoft Office.

</details>

<details>
<summary><b>GenOffice fonctionne-t-il hors ligne ?</b></summary>

L'édition de documents est entièrement locale — les fichiers ne quittent
jamais votre machine pour être ouverts, modifiés, enregistrés ou convertis.
Les fonctionnalités IA (agents, recherche, outils d'image) nécessitent une
connexion réseau, avec soit une connexion Genspark, soit votre propre clé
d'API de modèle.

</details>

<details>
<summary><b>GenOffice peut-il modifier des fichiers PDF ?</b></summary>

Oui — une véritable édition du texte et des images du PDF, qui réécrit le
flux de contenu de la page en préservant les polices d'origine, et non de
simples annotations de recouvrement.

</details>

<details>
<summary><b>GenOffice peut-il convertir un PDF en Word, Excel ou PowerPoint ?</b></summary>

Oui — entièrement en local : extraction au niveau des caractères via PDFium
combinée à une analyse de mise en page fondée sur la géométrie, sans service
cloud, sans envoi. Les pages scannées sont également prises en charge : sur
macOS et Windows, l'OCR du système les lit, si bien qu'elles sont converties
en texte modifiable plutôt qu'en image de page.

</details>

<details>
<summary><b>Puis-je utiliser mon propre modèle IA ou ma propre clé d'API ?</b></summary>

Oui. Outre la connexion Genspark sans clé, GenOffice prend en charge
l'utilisation de votre propre clé pour Claude, OpenAI, Gemini, DeepSeek,
Kimi, GLM, Qwen, Doubao, MiniMax, Grok, Mistral, OpenRouter et OpenCode
Zen/Go, ainsi que tout point de terminaison compatible OpenAI — y compris les
serveurs de modèles locaux. La recherche, la génération d'images et
l'analyse d'images/vidéos utilisent leurs propres clés sous Paramètres → IA
Médias et Recherche.

</details>

<details>
<summary><b>GenOffice peut-il convertir du HTML en Word ?</b></summary>

Oui — Exporter en Word dans l'application HTML produit un `.docx` natif et
modifiable, entièrement en local. La page est rendue dans le moteur Chromium
intégré puis réduite à de véritables structures Word : titres, paragraphes,
listes, tableaux, cartes, lignes de KPI, champs de formulaire et arrière-plans
de page ; seuls les éléments visuels sans équivalent Word (graphiques,
icônes, encadrés décorés) sont intégrés sous forme d'images.

</details>

<details>
<summary><b>GenOffice collecte-t-il des données ?</b></summary>

Les versions officielles empaquetées envoient par défaut des analyses
d'usage limitées, et vous pouvez désactiver ce signalement à tout moment sous
Paramètres → Général. Les analyses n'envoient jamais le contenu des
documents, les noms de fichiers, les chemins de fichiers, l'identité du
compte ou les adresses e-mail. Consultez
[GenOffice Privacy](../../PRIVACY.md) pour la liste complète des événements
et des données divulguées.

</details>

## Sécurité

Consultez [SECURITY.md](../../SECURITY.md) pour la posture de sécurité des
processus (isolation du renderer, validation IPC, contrôle des liens
externes) et les modèles de menace concernant le contenu généré par l'IA.

## Remerciements

GenOffice ne serait pas possible sans ces projets open source :

- [Electron](https://www.electronjs.org/) — le runtime de bureau pour chaque
  application.
- [Univer](https://github.com/dream-num/univer) (Apache-2.0) — le cœur
  d'interface de feuille de calcul que Sheets étend.
- [PDFium](https://pdfium.googlesource.com/pdfium/) (BSD-3-Clause, intégré
  via [@embedpdf/pdfium](https://github.com/embedpdf/embed-pdf-viewer)) — le
  moteur de flux de contenu derrière la véritable édition de texte et
  d'images des PDF.
- [pdf.js](https://github.com/mozilla/pdf.js) (Apache-2.0) et
  [pdf-lib](https://github.com/Hopding/pdf-lib) (MIT) — le rendu PDF et
  l'assemblage de documents.
- [Tiptap](https://tiptap.dev/) / [ProseMirror](https://prosemirror.net/) —
  les éditeurs par blocs de Docs et Markdown.
- [CodeMirror](https://codemirror.net/) (MIT) — l'éditeur de source dans
  HTML.
- [Konva](https://konvajs.org/) — le rendu canevas pour Slides et les
  graphiques de Sheets.
- [HarfBuzz](https://github.com/harfbuzz/harfbuzz) (wasm) — les métriques de
  mise en forme de texte pour les écritures complexes.
- [calamine](https://github.com/tafia/calamine) et
  [IronCalc](https://github.com/ironcalc/IronCalc) — les couches de lecture
  et de calcul du sidecar xlsx en Rust.
- [libeot](https://github.com/umanwizard/libeot) (MPL-2.0) — le décodeur
  MicroType Express pour les polices PowerPoint intégrées, porté en
  TypeScript.
- [React](https://react.dev/) (MIT) — la couche d'interface utilisateur de
  chaque application.
- [Mermaid](https://mermaid.js.org/) (MIT) et [KaTeX](https://katex.org/)
  (MIT) — les diagrammes et les formules mathématiques dans Markdown et
  Docs.
- [opentype.js](https://opentype.js.org/) (MIT) — l'analyse des polices pour
  les métriques et la recherche de glyphes.
- [JSZip](https://stuk.github.io/jszip/) (MIT) et
  [fast-xml-parser](https://github.com/NaturalIntelligence/fast-xml-parser)
  (MIT) — les couches de conteneur OOXML et de XML.
- [Fluent UI System Icons](https://github.com/microsoft/fluentui-system-icons)
  (MIT) — le jeu d'icônes des rubans.
- [electron-updater](https://www.electron.build/) (MIT) — les mises à jour
  intégrées à l'application.
- Les polices Liberation, Carlito, Caladea et Noto CJK (OFL/Apache-2.0) — les
  polices de document intégrées.

`npm run notices` régénère le résumé des licences tierces intégrées
(`tools/gen-third-party-notices.mjs`) ; toutes les dépendances d'exécution
sont sous licence MIT/Apache-2.0/BSD-3-Clause/OFL.

## Licence

GenOffice est distribué sous [licence Apache 2.0](../../LICENSE), avec une
exception : le répertoire `ee/` est réservé aux futurs modules d'entreprise
et est couvert par la [GenOffice Enterprise License](../../ee/LICENSE).

Les noms et logos GenOffice et Genspark sont des marques déposées de
Mainfunc, Inc. La licence Apache-2.0 n'accorde pas la permission de les
utiliser (voir la section 6) ; les forks doivent utiliser leur propre
identité visuelle.
