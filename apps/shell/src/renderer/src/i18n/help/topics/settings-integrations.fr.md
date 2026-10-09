# Paramètres, langue, thème et intégrations MCP

## Ouvrir les paramètres

La ligne de compte en bas à gauche de l’accueil ouvre le panneau des paramètres (elle affiche Se connecter lorsque vous êtes déconnecté). Il comporte six sections : Compte, Modèle IA, Médias IA et recherche, Général, Intégrations et À propos.

![Paramètres ▸ Général, où se trouvent la langue, le thème, l’enregistrement automatique et l’interrupteur des statistiques d’utilisation](img/settings-general.png)

La configuration des modèles a son propre article ; dans **Médias IA et recherche**, vous activez, par fournisseur, la génération d’images, l’analyse d’images, l’analyse vidéo, la recherche web et la recherche de fichiers locaux.

## Langue

- Les paramètres proposent **21 langues d’interface** : anglais, chinois simplifié, japonais, coréen, français, allemand, espagnol, thaï, indonésien, russe, arabe, portugais, italien, polonais, tchèque, néerlandais, malais, hébreu, hindi, chinois traditionnel, vietnamien.
- Le changement s’applique immédiatement et est conservé ; la barre de menus native est reconstruite dans la langue choisie.

## Thème

Clair / Sombre / Suivre le système. Suivre le système suit l’apparence de l’OS, et les éditeurs changent de apparence en même temps sans clignotement.

## Général

- **Envoyer des statistiques d'utilisation anonymes** — activé par défaut. Utilise Google Analytics 4 et envoie votre adresse IP publique et les métadonnées de transport ; le contenu des documents et les noms de fichiers ne sont jamais collectés, et chaque événement ne porte qu'un type comme « ouverture d'un .docx ». Vous pouvez le désactiver ici à tout moment.
- **Position de la barre latérale IA** (à gauche ou à droite), **Taille du texte du panneau IA** et **Correction orthographique dans le chat IA**.
- **Ouvrir le panneau IA dans les nouveaux documents** — désactivé, un nouveau document démarre avec le panneau replié, à un clic de là.
- **Enregistrer automatiquement tous les documents** active l'enregistrement automatique par défaut dans tous les éditeurs ; vous pouvez toujours le désactiver pour une fenêtre.
- **Emplacement d'enregistrement** avec un bouton **Modifier** et **Application par défaut pour les documents Office** pour attribuer .docx / .xlsx / .pptx à GenOffice.

## Médias IA et recherche

Ce ne sont pas des interrupteurs : chaque capacité choisit le fournisseur qui la sert, et la clé et l'URL de base d'un fournisseur sont saisies une seule fois et partagées :

- **Recherche web**, **Génération d'images**, **Analyse d'images** et **Analyse vidéo**, chacune avec un fournisseur, un modèle, une clé et une URL de base.
- **Recherche de fichiers locaux** s'exécute sur cette machine. Dessous se trouve le **Reclassement Jev**, **désactivé par défaut**. Activez-le et les extraits des 20 meilleurs résultats locaux — jusqu'à 1 200 caractères par document, plus les noms de fichiers et de dossiers — sont envoyés au modèle Jev de TypeSafe pour être réordonnés par pertinence. Désactivé, rien ne quitte l'appareil.

## À propos

- **Version**, le lien GitHub du projet et un bouton **Mettre une étoile sur GitHub**.
- **Canal de mise à jour** : Stable ou Bêta. Une modification prend effet immédiatement et déclenche une recherche de mise à jour ; elle ne rétrograde pas une installation Bêta vers Stable.

## Associations d’applications par défaut

Les paramètres peuvent enregistrer GenOffice comme gestionnaire de .docx / .xlsx / .pptx / .pdf et formats similaires (enregistrement d’application par défaut au niveau de la plateforme ; confirmez lorsque l’on vous le demande).

## Mentions relatives aux logiciels tiers et mises à jour

- Aide ▸ Mentions relatives aux logiciels tiers : l’inventaire complet des licences open source livré avec l’application.
- Aide ▸ Vérifier les mises à jour : déclenche une vérification manuelle ; une version plus récente propose l’installation.

## Connexion à Genspark

- Le point de connexion (paramètres ou liste des projets cloud) utilise un flux par **code d’appareil** : GenOffice affiche un code et ouvre la connexion dans le navigateur ; la suite est automatique une fois terminée.
- La connexion sert uniquement à : la liste des projets cloud et aux modèles hébergés de Genspark. Sans elle, toutes les fonctions locales et les modèles personnalisés continuent de fonctionner.
- La déconnexion se fait en un clic dans les paramètres.

## Intégration MCP (pour les utilisateurs avancés / clients IA)

**Intégrations** est le panneau qui relie GenOffice à un agent de codage, et il a son propre article : Connecter un agent de codage. La version courte — choisissez une voie (la ligne de commande, ou MCP), suivez la section correspondante, puis ouvrez une nouvelle conversation et posez votre question.

![Paramètres ▸ Intégrations : les trois étapes, puis les lignes de compétences et les options MCP](img/settings-integrations.png)

Sous **Serveur HTTP local**, l’application peut aussi lancer le serveur elle-même — un interrupteur d’activation et un port —, et **Avancé** ajoute l’URL du test de santé et le fichier journal, au lieu de le laisser à l’assistant. Elle n’écoute que sur localhost.

## Aide-mémoire de la ligne de commande

| Commande           | Ce qu’elle fait                 |
| ------------------ | ------------------------------- |
| `genoffice <file>` | ouvrir un fichier               |
| `genoffice mcp`    | démarrer le serveur MCP local   |
| `genoffice --help` | toutes les commandes et options |
