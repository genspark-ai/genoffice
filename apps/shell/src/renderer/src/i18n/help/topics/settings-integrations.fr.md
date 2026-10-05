# Paramètres, langue, thème et intégrations MCP

## Ouvrir les paramètres

La ligne de compte en bas à gauche de l’accueil ouvre le panneau des paramètres (elle affiche Se connecter lorsque vous êtes déconnecté) ; les options liées à l’IA se trouvent dans sa section Modèle IA

![La fenêtre des paramètres](img/settings-integrations.png) — la configuration des modèles est traitée dans Modèles IA et paramètres.

## Langue

- Les paramètres proposent **21 langues d’interface** : anglais, chinois simplifié, japonais, coréen, français, allemand, espagnol, thaï, indonésien, russe, arabe, portugais, italien, polonais, tchèque, néerlandais, malais, hébreu, hindi, chinois traditionnel, vietnamien.
- Le changement s’applique immédiatement et est conservé ; la barre de menus native est reconstruite dans la langue choisie.

## Thème

Clair / Sombre / Suivre le système. Suivre le système suit l’apparence de l’OS, et les éditeurs changent de apparence en même temps sans clignotement.

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

GenOffice embarque un **serveur MCP** local pour que des clients IA externes (Claude Desktop, Cursor, …) puissent lire et écrire directement vos documents :

- Démarrage : `genoffice mcp` en ligne de commande (port et jeton d’authentification configurables ; boucle locale seule par défaut).
- Capacités : créer/ouvrir/modifier des docx, xlsx et pptx, lire le contenu, convertir des formats, exporter en PDF et plus encore — le même ensemble d’outils que les applications de bureau.
- Sécurité : l’authentification par jeton est facultative mais recommandée ; l’écoute reste sur la machine locale par défaut ; voir `genoffice mcp --help`.

## Aide-mémoire de la ligne de commande

| Commande           | Ce qu’elle fait                 |
| ------------------ | ------------------------------- |
| `genoffice <file>` | ouvrir un fichier               |
| `genoffice mcp`    | démarrer le serveur MCP local   |
| `genoffice --help` | toutes les commandes et options |
