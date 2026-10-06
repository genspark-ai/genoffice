# Modèles IA et paramètres

## Fournisseurs et modèles

Les modèles et les clés se configurent dans les Paramètres (la ligne de compte en bas à gauche de l’accueil) :

![La fenêtre des paramètres](img/settings-general.png)

- **Genspark hébergé** : connectez-vous (flux par code d’appareil) et utilisez-le — aucune configuration.
- **Points de terminaison personnalisés (BYOK)** : Paramètres ▸ IA prend une URL de base et une clé API par protocole — compatible OpenAI, Anthropic, Gemini, DeepSeek, DashScope (qwen) et plus. Les clés ne vivent que dans les en-têtes de requête — jamais sur le disque, dans les journaux ou dans l’environnement des sous-processus.
- Un modèle différent peut être choisi par capacité : chat/génération, génération d’images, analyse d’images.
- **Tester la connexion** : vérifie que le point de terminaison est joignable et que le modèle est visible avant d’enregistrer.
- Les URL de base peuvent porter un chemin et une chaîne de requête (style passerelle) ; les chemins de point de terminaison sont concaténés correctement.

## Intégration CLI (classe Codex)

- Les paramètres acceptent le chemin d’un programme CLI local (les répertoires personnel non ASCII et un préfixe ~ fonctionnent ; ~ est développé automatiquement) ; Détecter les modèles sonde les modèles disponibles du CLI.
- La validation vérifie uniquement l’existence — aucune restriction de jeu de caractères.

## Quand les changements s’appliquent

- Les changements de modèle et de point de terminaison s’appliquent immédiatement ; une conversation en cours conserve l’ancienne configuration jusqu’à son tour suivant.
