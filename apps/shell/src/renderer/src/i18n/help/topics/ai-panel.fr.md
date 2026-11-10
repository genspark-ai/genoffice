# Le panneau d’assistant IA

Chaque éditeur peut appeler le panneau IA: sélectionnez quelque chose, donnez une instruction, regardez le résultat en flux.

## L’ouvrir et s’en servir

![Le panneau IA dans Docs](img/ai-panel.png)

- Entrées: le **bouton IA** du ruban de chaque éditeur, **Demander à l’IA** dans les menus contextuels, ou Demander à l’IA sur la barre de marquage.
- Décrivez la tâche en langage naturel (réécrire ceci / transformer cette colonne en pourcentages / remettre cette page en page…) et appuyez sur Entrée.
- Les réponses sont rendues **en flux**; lorsque l’IA a besoin d’outils (lire le document, le modifier), elle les exécute et continue jusqu’au bout.
- **Arrêter**: interrompez le tour en cours à tout moment.
- **File d'attente** : pendant qu'une réponse arrive, Entrée met votre message en file plutôt que de le jeter — une bande au-dessus de la saisie les compte ; ouvrez-la pour modifier, retirer, réordonner, mettre en pause ou vider. Le message suivant part tout seul à la fin du tour ; Échap arrête toujours le tour.

## Ce qu’il sait faire

- **Docs**: réécriture/développement/traduction/résumé, insertion de tableaux et d’images, ajustement de la mise en forme; chaque tour enregistre d’abord un instantané.
- **Sheets**: formules, remplissage de données, transformations par lots, mise en forme.
- **Slides**: génération d’un diaporama complet, ajustement de la disposition, réécriture des textes.
- **PDF**: questions et résumés sur le texte ou les pages sélectionnées.
- **Markdown / HTML**: réécriture, développement, traduction.
- **Accueil ▸ Ask AI** répond à partir de vos fichiers étoilés et les cite — le panneau de base de connaissances a sa propre section au chapitre Accueil, et le panneau IA de Docs porte la même compétence.

## Retour arrière et sécurité

- Le panneau de Docs conserve une **liste de versions**: un instantané par tour, revenez à n’importe lequel, et ce retour arrière est lui-même annulable par Ctrl+Z. Les instantanés survivent à la réouverture du document.
- Les modifications de l’IA passent par le même pipeline d’édition que les modifications manuelles (annulables, soumises à l’enregistrement) — rien ne contourne votre confirmation d’enregistrement.

## Masquer une sélection à l'IA

- Dans Slides, HTML et Markdown, une sélection peut être soustraite aux yeux du modèle : clic droit ▸ **Masquer la sélection à l'IA**, nommez l'espace réservé, et le modèle lit `{{nom}}` à la place. (Markdown arme l'entrée via la bascule œil barré de son ruban ; celle de HTML vit dans le panneau source.)
- Le fichier garde le vrai texte — enregistrement, export et impression restent intacts, et la marque survit à la réouverture. La même entrée retire le masque (Markdown a une entrée distincte **Réafficher la sélection à l'IA**).
- HTML masque contenus d'éléments, valeurs d'attributs et littéraux de chaîne des scripts ; la page rendue ne change jamais.
- Slides masque aussi images, vidéos et audios entiers, et une diapositive au contenu masqué n'est jamais envoyée en capture : le contrôle automatique de mise en page saute la page et le dit.
- Cela ne vaut que pour l'IA de l'application — la CLI, les outils MCP et l'export headless lisent le fichier directement et voient toujours le texte.

## Confidentialité

- Les instructions et le contenu pertinent du document vont au **service de modèle que vous avez configuré** (Genspark hébergé ou un point de terminaison personnalisé, chapitre suivant); sans configuration, rien n’est envoyé.
- Les fichiers locaux ne sont envoyés nulle part ailleurs; les clés BYOK sont stockées dans le fichier de paramètres de l'application sur cet ordinateur et envoyées uniquement dans les en-têtes de requête.
