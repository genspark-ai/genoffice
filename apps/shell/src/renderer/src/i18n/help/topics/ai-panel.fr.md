# Le panneau d’assistant IA

Chaque éditeur peut appeler le panneau IA : sélectionnez quelque chose, donnez une instruction, regardez le résultat en flux.

## L’ouvrir et s’en servir

![Le panneau IA dans Docs](img/ai-panel.png)

- Entrées : le **bouton IA** du ruban de chaque éditeur, **Demander à l’IA** dans les menus contextuels, ou Demander à l’IA sur la barre de marquage.
- Décrivez la tâche en langage naturel (réécrire ceci / transformer cette colonne en pourcentages / remettre cette page en page…) et appuyez sur Entrée.
- Les réponses sont rendues **en flux** ; lorsque l’IA a besoin d’outils (lire le document, le modifier, exécuter un script), elle les exécute et continue jusqu’au bout.
- **Arrêter** : interrompez le tour en cours à tout moment.

## Ce qu’il sait faire

- **Docs** : réécriture/développement/traduction/résumé, insertion de tableaux et d’images, ajustement de la mise en forme ; chaque tour enregistre d’abord un instantané.
- **Sheets** : formules, remplissage de données, transformations par lots (au besoin via le bac à sable run_script), mise en forme.
- **Slides** : génération d’un diaporama complet, ajustement de la disposition, réécriture des textes.
- **PDF** : questions et résumés sur le texte ou les pages sélectionnées.
- **Markdown / HTML** : réécriture, développement, traduction.

## Retour arrière et sécurité

- Le panneau de Docs conserve une **liste de versions** : un instantané par tour, revenez à n’importe lequel, et ce retour arrière est lui-même annulable par Ctrl+Z. Les instantanés survivent à la réouverture du document.
- Les modifications de l’IA passent par le même pipeline d’édition que les modifications manuelles (annulables, soumises à l’enregistrement) — rien ne contourne votre confirmation d’enregistrement.

## Confidentialité

- Les instructions et le contenu pertinent du document vont au **service de modèle que vous avez configuré** (Genspark hébergé ou un point de terminaison personnalisé, chapitre suivant) ; sans configuration, rien n’est envoyé.
- Les fichiers locaux ne sont envoyés nulle part ailleurs ; les clés BYOK ne vivent que dans les en-têtes de requête — jamais sur le disque ni dans les journaux.
