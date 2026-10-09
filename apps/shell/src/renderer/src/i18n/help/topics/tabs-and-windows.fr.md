# Onglets et gestion des fenêtres

Tous les fichiers ouverts partagent une seule fenêtre ; la barre d’onglets en haut permet de passer de l’un à l’autre, à la manière d’un navigateur.

![La barre d’onglets : Accueil plus deux documents](img/tabs.png)

## Principes

- **Basculer** : cliquez sur un onglet, ou faites défiler la molette au-dessus de la barre pour les parcourir.
- **Fermer** : la × de l’onglet, ou ⌘W/ctrl+W. Les modifications non enregistrées déclenchent une demande de confirmation.
- **Nouveau** : le + à l’extrémité droite de la barre.
- **L’activation est immédiate** : appuyer sur un onglet bascule tout de suite, sans attendre la fin du clic.

## Renommer : double-cliquez sur un onglet

**Double-cliquez sur n’importe quel onglet de fichier** et le titre devient un champ de saisie sur place : Entrée valide, Échap annule, la perte de focus valide ; l’Entrée qui confirme un candidat de l’IME n’est pas prise pour une validation. Le renommage passe par les mêmes contrôles de sécurité que celui d’une ligne de l’accueil (caractères interdits, conflits de nom), renomme le fichier sur le disque et actualise la liste des récents.

## Réordonner par glisser-déposer

Maintenez un onglet et faites-le glisser latéralement pour le réordonner ; les voisins s’écartent en direct et l’ordre se fixe au relâchement. Une zone morte de 4 pixels évite que les simples clics ne fassent bouger les onglets.

## Menu contextuel

Un clic droit sur un onglet propose **Ouvrir dans une nouvelle fenêtre** — disponible sur chaque onglet de document détachable — et **Fermer**, grisé sur un onglet qui ne peut pas être fermé. Les deux suivent la langue de l’application.

## Détacher dans une fenêtre

**Faites glisser un onglet hors de la barre** (ou utilisez Ouvrir dans une nouvelle fenêtre) et il devient sa propre fenêtre, avec le document vivant ; faites-le glisser à nouveau pour le rattacher à la barre. Une fenêtre détachée est une fenêtre à part entière — continuez à y modifier et à y enregistrer.

## La liste de tous les onglets

Quand les onglets débordent, la petite icône de la barre d’onglets à l’extrémité droite ouvre la liste complète (un menu natif, jamais masqué par la zone de contenu) ; choisissez avec les flèches.

## La barre d’outils de chaque éditeur

Chaque onglet d’éditeur a une barre d’outils en haut (la disposition exacte varie légèrement selon l’éditeur) :

- **Enregistrer** (⌘S/ctrl+S) et **Enregistrer sous**.
- **Annuler / Rétablir** : sur plusieurs niveaux ; chaque éditeur conserve son propre historique.
- **Rechercher** (ctrl+F) : ouvre le panneau de recherche de cet éditeur.
- **Enregistrement automatique** : activé, les modifications sont écrites sur le disque à intervalles réguliers ; désactivé, seuls les enregistrements manuels écrivent (la marque de non-enregistrement sur l’onglet ou le titre vous y aide).

## L’onglet Accueil

L’onglet Accueil le plus à gauche ne peut pas être fermé ; pour revenir d’un éditeur, cliquez dessus ou utilisez Fichier ▸ **Retour à l’accueil**.
