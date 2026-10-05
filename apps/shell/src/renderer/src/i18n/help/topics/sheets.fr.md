# Sheets : tableurs

Sheets est l’éditeur de tableurs analogue à Excel ; le calcul s’exécute dans un processus séparé, un moteur Rust (une panne à cet endroit n’entraîne jamais l’arrêt de l’application). Il ouvre et enregistre de vrais .xlsx ; .csv et .tsv s’ouvrent comme des tables.

## L’interface

- **Ruban** : huit onglets, présentés un par un ci-dessous.
- **Barre de formule** : affiche et modifie la formule de la cellule active ; les fonctions courantes sont prises en charge.
- **Onglets de feuille** (en bas) : ajouter / renommer / supprimer / déplacer des feuilles.
- **Édition des cellules** : double-clic ou simple saisie ; Entrée confirme et descend, Tab va à droite, Échap annule (les habitudes d’Excel).
- **Raccourcis** : alignés sur la famille Excel (ctrl+C/V/X, ctrl+Z/Y, ctrl+F, …).

## Onglets du ruban

- **Accueil** : police, remplissage, bordures, formats de nombre (devise/pourcentage/milliers, décimales à la hausse ou à la baisse), alignement, fusion, insertion et dimension des lignes et colonnes, mise en forme conditionnelle, mise sous forme de tableau, styles de cellule, presse-papiers et copier la mise en forme, trier et filtrer.
- **Insertion** : formes, icônes, symboles, équation, capture d’écran et plus encore.
- **Mise en page** : couleurs et polices du thème, interrupteurs d’impression du quadrillage et des en-têtes, aperçu des sauts de page.
- **Formules** : Somme automatique et insertion de fonction, définir les noms (aussi depuis la sélection), repérer les antécédents et les dépendants, Fenêtre Espion, recalculer la feuille ou le classeur.
- **Données** : trier et filtrer (y compris filtre avancé, effacer le filtre), convertir, fusionner des classeurs, tout actualiser.
- **Révision** : parcourir les commentaires (afficher, précédent/suivant), traduire.
- **Affichage** : interrupteurs quadrillage et en-têtes, zoom, Normal / Aperçu des sauts de page.
- **Création de graphique** : apparaît lorsqu’un graphique est sélectionné — type de graphique, styles et couleurs, modification de la plage de données.

L’onglet Données, bouton par bouton (de gauche à droite sur l’image) :

![L’onglet Données](img/sheets-data.png)

- **Tableau croisé dynamique** : construit un tableau croisé à partir de la plage courante ; faites glisser les champs pour agréger.
- **Actualiser** : recalcule les données du tableau croisé courant.
- **À partir de texte/CSV** : importe un .csv/.txt comme nouvelle feuille, en le séparant selon un délimiteur.
- **Fusionner des classeurs** : amène les feuilles d’autres fichiers .xlsx dans celui-ci.
- **Actualiser tout** : recalcule tous les tableaux croisés et jeux de données externes.
- **Trier** (menu déroulant) : croissant / décroissant / tri personnalisé (règles sur plusieurs colonnes).
- **Filtrer** : ajoute des menus ▼ à la ligne d’en-tête ; cochez les valeurs à conserver.
- Les petits boutons empilés à côté : **Effacer** (revenir à toutes les lignes), **Réappliquer** (relancer le filtre courant), **Avancé** (filtrer avec une plage de critères).
- **Texte en colonnes** (menu déroulant) : scinde une colonne en plusieurs selon un délimiteur ou une largeur fixe.
- **Remplissage instantané** : donnez un exemple et le reste de la colonne se remplit sur le modèle (ctrl+E).
- **Supprimer les doublons** : retire les lignes en double d’après les colonnes sélectionnées.
- **Validation des données** (menu déroulant) : règles de saisie pour la sélection (listes déroulantes, plages de nombres…).
- **Consolider** : agrège plusieurs plages en un seul endroit par catégorie.
- **Analyse de scénarios** (menu déroulant) : valeur cible / tables de données.
- **Grouper / Dissocier** (menu déroulant) : groupes de lignes ou de colonnes avec repli et développement.
- **Sous-total** : insère des lignes de sous-total par catégorie.

L’onglet Formules, bouton par bouton :

![L’onglet Formules](img/sheets-formulas.png)

- **Insérer une fonction** (fx) : recherche des fonctions avec un assistant d’arguments.
- **Somme automatique** (menu déroulant) : SOMME en un clic, plus moyenne/compte/max/min.
- **Fonctions récentes / Financières / Logiques / Texte / Date et heure / Recherche et référence / Maths et trigonométrie / Plus** : parcourez et insérez les fonctions par catégorie.
- **Gestionnaire de noms** : afficher, créer et supprimer des plages nommées.
- **Définir un nom** (menu déroulant) : nomme la sélection ; **Utiliser dans la formule** insère un nom existant ; **Créer à partir de la sélection** nomme des plages d’après leur ligne ou colonne d’en-tête.
- **Repérer les antécédents / Repérer les dépendants** : des flèches bleues montrent d’où viennent les données d’une formule et où elles alimentent ; **Supprimer les flèches** les efface.
- **Afficher les formules** : les cellules montrent la formule elle-même au lieu du résultat.
- **Vérification des erreurs** : localise et explique les erreurs de formule.
- **Fenêtre Espion** : épinglez les cellules qui vous intéressent et suivez leurs valeurs en direct.
- **Options de calcul** (menu déroulant) : recalcul automatique ou manuel ; en mode manuel, **Calculer maintenant** et **Calculer la feuille** déclenchent le recalcul à la main.

## Nombres et formats

- Formats de nombre : standard, nombre, devise, pourcentage, date/heure, fraction, scientifique et plus encore.
- Alignement, retour à la ligne, cellules fusionnées, bordures et remplissages.
- Hauteurs de lignes et largeurs de colonnes par glissement ; double-cliquez sur une bordure pour ajuster automatiquement.

## Données

**Trier et filtrer** (par exemple, par ordre décroissant sur une colonne) :

1. Cliquez sur **n’importe quelle cellule de cette colonne** (inutile de sélectionner toute la colonne).
2. Onglet Accueil ▸ **Trier et filtrer** ▸ **Décroissant** ; les lignes entières se réordonnent ensemble (la zone est triée comme un tout).
3. Pour des règles personnalisées (plusieurs colonnes, par couleur) : même chemin, **Tri personnalisé**.
4. Filtrage : sélectionnez la ligne d’en-tête et cliquez sur **Trier et filtrer ▸ Filtrer** — chaque en-tête reçoit un menu ▼ où vous cochez les valeurs à garder ; effacer le filtre ramène tout.

- Trier et filtrer.
- Figer les volets.
- .csv / .tsv : s’ouvrent directement comme une table (un tsv séparé par des tabulations est analysé comme une seule table) ; l’enregistrement réécrit au format d’origine.

## Menus contextuels

- **Dans la grille** : le menu propre à l’éditeur (Univer) — couper/copier/coller, insérer et supprimer lignes/colonnes, masquer, fusionner des cellules, figer les volets et autres items du quotidien.
- **Sur la barre d’état du bas** : choisissez les statistiques affichées dans la barre d’état (moyenne / nombre / somme, …) ; le choix est conservé.
- **Sur un onglet de feuille en bas** : ajouter / renommer / supprimer / colorer / masquer des feuilles (menu d’onglet d’Univer).
- Le menu contextuel de la barre d’onglets en haut est traité dans [Onglets et gestion des fenêtres](help://tabs-and-windows).

## L’éditeur de scripts (avancé)

Sheets fournit un **éditeur de scripts** avec une API calquée sur Google Apps Script pour les traitements par lots.

- Ouverture : Outils ▸ Éditeur de scripts (ou le menu développeur, selon la version).
- Interface : bibliothèque de scripts à gauche (nouveau/supprimer), éditeur de code et volet de sortie à droite ; boutons **Exécuter / Arrêter**.
- L’API est asynchrone :

```js
const sheet = await SpreadsheetApp.getActiveSpreadsheet()
const active = await sheet.getActiveSheet()
const range = await active.getRange('A1:C10')
const values = await range.getValues() // tableau 2D
await range.setValues(values.map((row) => row.map((v) => v * 2)))
Logger.log('done')
```

- `SpreadsheetApp` (point d’entrée), `Sheet` (getName/getRange/getLastRow…), `Range` (getValue(s)/setValue(s)/clear…), `Logger.log`, `Utilities.sleep`.
- Les scripts s’exécutent dans un **worker isolé** : pas de réseau, pas de système de fichiers, pas de DOM — ils ne peuvent toucher que par l’API ci-dessus, si bien qu’un script bogué ou malveillant ne peut atteindre rien d’autre.
- Le volet de sortie plafonne son nombre de lignes pour qu’un journal énorme ne bloque pas l’interface.
- **L’IA peut aussi exécuter des scripts** : l’outil `run_script` de l’assistant exécute la même API isolée — idéal pour les transformations par lots fondées sur des règles.

## IA

- Le panneau IA latéral : sélectionnez une plage et donnez votre instruction en langage naturel (reformater, générer des données, écrire des formules).
- Les modifications de l’IA peuvent être annulées depuis le panneau.

## Enregistrement et export

- Enregistre du .xlsx (formules et formats préservés) ; Enregistrer sous ; l’export PDF suit la pagination d’impression.
- L’enregistrement automatique suit la règle globale (activé après la première sauvegarde manuelle).

## Stabilité

- Le moteur de calcul est isolé de l’interface au niveau du processus : si des données extrêmes le tuent, vous obtenez un message et une tentative de récupération de session — pas un plantage de l’application.
