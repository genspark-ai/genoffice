# PDF : lecture, annotation et caviardage

L’éditeur PDF a cinq onglets de ruban : **Accueil / Annoter / Édition / Pages / Affichage**. Il lit et écrit : le texte est modifiable, le contenu peut être caviardé, signé, et les formulaires peuvent être remplis.

## Lecture et navigation

- Barre latérale gauche : les **Vignettes** (cliquez pour sauter, la plage visible est mise en évidence) ou les **Signets** (le sommaire, lorsqu’il existe).
- Zoom : le contrôle de ratio en bas à droite ; ctrl+molette règle le zoom par crans.
- Rotation : page par page ou pour toutes les pages depuis le menu Pages ; les rotations sont réécrites à l’enregistrement.
- Recherche : ctrl+F en texte intégral, toutes les occurrences sont mises en évidence.
- PDF chiffrés : une demande de mot de passe (dans sa propre petite fenêtre) permet de les ouvrir ; le mot de passe ne sert que pour cette session.

## Sélection de texte et marquage (Annoter)

Essayez sur n’importe quel paragraphe :

1. **Faites glisser la souris sur une phrase** — au relâchement, une barre d’annotation flotte au-dessus :

![La barre d’annotation après la sélection de texte](img/pdf-highlight.png)

2. Choisissez **Surligner** (la pastille jaune ouvre une palette de couleurs), **Souligner** ou **Barrer** ; **Demander à l’IA** envoie la sélection avec votre question au panneau IA.
3. Pour annuler une annotation, sélectionnez de nouveau le même passage par glissement et cliquez sur le bouton actif de la barre (un interrupteur à la Word), ou sélectionnez-le et appuyez sur Suppr.

Référence :

- Faites glisser sur du texte et une barre contextuelle apparaît : **surligner / souligner / barrer / copier / Demander à l’IA**.
- Les couleurs viennent de la palette ; **appliquer le même marquage à une plage déjà marquée le retire** (interrupteur à la Word).
- Les marquages déjà enregistrés dans le fichier peuvent être sélectionnés et supprimés (menu ⋯ ou Suppr).
- **Remarque** : tant qu’un outil de dessin est armé, la couche de texte n’est pas sélectionnable — l’outil se désarme après chaque placement, vous êtes donc de nouveau en mode sélection pour la suite.

## Outils de dessin (Annoter)

Six outils : **Dessin, Rectangle, Ellipse, Flèche, Note**, plus le **Caviarder une zone** sur l’onglet Édition.

- Chaque outil est un interrupteur : cliquez pour l’armer ; **il se désarme dès qu’une forme est placée** (cliquez de nouveau sur l’outil pour continuer) ; cliquer sur l’outil armé le désarme aussi.
- Le dessin suit l’épaisseur du trait ; rectangle/ellipse/flèche se tracent au glissement ; les couleurs viennent de la palette de dessin.
- Les formes placées peuvent être sélectionnées, supprimées, déplacées et (rectangle/ellipse) redimensionnées.
- **Le caviardage, procédure complète** (pour masquer une ligne de texte) :

  1. Onglet Annoter ▸ cliquez sur **Caviarder une zone** (l’outil s’arme).
  2. **Faites glisser un rectangle sur le contenu** — il est couvert d’une marque hachurée, et la barre d’outils gagne les boutons **effacer les marques / appliquer le caviardage** :

  ![La page après avoir marqué un caviardage](img/pdf-redact.png)

  3. Cliquez sur **Appliquer le caviardage** et confirmez — le résultat est une copie de travail où le texte et les images couverts sont physiquement supprimés (et non masqués), et cela est irréversible ; le document d’origine reste intact.

  Une erreur ? Le bouton effacer les marques supprime les marques actuelles pour vous permettre de redessiner.

## Pense-bêtes et fils de commentaires

- L’**outil Note** pose une épingle et ouvre une carte en marge pour ce texte (nom d’auteur configurable) ; une fois validé, elle est enregistrée comme annotation de texte PDF standard.
- Cliquez sur une épingle pour ouvrir le fil : **Répondre** (fils plats à la WPS/Acrobat), **Modifier** votre commentaire, **Supprimer** un commentaire ou un fil entier.
- Les modifications en cours survivent jusqu’à ce que l’enregistrement réécrive le nouveau texte dans la même annotation du fichier, ce qui préserve l’enchaînement des réponses.

## Modifier le contenu du PDF (Édition)

- **Modifier le texte** : cliquez sur le texte pour le modifier bloc par bloc (moteur pdfium ; polices appariées au mieux).
- **Insérer du texte** : placez un texte consultable avec la police, la taille et la couleur de votre choix.
- **Insérer une image / un tampon**.
- Formulaires : les champs AcroForm se remplissent directement ; les valeurs sont écrites à l’enregistrement.

## Signatures

- **Signature manuscrite** : dessinez-la ; elle peut être liée à un champ de signature de formulaire.
- **Signature image** : placez une image comme signature.
- Les signatures enregistrées peuvent être réutilisées.

## Opérations sur les pages (Pages)

- **Rotation / suppression / réorganisation** : faites glisser les vignettes pour réorganiser ; la suppression demande confirmation.
- **Extraire des pages** : exportez les pages sélectionnées dans un nouveau PDF.
- **Diviser** : par plages, en plusieurs fichiers.
- **Fusionner** : ajoute d’autres PDF. Les tailles sont additionnées **avant** toute lecture et **tout total supérieur à 1 Gio est refusé** avec une erreur lisible (cela garde la mémoire bornée).
- Les changements au niveau des pages sont réécrits à l’enregistrement suivant ; Enregistrer sous laisse l’original intact.

## Export et impression

- **PDF en Word** : conversion locale en .docx.
- **Imprimer** : l’ordre et les rotations courants via la boîte de dialogue du système ; plages de pages prises en charge.

## Enregistrement

- L’enregistrement normal ou automatique réécrit les annotations et les modifications dans le fichier (de façon atomique).
- **Le caviardage passe par sa procédure Appliquer**, qui produit une copie et laisse l’original intact, afin que le contenu sensible ne reste pas dedans.
