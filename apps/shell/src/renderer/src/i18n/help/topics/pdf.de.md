# PDF: lesen, kommentieren und schwärzen

Der PDF-Editor hat fünf Tabs im Menüband: **Start / Anmerken / Bearbeiten / Seiten / Ansicht**. Er liest und schreibt: Text ist bearbeitbar, Inhalte lassen sich schwärzen und signieren, und Formulare lassen sich ausfüllen.

## Lesen und navigieren

- Linke Seitenleiste: **Miniaturen** (anklicken zum Springen, der sichtbare Bereich ist hervorgehoben) oder **Gliederung** (Lesezeichen, sofern vorhanden).
- Zoom: die Verhältnissteuerung unten rechts; Strg+Radrad zoomt stufenweise.
- Drehung: einzelne oder alle Seiten über das Menü „Seiten“; Drehungen werden beim Speichern zurückgeschrieben.
- Suche: Strg+F als Volltextsuche, alle Treffer werden hervorgehoben.
- Verschlüsselte PDFs: Eine Passwortabfrage (in einem eigenen kleinen Fenster) öffnet sie; das Passwort gilt nur für diese Sitzung.

## Textauswahl und Markierungen (Anmerken)

Probieren Sie es an einem beliebigen Absatz aus:

1. **Ziehen Sie die Maus über einen Satz** — beim Loslassen schwebt eine Annotationsleiste darüber:

![Die Annotationsleiste nach der Textauswahl](img/pdf-highlight.png)

2. Wählen Sie **Hervorheben** (das gelbe Feld öffnet eine Farbpalette), **Unterstreichen** oder **Durchstreichen**; **KI fragen** sendet die Auswahl mit Ihrer Frage an den KI-Bereich.
3. Um eine Annotation rückgängig zu machen, wählen Sie dieselbe Stelle erneut durch Ziehen aus und klicken Sie auf die aktive Schaltfläche der Leiste (ein Umschalter im Word-Stil) oder wählen Sie sie aus und drücken Sie Entf.

Referenz:

- Über Text ziehen lässt eine Popup-Leiste erscheinen: **hervorheben / unterstreichen / durchstreichen / kopieren / KI fragen**.
- Die Farben kommen aus der Palette; **dieselbe Markierung erneut auf einen bereits markierten Bereich anzuwenden entfernt sie** (Umschalter im Word-Stil).
- Bereits in der Datei gespeicherte Markierungen lassen sich auswählen und löschen (⋯-Menü oder Entf).
- **Hinweis**: Solange ein Zeichnungswerkzeug aktiviert ist, ist die Textebene nicht auswählbar — das Werkzeug deaktiviert sich nach jeder Platzierung, Sie sind also für den nächsten Schritt wieder im Auswahlmodus.

## Zeichnungswerkzeuge (Anmerken)

Sechs Werkzeuge: **Freihand, Rechteck, Ellipse, Pfeil, Notiz** sowie **Bereich schwärzen** auf dem Tab „Anmerken“.

- Jedes Werkzeug ist ein Umschalter: Klicken Sie, um es zu aktivieren; **es deaktiviert sich, sobald eine Form platziert ist** (für weitere Formen erneut anklicken); ein Klick auf das aktive Werkzeug deaktiviert es ebenfalls.
- Freihand folgt der Strichstärke; Rechteck/Ellipse/Pfeil werden durch Ziehen aufgezogen; die Farben kommen aus der Zeichnungspalette.
- Platzierte Formen lassen sich auswählen, löschen, verschieben und (Rechteck/Ellipse) in der Größe ändern.
- **Schwärzen, der vollständige Ablauf** (um eine Textzeile zu verbergen):

  1. Tab „Anmerken“ ▸ klicken Sie auf **Bereich schwärzen** (das Werkzeug wird aktiviert).
  2. **Ziehen Sie ein Rechteck über den Inhalt** — er wird mit einer Schraffur überdeckt, und die Werkzeugleiste erhält die Schaltflächen **Markierungen löschen / Schwärzungen anwenden**:

  ![Die Seite nach dem Markieren einer Schwärzung](img/pdf-redact.png)

  3. Klicken Sie auf **Schwärzungen anwenden** und bestätigen — das Ergebnis ist eine Arbeitskopie, in der der überdeckte Text und die Bilder physisch entfernt sind (nicht bloß verdeckt), und das lässt sich nicht rückgängig machen; das Originaldokument bleibt unberührt.

  Ein Fehler? Die Schaltfläche „Markierungen löschen“ entfernt die aktuellen Markierungen, damit Sie neu zeichnen können.

## Haftnotizen und Kommentarverläufe

- Das **Notiz-Werkzeug** setzt eine Stecknadel und öffnet eine Randkarte für den Text (Autorenname einstellbar); nach dem Bestätigen wird sie als Standard-PDF-Textannotierung gespeichert.
- Klicken Sie auf eine Stecknadel, um den Verlauf zu öffnen: **Antworten** (flache Verläufe im WPS-/Acrobat-Stil), **Bearbeiten** Ihres Kommentars, **Löschen** eines Kommentars oder eines ganzen Verlaufs.
- In Bearbeitung befindliche Inhalte bleiben erhalten, bis das Speichern den neuen Text in dieselbe Annotierung der Datei zurückschreibt; so bleiben die Antwortketten intakt.

## PDF-Inhalte bearbeiten (Bearbeiten)

- **Text bearbeiten**: Klicken Sie auf Text, um ihn Block für Block zu bearbeiten (pdfium-Engine; Schriften werden möglichst passend gewählt).
- **Text einfügen**: Setzen Sie durchsuchbaren Text mit Schriftart, Größe und Farbe Ihrer Wahl ab.
- **Bild / Stempel einfügen**.
- Formulare: AcroForm-Felder lassen sich direkt ausfüllen; die Werte werden beim Speichern geschrieben.

## Signaturen

- **Handschriftliche Signatur**: zeichnen Sie sie; sie kann einem Signaturfeld eines Formulars zugeordnet werden.
- **Bildsignatur**: setzen Sie ein Bild als Signatur ab.
- Gespeicherte Signaturen lassen sich wiederverwenden.

## Seitenvorgänge (Seiten)

- **Drehen / löschen / neu anordnen**: Ziehen Sie Miniaturen zum Umsortieren; beim Löschen wird nachgefragt.
- **Seiten importieren**: ziehen Sie Seiten aus einem anderen PDF herein. **Leere Seite einfügen** fügt eine leere Seite hinzu.
- **Seiten ersetzen** tauscht einen Bereich gegen Seiten aus einer anderen Datei; **Seiten zuschneiden** beschneidet, mit der Option, das auf alle Seiten anzuwenden.
- **Seitengröße** skaliert alle Seiten auf ein einziges Papierformat; **Reihenfolge umkehren** dreht das Dokument von hinten nach vorne.
- **Seite extrahieren**: ausgewählte Seiten in ein neues PDF exportieren.
- **PDF aufteilen**: zwei Varianten — nach Bereichen in mehrere Dateien aufteilen oder jede Seite in ein Raster kleinerer Seiten schneiden.
- **PDF zusammenführen**: zwei Varianten — weitere PDFs anhängen oder mehrere Seiten auf ein Blatt legen. Die Größen werden summiert, **bevor** irgendetwas gelesen wird, und **alles über 1 GiB insgesamt wird mit einer lesbaren Fehlermeldung abgelehnt** (so bleibt der Speicher begrenzt).
- Änderungen auf Seitenebene werden beim nächsten Speichern zurückgeschrieben; „Speichern unter“ lässt das Original unberührt.

## Export und Druck

- **Als Word exportieren… / PowerPoint… / Excel…** im Menü Datei, oder dieselben drei über **PDF umwandeln** im Menüband — alles lokal, kein Upload. Die .pptx kommt mit einem Foliensatz pro Seite heraus, die .xlsx mit einem Arbeitsblatt pro Seite. Jede fragt nach dem Speicherort.
- **Druck**: aktuelle Reihenfolge und Drehungen über den Systemdialog; Seitenbereiche werden unterstützt.

## Speichern

- Normales Speichern bzw. AutoSpeichern schreibt Annotationen und Änderungen in die Datei zurück (atomar).
- **Schwärzen läuft über den Ablauf „Anwenden“** und erzeugt eine Kopie, das Original bleibt unberührt — so bleibt sensible Information nicht darin zurück.
