# Der Startbildschirm: wo Ihre Dateien liegen

Die Startseite ist GenOffices Startseite: eine Navigationsleiste links, Dateilisten und Schnellerstellungskarten rechts.

![Der Startbildschirm](img/home-screen.png)

## Navigation in der Seitenleiste

- **Zuletzt verwendet**: die zuletzt geöffneten Dateien. Jede Zeile trägt ihren Zeitpunkt — heute, gestern oder das Datum.
- **Favoriten**: die von Ihnen mit einem Stern versehenen Dateien. Fahren Sie über eine Dateizeile und klicken Sie auf den Stern, um sie hinzuzufügen oder zu entfernen.
- **Benutzerhandbuch**: öffnet dieses Handbuch.
- **Genspark Projects**: nach der Anmeldung an Ihrem Genspark-Konto werden die im Web mit Genspark AI erstellten Projekte angezeigt; klicken Sie auf eines, um die Bearbeitung im Browser fortzusetzen. Suche, Sortierung nach Zeit, Aktualisierung und Nachladen werden unterstützt.
- **Ordner**: Verzeichnisse, die Sie mit **Ordner hinzufügen…** zur Seitenleiste hinzufügen oder aus dem Dateimanager dorthin ziehen. Jedes wird zu einem Stammverzeichnis, das Sie öffnen, mit Unterordnern füllen, umbenennen und daraus entfernen können; geht eines offline, wird es als nicht verfügbar angezeigt und lässt sich aus der Liste nehmen. **Neuer Ordner** legt ein weiteres an.

Hier gibt es keinen Eintrag für den Papierkorb. Gelöschte Dateien landen im Papierkorb des Systems, und sie wiederherzustellen ist Sache des Betriebssystems.

## Die Dateiliste

Jede Zeile zeigt ein Symbol, den Dateinamen, die Änderungszeit und mehr. Das **⋯-Menü** der Zeile bietet:

- **Öffnen** und **Im Ordner anzeigen**, um die Datei im Dateimanager zu finden.
- **Pfad kopieren**.
- **In Ordner verschieben…**: öffnet eine Ordnerauswahl und verschiebt die Datei tatsächlich; ist am Ziel schon eine Datei gleichen Namens, können Sie überspringen, überschreiben oder umbenennen.
- **Umbenennen**: direkt an Ort und Stelle, die Dateiendung bleibt automatisch erhalten.
- **Zu Favoriten hinzufügen / Aus Favoriten entfernen** — Sterne bleiben über Neustarts hinweg erhalten und wandern mit der Datei mit, wenn Sie sie umbenennen.
- **Duplizieren**: erstellt eine Kopie im selben Ordner.
- **Löschen**: verschiebt die Datei in den Papierkorb des Systems — kein endgültiges Löschen.
- **Aus Liste entfernen**, in der Ansicht „Zuletzt verwendet“, um einen Eintrag zu streichen, ohne die Datei anzufassen.

### Mehrere Dateien auf einmal

Haken Sie das Kontrollkästchen einer Zeile an oder klicken Sie mit ⌘/Strg, um eine Auswahl zusammenzustellen; das Kontrollkästchen in der Kopfzeile wählt alles aus, was gerade aufgeführt ist, und eine Leiste über der Liste meldet (**{n} ausgewählt**), wie viele ausgewählt sind, und bietet **In Ordner verschieben…** sowie **Dateien löschen** für die ganze Auswahl. Sie können eine Mehrfachauswahl auch auf einen Ordner in der Seitenleiste ziehen.

## Suche

Das Suchfeld oben erfasst gleichzeitig zwei Dinge:

- **Dateinamen**: schnelles Filtern nach Namen.
- **Dateinhalte**: GenOffice indexiert Ihre Dateien im Hintergrund (Text in docx/xlsx/pptx/pdf/md/html), sodass die Volltextsuche auch Dateien findet. Umfang und Schalter stehen in den Sucheinstellungen.

## Schnellstart-Karten

Die Karten über den Listen erzeugen in einem Schritt ein neues Dokument. Ein Klick auf eine Karte erzeugt eine Datei dieses Typs und öffnet den passenden Editor — Sie legen sofort los oder lassen die KI einen Entwurf schreiben (jeder Editor hat eine **KI-Schaltfläche** im Menüband und **KI fragen** im Kontextmenü der Auswahl).

Die neue Datei landet im aktuell in der Seitenleiste ausgewählten Ordner; ohne Auswahl geht sie in den Standardordner.

Was die einzelnen Karten tun:

- **AI Docs** (.docx): ein leeres Textdokument im Docs-Editor. Die Datei wird erst beim **ersten Speichern** auf den Datenträger geschrieben; neue Dokumente öffnen mit ausgeklapptem KI-Bereich (abschaltbar unter Einstellungen → „KI-Panel in neuen Dokumenten öffnen“).
- **AI Sheets** (.xlsx): eine leere Tabelle im Sheets-Editor. Bis zum Speichern existiert keine Datei auf dem Datenträger — der Name ist für das erste Speichern reserviert; nach der ersten KI-Generierung kann die Datei außerdem automatisch aus ihrem Inhalt benannt werden.
- **AI Slides** (.pptx): eine leere Präsentation im Slides-Editor.
- **AI Markdown** (.md): ein leeres Markdown-Dokument im Markdown-Editor.
- **AI HTML** (.html): eine leere Webseite im HTML-Editor.
- **AI PDF** (.pdf): anders als die übrigen — es erzeugt **sofort** eine echte leere einseitige PDF im Zielordner und öffnet sie als gewöhnliche Datei (der PDF-Editor arbeitet mit echten Dateien). Praktisch zum Kommentieren, Schwärzen oder Hinzufügen von Text; die Datei kann beim ersten Speichern automatisch aus ihrem Inhalt benannt werden.
- **Lokale Datei öffnen**: eine Systemdateiauswahl für Word (.docx/.doc), Excel (.xlsx/.xlsm/.xls/.csv/.tsv), PowerPoint (.pptx/.ppt), PDF, Markdown (.md/.markdown) und Webseiten (.html/.htm). Mehrfachauswahl ist möglich; jede Datei bekommt ihren eigenen Tab.

> Tipp: Datei ▸ Neu in der Menüleiste erzeugt dieselben Dokumenttypen (⌘N/Ctrl+N erzeugt standardmäßig ein Textdokument); eine ins Fenster gezogene Datei wird direkt geöffnet.

## Cloud-Projekte (Genspark Projects)

- Für die erste Nutzung ist eine Anmeldung an Ihrem Genspark-Konto nötig (Gerätecode-Verfahren: GenOffice zeigt einen Code, Sie schließen die Anmeldung im Browser ab).
- Die Projektliste wird mit dem Web synchronisiert; „In Browser öffnen“ führt dorthin, um weiterzuarbeiten.
- Ohne Anmeldung ist keine der lokalen Funktionen beeinträchtigt.
