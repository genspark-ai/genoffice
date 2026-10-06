# Sheets: Tabellenkalkulation

Sheets ist der Excel-ähnliche Editor; die Berechnung läuft in einem separaten Rust-Engine-Prozess (ein Absturz dort bringt die App nie zum Absturz). Er öffnet und speichert echte .xlsx; .csv und .tsv öffnen sich als Tabellen.

## Die Oberfläche

- **Menüband**: acht Tabs, die unten einzeln aufgeführt sind.
- **Bearbeitungsleiste**: zeigt und bearbeitet die Formel der aktiven Zelle; gängige Funktionen werden unterstützt.
- **Blattreiter** (unten): Blätter hinzufügen / umbenennen / löschen / verschieben.
- **Zellenbearbeitung**: Doppelklick oder einfach tippen; Eingabetaste bestätigt und geht nach unten, Tab geht nach rechts, Esc bricht ab (Excel-Gewohnheiten).
- **Tastenkürzel**: an der Excel-Familie ausgerichtet (ctrl+C/V/X, ctrl+Z/Y, ctrl+F, …).

## Tabs im Menüband

- **Start**: Schriftart, Füllung, Rahmen, Zahlenformate (Währung/Prozent/Tausendertrennzeichen, Dezimalstellen hinzufügen/entfernen), Ausrichtung, Verbinden, Zeilen-/Spalten einfügen und Größe, bedingte Formatierung, Als Tabelle formatieren, Zellformate, Zwischenablage und Format übertragen, Sortieren und Filtern.
- **Einfügen**: Formen, Symbole, Zeichen, Gleichung, Bildschirmfoto und mehr.
- **Seitenlayout**: Designfarben und -schriften, Schalter zum Drucken von Gitternetzlinien und Überschriften, Umbruchvorschau.
- **Formeln**: AutoSumme und Funktionseinfügung, Namen definieren (auch aus der Auswahl), Spur zum Vorgänger/Nachfolger, Überwachungsfenster, Blatt/Arbeitsmappe neu berechnen.
- **Daten**: Sortieren und Filtern (einschließlich erweiterter Filter, Filter löschen), Text in Spalten, Arbeitsmappen zusammenführen, alle aktualisieren.
- **Überprüfen**: Kommentare durchsehen (anzeigen, vorheriger/nächster), übersetzen.
- **Ansicht**: Schalter für Gitternetzlinien und Überschriften, Zoom, Normal/Umbruchvorschau.
- **Diagrammentwurf**: erscheint, wenn ein Diagramm ausgewählt ist — Diagrammtyp, Stile und Farben, Datenbereich bearbeiten.

Der Tab „Daten“, Knopf für Knopf (im Bild von links nach rechts):

![Der Tab Daten](img/sheets-data.png)

- **PivotTable**: erstellt aus dem aktuellen Bereich eine Pivot-Tabelle; Felder zum Aggregieren ziehen.
- **Aktualisieren**: berechnet die Daten der aktuellen Pivot-Tabelle neu.
- **Aus Text/CSV**: importiert eine .csv/.txt als neues Blatt, getrennt nach Trennzeichen.
- **Arbeitsmappen zusammenführen**: holt Blätter aus anderen .xlsx-Dateien in diese Arbeitsmappe.
- **Alle aktualisieren**: berechnet alle Pivot-Tabellen und externen Datenbestände neu.
- **Sortieren** (Dropdown): aufsteigend / absteigend / Benutzerdefiniertes Sortieren (Regeln über mehrere Spalten).
- **Filtern**: fügt der Kopfzeile ▼-Dropdowns hinzu; haken Sie die zu behaltenden Werte an.
- Die kleinen senkrecht gestapelten Schaltflächen daneben: **Löschen** (alle Zeilen zurückholen), **Erneut anwenden** (den aktuellen Filter erneut ausführen), **Erweitert** (mit einem Kriterienbereich filtern).
- **Text in Spalten** (Dropdown): teilt eine Spalte anhand eines Trennzeichens oder einer festen Breite in mehrere auf.
- **Blitzvorschau**: ein Beispiel genügt, und der Rest der Spalte füllt sich nach dem Muster (ctrl+E).
- **Duplikate entfernen**: entfernt doppelte Zeilen anhand der ausgewählten Spalten.
- **Datenüberprüfung** (Dropdown): Eingaberegeln für die Auswahl (Dropdown-Listen, Zahlenbereiche, …).
- **Konsolidieren**: fasst mehrere Bereiche nach Kategorien an einer Stelle zusammen.
- **Was-wäre-wenn-Analyse** (Dropdown): Zielwertsuche — löst eine Eingabezelle so, dass eine Formelzelle den Zielwert erreicht.
- **Gruppieren / Gruppierung aufheben** (Dropdown): Zeilen- oder Spaltengruppen mit Ein- und Ausklappen.
- **Teilergebnis**: fügt je Kategorie Teilergebniszeilen ein.

Der Tab „Formeln“, Knopf für Knopf:

![Der Tab Formeln](img/sheets-formulas.png)

- **Funktion einfügen** (fx): sucht Funktionen mit einem Argumentassistenten.
- **AutoSumme** (Dropdown): SUMME mit einem Klick, dazu Mittelwert/Anzahl/Maximum/Minimum.
- **Zuletzt verwendet / Finanzen / Logik / Text / Datum und Uhrzeit / Nachschlagen und Verweisen / Mathematik und Trigonometrie / Weitere**: Funktionen nach Kategorie durchsuchen und einfügen.
- **Namens-Manager**: benannte Bereiche anzeigen, erstellen und löschen.
- **Namen definieren** (Dropdown): benennt die Auswahl; **In Formel verwenden** fügt einen vorhandenen Namen ein; **Aus Auswahl erstellen** benennt Bereiche nach ihrer Kopfzeile oder -spalte.
- **Spur zum Vorgänger / Spur zum Nachfolger**: blaue Pfeile zeigen, woher die Daten einer Formel kommen und wohin sie fließen; **Pfeile entfernen** löscht sie.
- **Formeln anzeigen**: Zellen zeigen die Formel selbst statt des Ergebnisses.
- **Fehlerüberprüfung**: findet und erklärt Formelfehler.
- **Überwachungsfenster**: heben Sie Zellen, die Sie interessieren, fest und behalten Sie ihre aktuellen Werte im Blick.
- **Berechnungsoptionen** (Dropdown): automatische oder manuelle Neuberechnung; im manuellen Modus lösen **Jetzt berechnen** und **Blatt berechnen** sie von Hand aus.

## Zahlen und Formatierung

- Zahlenformate: Standard, Zahl, Währung, Prozent, Datum/Uhrzeit, Bruch, wissenschaftlich und mehr.
- Ausrichtung, Zeilenumbruch, verbundene Zellen, Rahmen und Füllungen.
- Zeilenhöhen und Spaltenbreiten per Ziehen; ein Doppelklick auf eine Grenze passt automatisch an.

## Daten

**Sortieren und Filtern** (beispielsweise nach einer Spalte absteigend):

1. Klicken Sie auf **eine beliebige Zelle dieser Spalte** (die ganze Spalte muss nicht markiert werden).
2. Tab Start ▸ **Sortieren und Filtern** ▸ **Absteigend**; ganze Zeilen werden gemeinsam neu geordnet (der Bereich wird als Ganzes sortiert).
3. Für eigene Regeln (mehrere Spalten, nach Farbe): derselbe Weg, **Benutzerdefiniertes Sortieren**.
4. Filtern: Markieren Sie die Kopfzeile und klicken Sie auf **Sortieren und Filtern ▸ Filtern** — jede Kopfzeile erhält ein ▼-Dropdown, in dem Sie die zu behaltenden Werte anhaken; **Filter löschen** holt alles zurück.

- Sortieren und Filtern.
- Fenster fixieren.
- .csv / .tsv: öffnet direkt als Tabelle (tabulatorgetrenntes tsv wird als eine Tabelle gelesen); beim Speichern wird im Originalformat zurückgeschrieben.

## Rechtsklick-Menüs

- **Im Raster**: das eigene Menü des Editors (Univer) — Ausschneiden/Kopieren/Einfügen, Zeilen/Spalten einfügen und löschen, ausblenden, Zellen verbinden, Fenster fixieren und weitere alltägliche Einträge.
- **Auf der Statusleiste unten**: wählen Sie, welche Statistiken die Statusleiste anzeigt (Mittelwert / Anzahl / Summe, …); die Auswahl bleibt erhalten.
- **Auf einem Blattreiter unten**: Blätter hinzufügen / umbenennen / löschen / einfärben / ausblenden (Tab-Menü von Univer).
- Das Kontextmenü der Tab-Leiste oben ist in [Tabs und Fensterverwaltung](help://tabs-and-windows) behandelt.

## KI

- Der KI-Bereich seitlich: Bereich auswählen und in normaler Sprache anweisen (umformatieren, Daten erzeugen, Formeln schreiben).
- Änderungen der KI lassen sich im Bereich zurückrollen.

## Speichern und exportieren

- Speichert .xlsx (Formeln und Formatierungen bleiben erhalten); Speichern unter; Export nach PDF folgt dem Druckumbruch.
- AutoSpeichern folgt der globalen Regel (aktiv nach dem ersten manuellen Speichern).

## Stabilität

- Die Rechen-Engine ist prozessisoliert von der Oberfläche: Bringt extremen Daten sie zum Absturz, bekommen Sie eine Meldung und einen Versuch, die Sitzung wiederherzustellen — keinen App-Absturz.
