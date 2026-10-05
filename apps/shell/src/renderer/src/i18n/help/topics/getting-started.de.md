# Schnellstart: die Oberfläche und die Grundlagen

GenOffice ist ein Office-Paket, das vollständig auf Ihrem Rechner läuft: ein Fenster, eine Reihe von Tabs, darin sechs Editoren — Docs (Textverarbeitung), Sheets (Tabellenkalkulation), Slides (Präsentationen), PDF, Markdown und HTML. Die Dateien sind echte .docx / .xlsx / .pptx / .pdf und lassen sich vollständig mit Word, Excel und PowerPoint austauschen. Kein Netzwerk nötig.

## Überblick über die Oberfläche

![Der Startbildschirm](img/home-screen.png)

Das Fenster besteht aus drei Teilen:

- **Tab-Leiste (oben)**: Jede geöffnete Datei ist ein Tab. Der Start-Tab ganz links ist immer vorhanden und lässt sich nicht schließen; die übrigen Tabs sind Ihre Dokumente. Ein Doppelklick auf einen Tab benennt die Datei direkt dort um.
- **Inhaltsbereich**: der zum aktiven Tab gehörende Editor (oder die Startseite).
- **Menü**: unter macOS in der Systemmenüleiste, unter Windows/Linux oben im Fenster. Die Menüs Datei/Bearbeiten/Ansicht wechseln passend zum aktiven Editor.

## Ein Dokument anlegen

Eines davon genügt:

- Klicken Sie auf eine Schnellerstellungskarte im Bereich [Schnellstart](help://getting-started) der **Startseite** (AI Docs, AI Sheets, AI Slides, …).
- Menü **Datei ▸ Neu**: Docs (⌘N/ctrl+N), Sheets, Slides, Markdown, HTML oder PDF.
- Ziehen Sie eine Datei auf das Fenster oder doppelklicken Sie sie im Dateimanager (wenn GenOffice die Standard-App ist).

Ein neues Dokument öffnet sich ohne Namen; die Datei auf dem Datenträger entsteht erst beim ersten Speichern.

## Dateien öffnen

- Menü **Datei ▸ Öffnen** (⌘O/ctrl+O) öffnet die Systemauswahl: .docx / .doc / .xlsx / .xls / .csv / .tsv / .pptx / .ppt / .pdf / .md / .html.
- Klicken Sie auf einen Eintrag in der Liste **Zuletzt verwendet** auf der Startseite.
- `genoffice <file>` in einem Terminal öffnet Dateien ebenfalls.

## Das Speichermodell

- **Manuelles Speichern**: ⌘S/ctrl+S oder Datei ▸ Speichern / Speichern unter. Beim ersten Speichern werden Ort und Name abgefragt.
- **AutoSpeichern** schaltet sich erst ein, nachdem Sie die Datei mindestens einmal manuell gespeichert haben — eine nur gelesene PDF wird nie stillschweigend neu geschrieben. Das AutoSpeichern greift kurz nach Inhaltsänderungen.
- Beim Schließen eines Tabs mit ungespeicherten Änderungen wird zuerst gefragt: Speichern / Verwerfen / Abbrechen.
- Jeder Schreibvorgang erfolgt atomar (temporäre Datei plus Umbenennung), sodass ein Stromausfall keine halbe Datei hinterlässt.

## Häufige Tastenkürzel

| Aktion                 | macOS | Windows / Linux |
| ---------------------- | ----- | --------------- |
| Neues Dokument         | ⌘N    | ctrl+N          |
| Öffnen                 | ⌘O    | ctrl+O          |
| Speichern              | ⌘S    | ctrl+S          |
| Tab schließen          | ⌘W    | ctrl+W          |
| Dieses Handbuch öffnen | F1    | F1              |

Tastenkürzel innerhalb der einzelnen Editoren (Format übertragen, Suchen und Ersetzen, Tabellenbefehle, …) stehen in den jeweiligen Kapiteln; Docs bringt zusätzlich einen durchsuchbaren Dialog für Tastenkombinationen mit (siehe sein Kapitel).

## Wie Sie weiterfahren

- Wo Dateien liegen: [Der Startbildschirm](help://home-screen).
- Viele geöffnete Dateien verwalten: [Tabs und Fensterverwaltung](help://tabs-and-windows).
- Die KI die Arbeit machen lassen: [Der KI-Bereich](help://ai-panel).
- Sprache, Design, Standard-Apps: [Einstellungen, Sprache, Design und MCP-Integrationen](help://settings-integrations).
