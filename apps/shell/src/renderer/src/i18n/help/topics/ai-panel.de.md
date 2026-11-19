# Der KI-Bereich

Jeder Editor kann den KI-Bereich aufrufen: Etwas auswählen, eine Anweisung geben und das gestreamte Ergebnis ansehen.

## Öffnen und verwenden

![Der KI-Bereich in Docs](img/ai-panel.png)

- Einstiegspunkte: die **KI-Schaltfläche** im Menüband jedes Editors, **KI fragen** in den Kontextmenüs oder KI fragen in der Markierungsleiste.
- Beschreiben Sie die Aufgabe in normaler Sprache (dies umformulieren / diese Spalte in Prozent umrechnen / diese Seite neu setzen …) und drücken Sie die Eingabetaste.
- Antworten werden **gestreamt**; wenn die KI Werkzeuge braucht (Dokument lesen, bearbeiten), führt sie sie aus und macht weiter, bis die Aufgabe erledigt ist.
- **Stopp**: unterbrechen Sie den laufenden Durchgang jederzeit.
- **Warteschlange**: Während eine Antwort läuft, stellt Enter Ihre Nachricht nicht ab, sondern in die Warteschlange — ein Streifen über dem Eingabefeld zählt sie; aufklappen zum Bearbeiten, Entfernen, Umsortieren, Pausieren oder Leeren. Die nächste Nachricht sendet sich selbst, wenn der Zug endet; Esc bricht weiterhin ab.

## Was er kann

- **Docs**: umformulieren, ausbauen, übersetzen, zusammenfassen, Tabellen und Bilder einfügen, Formatierung anpassen; jeder Durchgang legt zuerst einen Snapshot an.
- **Sheets**: Formeln, Datenfüllung, Stapeltransformationen, Formatierung.
- **Slides**: Erzeugen eines kompletten Decks, Layout anpassen, Texte umformulieren.
- **PDF**: Fragen und Zusammenfassungen zum ausgewählten Text oder zu ausgewählten Seiten.
- **Markdown / HTML**: umformulieren, ausbauen, übersetzen.
- **Start ▸ Ask AI** antwortet aus Ihren markierten Dateien und nennt die Quellen — das Wissensbasis-Panel hat im Start-Kapitel einen eigenen Abschnitt, und das Docs-KI-Panel beherrscht dieselbe Fähigkeit.

## Zurückrollen und Sicherheit

- Der Docs-Bereich führt eine **Versionsliste**: einen Snapshot pro Durchgang, Sie können zu jedem davon zurückrollen, und auch das Zurückrollen selbst lässt sich mit Strg+Z rückgängig machen. Snapshots überstehen das erneute Öffnen des Dokuments.
- KI-Änderungen laufen durch dieselbe Bearbeitungspipeline wie manuelle Änderungen (rückgängig machbar, speichern erforderlich) — nichts umgeht Ihre Speicherbestätigung.

## Eine Auswahl vor der KI verbergen

- In Slides, HTML und Markdown lässt sich eine Auswahl vor den Augen des Modells verbergen: Rechtsklick ▸ **Auswahl vor der KI verbergen**, dem Platzhalter einen Namen geben, und das Modell liest stattdessen `{{Name}}`. (Markdown aktiviert den Eintrag über den Augenaus-Schalter im Menüband; der HTML-Eintrag liegt im Quelltext-Panel.)
- Die Datei behält den echten Text — Speichern, Export und Druck bleiben unberührt, und die Markierung übersteht das erneute Öffnen. Dieselbe Menüposition hebt die Markierung auf (Markdown hat zusätzlich **Auswahl der KI wieder zeigen**).
- HTML verbirgt Elementinhalte, Attributwerte und String-Literale in Skripten; die gerenderte Seite ändert sich nie.
- Slides verbirgt auch ganze Bilder, Videos und Audiodateien; eine Folie mit verborgenem Inhalt wird nie als Screenshot gesendet: Die automatische Layoutprüfung überspringt die Seite und sagt das auch.
- Das gilt nur für die KI in der App — CLI, MCP-Tools und Headless-Export lesen die Datei direkt und sehen den Text weiterhin.

## Datenschutz

- Anweisungen und der betreffende Dokumentinhalt gehen an den **von Ihnen konfigurierten Modelldienst** (Genspark gehostet oder ein eigener Endpunkt, siehe nächstes Kapitel); ohne Konfiguration wird nichts gesendet.
- Lokale Dateien werden nirgendwohin sonst hochgeladen; BYOK-Schlüssel werden in der Einstellungsdatei der App auf diesem Rechner gespeichert und nur in den Anfrage-Headern gesendet.
