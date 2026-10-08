# Der KI-Bereich

Jeder Editor kann den KI-Bereich aufrufen: Etwas auswählen, eine Anweisung geben und das gestreamte Ergebnis ansehen.

## Öffnen und verwenden

![Der KI-Bereich in Docs](img/ai-panel.png)

- Einstiegspunkte: die **KI-Schaltfläche** im Menüband jedes Editors, **KI fragen** in den Kontextmenüs oder KI fragen in der Markierungsleiste.
- Beschreiben Sie die Aufgabe in normaler Sprache (dies umformulieren / diese Spalte in Prozent umrechnen / diese Seite neu setzen …) und drücken Sie die Eingabetaste.
- Antworten werden **gestreamt**; wenn die KI Werkzeuge braucht (Dokument lesen, bearbeiten), führt sie sie aus und macht weiter, bis die Aufgabe erledigt ist.
- **Stopp**: unterbrechen Sie den laufenden Durchgang jederzeit.

## Was er kann

- **Docs**: umformulieren, ausbauen, übersetzen, zusammenfassen, Tabellen und Bilder einfügen, Formatierung anpassen; jeder Durchgang legt zuerst einen Snapshot an.
- **Sheets**: Formeln, Datenfüllung, Stapeltransformationen, Formatierung.
- **Slides**: Erzeugen eines kompletten Decks, Layout anpassen, Texte umformulieren.
- **PDF**: Fragen und Zusammenfassungen zum ausgewählten Text oder zu ausgewählten Seiten.
- **Markdown / HTML**: umformulieren, ausbauen, übersetzen.

## Zurückrollen und Sicherheit

- Der Docs-Bereich führt eine **Versionsliste**: einen Snapshot pro Durchgang, Sie können zu jedem davon zurückrollen, und auch das Zurückrollen selbst lässt sich mit Strg+Z rückgängig machen. Snapshots überstehen das erneute Öffnen des Dokuments.
- KI-Änderungen laufen durch dieselbe Bearbeitungspipeline wie manuelle Änderungen (rückgängig machbar, speichern erforderlich) — nichts umgeht Ihre Speicherbestätigung.

## Datenschutz

- Anweisungen und der betreffende Dokumentinhalt gehen an den **von Ihnen konfigurierten Modelldienst** (Genspark gehostet oder ein eigener Endpunkt, siehe nächstes Kapitel); ohne Konfiguration wird nichts gesendet.
- Lokale Dateien werden nirgendwohin sonst hochgeladen; BYOK-Schlüssel werden in der Einstellungsdatei der App auf diesem Rechner gespeichert und nur in den Anfrage-Headern gesendet.
