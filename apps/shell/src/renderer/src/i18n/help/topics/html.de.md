# Der HTML-Editor

Der HTML-Editor öffnet .html / .htm mit zwei Modi: **Vorschau** (die gerenderte Seite) und **Quelltext**.

- **Vorschau**: echtes Rendering; relative Stylesheets und Bilder werden neben der Datei geladen.
- **Vorschau-Inspektor**: Klicken, um ein Element auszuwählen, Doppelklick, um dessen Text direkt zu bearbeiten, Löschen über die Werkzeugleiste und KI fragen zur Auswahl.
- **Quelltextmodus**: HTML bearbeiten; Strg+F zum Suchen, „Alle ersetzen“ speichert das neu geschriebene Markup.
- **Speichern**: bytegetreu (BOM/CRLF/abschließender Zeilenumbruch bleiben erhalten); Speichern ohne Änderung schreibt nicht neu.
- **Zoom**: Strg+Radrad / Ziehen skaliert die Vorschau; Strg+Z in der Vorschau macht die letzte Änderung rückgängig.

## Die Werkzeugleiste

Klicken Sie in der Vorschau auf ein beliebiges Element, und eine Werkzeugleiste schwebt darüber:

![Die schwebende Werkzeugleiste über einem ausgewählten Element](img/html-toolbar.png)

- **Datei und Verlauf**: Speichern, Speichern unter, Rückgängig, Wiederholen, Suchen; der Schalter **AutoSpeichern** schreibt Änderungen in regelmäßigen Abständen.
- Umschalter **Vorschau / Quelltext**; **Vollbild** zeigt die Seite im Vollbild.
- **Formatierung**: Fett, Kursiv, Schriftgrad vergrößern/verkleinern; der **Stilbereich** für das ausgewählte Element (Farben und mehr).
- **Einfügen**: Überschrift, Absatz, Tabelle, Bild (per Link), weitere.
- **Bildaktionen** (wenn ein Bild ausgewählt ist): Zuschneiden, **Hintergrund entfernen**, ersetzen, Seitenverhältnis sperren.
- **Elementaktionen** (wenn im Vorschau-Inspektor ein Element ausgewählt ist): löschen, duplizieren, nach oben/nach unten verschieben.
- **KI-Schaltfläche**: öffnet den KI-Bereich; fragen Sie direkt zum ausgewählten Element.
