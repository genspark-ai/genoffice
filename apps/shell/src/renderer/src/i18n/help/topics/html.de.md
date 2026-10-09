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

## Exportieren

Menü Datei, alles lokal und alles fragt nach, wohin das Ergebnis soll:

- **Als Word exportieren…** und **Als PDF exportieren…** schreiben eine echte .docx bzw. .pdf.
- **Als Einzeldatei-HTML exportieren…** schreibt eine einzelne .html mit eingebetteten Bildern. Sie überschreibt nicht die Datei, die Sie gerade geöffnet haben, und meldet, wie viele Bilder sie nicht einbetten konnte.

## Gerüst einfügen

Für eine leere Seite schreibt **Einfügen ▸ Gerüst einfügen** ein minimales Dokument im Standardmodus:

```html
<!DOCTYPE html>
<html lang="en">
  <head>
    <meta charset="UTF-8" />
    <title></title>
  </head>
  <body></body>
</html>
```

Jeder Teil hat seinen Grund, weshalb das ein Befehl ist und nicht etwas, das Sie tippen müssten:

- den **doctype**, sonst läuft die Vorschau im Quirks-Modus, in dem Boxgrößen und Tabellenlayout anderen Regeln folgen als erwartet;
- das **`lang`**, sonst hat ein Screenreader keine Sprache, um die Seite vorzulesen, und der Browser wählt Schriftart und Rechtschreibprüfung für die falsche Sprache;
- den **charset**, sonst kann eine Seite mit nicht lateinischer Schrift als Zeichenbrei dekodiert werden.

Die Viewport-Meta-Angabe fehlt bewusst: Sie wird in einer Desktop-Fläche gerendert, auf die kein mobiler Viewport wirkt.

Das `lang` folgt der Oberflächensprache der App; das eingefügte Gerüst ist also das, für das Ihre Werkzeuge bereits eingerichtet sind. Bearbeiten Sie es danach frei.

Der Eintrag erscheint nur im Bearbeitungsmodus und nur, solange das Dokument leer ist — sobald Inhalt da ist, gibt es nichts mehr, _worein_ ein Gerüst einzufügen wäre.
