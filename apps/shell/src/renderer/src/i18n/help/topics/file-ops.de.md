# Dateivorgänge: umbenennen, löschen, exportieren

Dieses Kapitel behandelt die Dateivorgänge, die allen Editoren gemeinsam sind; die eigenen Exportoptionen jedes Editors stehen in seinem Kapitel.

Das **⋯-Menü** einer Zeile (fahren Sie über eine Dateizeile) bündelt diese Aktionen:

![Das ⋯-Menü einer Dateizeile](img/file-ops.png)

## Umbenennen

Zwei Einstiegspunkte, dieselben Prüfungen:

- **⋯ ▸ Umbenennen** auf einer Zeile der Startseite.
- **Doppelklicken Sie auf einen Dateitab**, um direkt dort umzubenennen (siehe [Tabs und Fensterverwaltung](help://tabs-and-windows)).

Regeln: Die Dateiendung bleibt automatisch erhalten; unzulässige Zeichen, Punkte am Ende und reservierte Namen (CON/NUL und Verwandte) werden mit einer Meldung abgelehnt; auch ein Namenskonflikt im selben Ordner wird verhindert. Die echte Datei auf dem Datenträger wird umbenannt, und Zuletzt verwendet und Favoriten folgen.

## Löschen

- **⋯ ▸ Löschen** auf der Startseite: fragt ab, welche Dateien, und verschiebt sie in den **Papierkorb des Systems**, von dort über das Betriebssystem wiederherstellbar. GenOffice führt dafür kein eigenes Rückgängig — die Wiederherstellung ist Sache des Papierkorbs, nicht einer Meldung.

## Duplizieren

**⋯ ▸ Duplizieren** erstellt im selben Ordner eine Kopie namens <name> copy; bei Namensgleichheit wird automatisch ein Zähler angehängt. Die Kopie landet in **Zuletzt verwendet**, statt geöffnet zu werden — sie ist einen Klick entfernt, nicht vor Ihnen aufgeschlagen.

## Speichern und Speichern unter

- **⌘S / ctrl+S** speichert die aktuelle Datei; bei einer Datei ohne Namen werden zuerst Ort und Name abgefragt.
- **Speichern unter** schreibt eine neue Datei und lässt das Original unberührt; weitere Änderungen gelten dann der neuen Datei.
- Jedes Speichern erfolgt atomar (temporäre Datei plus Umbenennung); ein Abbruch mitten im Schreiben kann die Datei nicht beschädigen.
- AutoSpeichern greift erst nach dem ersten manuellen Speichern (siehe [Schnellstart](help://getting-started)).

## Als PDF exportieren

- **Docs**: Datei ▸ Als PDF exportieren (oder die Schaltfläche im Menüband), paginiert nach dem aktuellen Seitenumbruch.
- **Slides**: Der Export rasternt Seite für Seite, mit Fortschritt bei großen Decks.
- **Sheets**: Der Export folgt dem Druckumbruch.
- Exporte werden in einem verborgenen Fenster gerendert und landen dort, wo Sie es wählen.

## Als Word / Bilder exportieren

- **PDF ▸ PDF in Word**: wandelt die PDF in .docx um (lokale Konvertierung; komplexe Layouts werden nach bestem Bemühen übernommen).
- **Docs** kann Seiten als Bilder exportieren (eine PNG-Datei pro Seite).

## Drucken

Datei ▸ Drucken in jedem Editor (⌘P/ctrl+P) öffnet den Systemdialog; PDFs werden mit der aktuellen Seitenreihenfolge und den aktuellen Drehungen gedruckt.

## Wo Dateien ohne Namen liegen

Der Ort, den Sie beim ersten Speichern wählen, ist ihr Zuhause; davor existiert das Dokument nur im Speicher. AutoSpeichern übernimmt erst nach diesem ersten Speichern.
