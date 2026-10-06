# Docs: Textverarbeitung

Docs ist der Word-ähnliche Textprozessor: Er liest und schreibt echte .docx mit treuem WYSIWYG-Seitenumbruch.

## Das Menüband

Tabs: **Start / Einfügen / Layout / Entwurf / Referenzen / Überprüfen / Ansicht**, dazu kontextbezogene Tabs für das ausgewählte Objekt (Tabellenentwurf, Bilder).

- **Start**: Zwischenablage; Schriftart (einschließlich ostasiatischer Schriftgrade und Hervorhebungszeichen); Absatz (Ausrichtung/Einzug/Abstand/Listen); Formatvorlagen (Überschrift 1-6/Standard/Zitat, änderbar).
- **Einfügen**: Seiten- und Abschnittswechsel, Tabellen (einschließlich gezeichneter und Schnelltabellen), Bilder, Formen, Hyperlinks, Kopf- und Fußzeile, Seitenzahl, Datum, Textfelder.
- **Layout**: Seitenränder, Ausrichtung und Papierformat, Spalten, Absatzeinzüge und Abstände.
- **Entwurf**: Designs, Farbsätze, Wasserzeichen, Seitenrahmen.
- **Referenzen**: Inhaltsverzeichnis (aktualisierbar), Fuß- und Endnoten, Beschriftungen, Querverweise.

  ![Der Tab Referenzen](img/docs-references.png)

- **Überprüfen**: Rechtschreibprüfung, Kommentare, Änderungen nachverfolgen (Ansichten „Alle Markups“ / „Einfaches Markup“), Wörter zählen.
- **Ansicht**: Lineal, Gitternetzlinien, Navigationsbereich, Zoom und der durchsuchbare Dialog für **Tastenkombinationen**.

## Der Navigationsbereich

**Ansicht ▸ Navigationsbereich** öffnet einen seitlichen Bereich mit der Überschriftengliederung des Dokuments, einem Suchfeld für das ganze Dokument und einer Miniaturansicht pro Seite. Ob er geöffnet ist, bleibt zwischen den Starts erhalten — ein Dokument, das Sie über die Gliederung erschließen, bleibt also erschließbar.

**Die Gliederung** ist der Überschriftenbaum. Mit der rechten Maustaste auf eine Überschrift darin klicken, um sie einzuklappen oder umzustrukturieren statt nur zu navigieren:

- **Reduzieren / Erweitern** an einer Überschrift klappt deren gesamten Teilbaum ein bzw. aus — das Kapitel verschwindet, sein Text bleibt im Dokument.
- **Alle reduzieren / Alle erweitern** klappt alles auf einmal ein bzw. aus. Bei einem langen Bericht ist das der Unterschied zwischen einer lesbaren Gliederung und einer Wand aus Text.
- **Überschriftenebenen anzeigen** filtert den Baum auf die Ebenen, die Sie interessieren, sodass *Überschrift 1 anzeigen* Ihnen ein Inhaltsverzeichnis übrig lässt, das Sie wirklich überfliegen können.
- **Höher stufen / Tiefer stufen** ändern die Ebene der Überschrift und damit auch die Ebene, die alle Überschriften darunter erben — so wie aus einem Kapitel ein Abschnitt wird.
- **Neue Überschrift davor / danach** fügt eine an der Cursorposition ein, ohne den Bereich zu verlassen.
- **Löschen** entfernt die Überschrift *und alles darunter* — darauf ist zu achten: Es löscht einen Teilbaum, keine Zeile.
- **Überschrift und Inhalt auswählen** markiert von der Überschrift bis zum Ende ihres Teilbaums, bereit für eine Bearbeitung des ganzen Abschnitts.

## Rechtsklick-Menü

Rechtsklicken Sie irgendwo im Text — das Menü richtet sich nach dem, was Sie angeklickt haben. Wichtigste Gruppen:

- **Zwischenablage**: Ausschneiden / Kopieren / Einfügen / **Unformatierter Unicode-Text einfügen**.
- **Schriftart, Absatz**: Schriftart und -größe, Fett/Kursiv/Unterstrichen, Ausrichtung/Einzug/Abstand ändern, ohne zum Menüband zu gehen.
- **Synonyme**: listet Synonyme zum ausgewählten Wort auf; ein Klick ersetzt es.
- **Übersetzen** (KI): übersetzt die Auswahl über den KI-Bereich in die Zielsprache (Englisch, vereinfachtes Chinesisch, Japanisch, Koreanisch, Französisch, Deutsch, Spanisch, …).
- **Neuer Kommentar**: hängt einen Kommentar an die Auswahl an.
- **Rechtschreibung** (auf einem falsch geschriebenen Wort): Ersetzungsvorschläge, alle ignorieren, zum Wörterbuch hinzufügen, Rechtschreibsprache festlegen.
- **Hyperlink**: öffnen / bearbeiten / Link kopieren / Hyperlink entfernen.
- **Bild**: Bild anzeigen, Bild speichern unter, **Textumbruch** (im Text / links und rechts / oben und unten / hinter dem Text / vor dem Text), Anordnungsreihenfolge.
- **Felder** (Inhaltsverzeichnis, Seitenzahlen): Feld aktualisieren / Feldfunktionen ein-/ausblenden / Feld bearbeiten.
- **Nummerierung** (in einer Liste): Nummerierung neu beginnen / Nummerierung fortsetzen / Listenebene ändern / Nummerierungswert festlegen.
- **Tabelle** (Cursor in einer Tabelle): Zeilen/Spalten einfügen, Zellen verbinden / teilen, Tabelle teilen, automatisch anpassen, Zellenausrichtung, Zeilen/Spalten verteilen, Tabelleneigenschaften, Lösch-Menü, Auswählen.

## Bearbeiten

- Suchen und Ersetzen (ctrl+F / ctrl+H): Groß-/Kleinschreibung, ganzes Wort, reguläre Ausdrücke.
- Format übertragen; mehrstufiges Rückgängig/Wiederholen; Einfügeoptionen.
- Tabellen: Zellen verbinden/teilen, Zeilen-/Spaltenbefehle, Rahmen und Schattierung, Sortieren, Formeln.
- Bilder: Textumbruch, Zuschneiden, Komprimieren; Zeichenfläche.

## Ostasiatische Typografie

- Zeichensatzkompression und Kinsoku-Zeilenumbrüche entsprechen Word; Umwandlung zwischen voller und halber Breite.
- Die Schriftkandidaten decken die gängigen ostasiatischen Familiennamen unter Windows und macOS ab.

## KI

- KI-Schaltfläche im Menüband und Seitenbereich: umformulieren, ausbauen, übersetzen, zusammenfassen, Tabellen-Voreinstellungen einfügen sowie freie Anweisungen.
- Jeder KI-Durchgang legt zuerst einen Snapshot an; Sie können aus der Versionsliste zurückrollen, und auch das Zurückrollen lässt sich rückgängig machen.

## Speichern und exportieren

- Speichert .docx und schreibt nur geänderte Absätze neu — unveränderter Inhalt bleibt byteidentisch.
- Exportiert PDF (mit aktuellem Seitenumbruch) und Bilder Seite für Seite.

## Drucken

ctrl+P über den Systemdialog, WYSIWYG-Seiten.
