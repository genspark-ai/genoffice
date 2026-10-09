# Der Markdown-Editor

Der Markdown-Editor öffnet .md / .markdown mit Quelltext und gerenderter Vorschau nebeneinander.

- **Öffnen**: über die Startseite oder Datei ▸ Öffnen; auch die Kommandozeile funktioniert.
- **Bearbeiten**: Bearbeitung als reiner Text; GFM-Erweiterungen (Tabellen, Aufgabenlisten, Durchgestrichen, automatische Links) werden in der Vorschau gerendert.
- **Vorschau**: in Echtzeit; relative Ressourcen wie Bilder werden neben dem Dokument aufgelöst.
- **Speichern**: bytegetreu — BOM, CRLF und das Vorhandensein eines abschließenden Zeilenumbruchs bleiben erhalten; ein Speichern ohne Änderung schreibt die Datei nicht neu.
- **Suchen und Ersetzen**: Strg+F durchsucht den Quelltext; „Alle ersetzen“ schreibt zurück.
- **KI**: Voreinstellungs-Schaltflächen lassen den Assistenten das Dokument umformulieren, ausbauen oder übersetzen.

## Die Werkzeugleiste

Eine Zeile Schaltflächen über dem Editor (zum Tooltip mit der Maus darüberfahren):

![Die Markdown-Werkzeugleiste](img/md-toolbar.png)

- **Datei und Verlauf**: Speichern, Speichern unter, Rückgängig, Wiederholen, Suchen; der Schalter **AutoSpeichern** rechts schreibt Änderungen in regelmäßigen Abständen auf den Datenträger.
- **KI-Schaltfläche**: öffnet den KI-Bereich; daneben liegen die Voreinstellungen zum Umformulieren, Ausbauen und Übersetzen.
- **Absatzformat** (Dropdown): zwischen Fließtext und den Überschriftebenen wechseln.
- **Zeichenformat**: **Fett**, _Kursiv_, ~~Durchgestrichen~~, `Inline-Code`, Link.
- **Listen**: Aufzählung, nummerierte Liste, Aufgabenliste.
- **Einfügen**: Tabelle, Bild, Trennlinie.
- **Eigenschaften**: fügt den YAML-Front-Matter-Block am Dateianfang ein oder springt dorthin.
- **Gliederung**: springt nach Überschriftenhierarchie.
- **Rechtschreibung**: schaltet die Rechtschreibprüfung für dieses Dokument um.

Drei schnelle Beispiele:

- **Überschrift**: Cursor in die Zeile setzen ▸ Dropdown „Absatzformat“ ▸ „Überschrift 1“.
- **Tabelle**: auf **Tabelle einfügen** klicken ▸ Zeilen- und Spaltenzahl ziehen ▸ Zellen ausfüllen; die Vorschau rendert sie sofort.
- **Aufgabenliste**: ein paar Zeilen auswählen ▸ auf **Aufgabenliste** klicken ▸ jede Zeile wird zu `- [ ]` und in der Vorschau als Liste mit Kontrollkästchen dargestellt.

## Exportieren

Menü Datei, alles lokal und alles fragt nach, wohin das Ergebnis soll:

- **Als Word exportieren…** und **Als PDF exportieren…** schreiben eine echte .docx bzw. .pdf.
- **Als Bilder exportieren…** schreibt ein PNG pro Seite in einen Ordner, den Sie auswählen.
- **In Docs umwandeln und öffnen** konvertiert nach .docx und öffnet sie im integrierten Docs-Tab hier in der App — es ist keine Übergabe an etwas in der Cloud, und die konvertierte Kopie liegt in einem Cache-Ordner, der nach etwa einer Woche aufgeräumt wird.

## Quelltextansicht

Die Werkzeugleiste trägt einen Schalter **Quelltext** (wie die App lokalisiert). Schalten Sie ihn ein, und der Editor wird durch das rohe Markdown ersetzt: genau der Text, den ein Speichern schreibt, nichts hübsch gemacht, nichts unter der Hand normalisiert.

- **Bearbeiten ist bytegetreu.** Ein Speichern aus der Quelltextansicht erzeugt dieselben Bytes wie ein Speichern aus dem Editor — BOM, CRLF und ein abschließender Zeilenumbruch bleiben alle erhalten.
- **Es ist dasselbe Dokument.** Wechseln Sie frei hin und her; der Quelltext ist der Text des Editors selbst, keine Kopie, die zusammengeführt werden müsste.
- **Die Formatierungsleiste steht nicht zur Verfügung**, solange die Ansicht offen ist, weil die meisten dieser Schaltflächen Editor-Strukturen einfügen, die auf der gerenderten Seite erst einen Sinn ergeben. Sie ist wieder da, wenn Sie die Ansicht schließen.
- **JSON und andere Dateien im Quelltextmodus** öffnen sich hier direkt: Es gibt nichts zu rendern, also _ist_ der Quelltext das Dokument.
