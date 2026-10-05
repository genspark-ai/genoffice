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
