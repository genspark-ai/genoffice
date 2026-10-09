# Kommandozeile und Agenten

Jede Installation bringt den Befehl `genoffice` mit, der dieselben Engines antreibt wie das Fenster: dieselben Parser, denselben Writer, denselben Renderer. Eine Datei, die die App speichert, und eine Datei, die der Befehl schreibt, sind dieselbe Datei, und eine Prüfung, die der KI-Bereich der App besteht, besteht auch der Befehl.

```sh
genoffice --help          # every command
genoffice guide slides    # the op reference, for writing your own
```

Das ist die gesamte Oberfläche auf einem Bildschirm — jeder Befehl mit einer Zeile dazu, was er bewirkt, dann die globalen Optionen und die Exit-Codes:

![Die echte Ausgabe von genoffice --help: die Versionsbanner, jeder Befehl mit einzeiliger Beschreibung sowie die globalen Optionen und die Exit-Codes](img/cli.png)

## Den Befehl bereitstellen

Unter macOS und Windows liefert die App ihn im Bundle mit. Um ihn beim Namen aufzurufen, führen Sie einmal `genoffice install-cli` aus: Das legt einen symbolischen Link zum mitgelieferten Binary in `/usr/local/bin` an, unter Windows im Benutzer-`PATH`.

## Die Befehle, die man kennen sollte

| Befehl            | Was er bewirkt                                                                                                                                                                                                                                                        |
| ----------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `open`            | Ein Dokument in der App öffnen; startet die App, wenn sie nicht läuft.                                                                                                                                                                                                |
| `selection`       | Was der Benutzer in dieser Datei gerade ausgewählt hat, solange die App läuft — der Zeiger auf „dieser“ / „hier“. Gibt einen Blockbereich, einen Bereich im Tabellenblatt, Folienelemente oder eine Seite zurück, je nachdem, welchem Editor die Datei gerade gehört. |
| `convert`         | Zwischen Formaten mit den Engines der App konvertieren.                                                                                                                                                                                                               |
| `create`          | Ein Dokument aus strukturiertem Inhalt erstellen.                                                                                                                                                                                                                     |
| `render`          | Eine PNG-Datei pro Seite, so wie der Renderer sie setzt.                                                                                                                                                                                                              |
| `pdf`             | Die Textebene einer PDF-Datei Seite für Seite lesen, ohne einen App-Prozess zu starten.                                                                                                                                                                               |
| `info`            | Metadaten und eine Strukturübersicht eines Dokuments.                                                                                                                                                                                                                 |
| `search`          | Web- oder Bildsuche über den in der App konfigurierten Anbieter.                                                                                                                                                                                                      |
| `image` / `media` | Ein Bild erzeugen oder eine Bild-, Video- oder Audiodatei beschreiben und Fragen dazu stellen.                                                                                                                                                                        |
| `merge`           | Die Platzhalter `{{key}}` in einer Vorlage `.docx`, `.pptx` oder `.xlsx` füllen.                                                                                                                                                                                      |
| `capabilities`    | Melden, welche Cloud-Funktionen auf diesem Rechner konfiguriert sind.                                                                                                                                                                                                 |
| `guide`           | Die op-Referenz und die Design-Leitfäden, erzeugt aus genau den Definitionen, gegen die der Executor prüft — sie kann also nicht von dem abweichen, was `apply` akzeptiert. `--json` liefert sie mit dem Schema jedes ops.                                            |
| `install-cli`     | `genoffice` in den `PATH` legen.                                                                                                                                                                                                                                      |
| `skill`           | Die auf diesem Rechner gefundenen Coding-Agenten auflisten und den GenOffice-Skill dort installieren oder aktualisieren.                                                                                                                                              |
| `mcp`             | Jeden Befehl als Model-Context-Protocol-Werkzeug bereitstellen. Siehe **Einen Coding-Agenten anbinden**.                                                                                                                                                              |

## Bearbeiten: Dokumente, Tabellen, Folien

`genoffice docs`, `genoffice sheet` und `genoffice slides` lesen und schreiben über denselben Pfad wie die App und teilen ein Vokabular: ein **op** ist eine einzelne Änderung, und eine **spec** ist eine Liste von ops, die der Reihe nach angewendet werden.

```sh
genoffice sheet read report.xlsx
genoffice slides apply deck.pptx --ops ops.json --dry-run
genoffice docs apply brief.docx --ops ops.json
```

`--dry-run` meldet, was ein Stapel tun würde, und schreibt nichts — so prüfen Sie eine spec günstig, bevor sie landet. Der KI-Bereich der App läuft auf genau diesen ops, Sie können also alles, was Sie ihm anvertrauen können, auch skripten.

## Model Context Protocol

`genoffice mcp` stellt jeden Befehl als MCP-Werkzeug bereit, und `genoffice mcp install <agent|all>` registriert es in der Konfiguration eines Coding-Agenten. Diese Seite steht in **Einen Coding-Agenten anbinden**.

## Was der Befehl nicht tut

Er liest und schreibt die Datei. Er ist nicht die App: es gibt kein Fenster, und der Aktualisierungsdialog in der App gilt hierfür nicht. Alles, was das Fenster braucht — der KI-Bereich, die QC-Prüfung einer gerenderten Folie —, wartet darauf, dass Sie die Datei öffnen. `genoffice render` liefert Ihnen die Pixel auch ohne Fenster, und `genoffice slides` prüft das Layout eines Decks eigenständig.
