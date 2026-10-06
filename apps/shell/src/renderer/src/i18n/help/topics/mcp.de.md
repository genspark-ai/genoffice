# Einen Coding-Agenten anbinden

GenOffice spricht das Model Context Protocol: Ein Coding-Agent kann Ihre Dokumente also über dieselben Engines lesen, schreiben und rendern wie die App. Der Agent rät nicht an einem Dateiformat — er erhält die typisierten op-Schemas aus genau den Definitionen, gegen die der Executor prüft.

## Einrichtung

Der übliche Fall ist ein Befehl:

```sh
genoffice mcp install all
```

Er findet die Coding-Agenten auf diesem Rechner — Claude Code, Codex, Cursor, Gemini CLI, Copilot CLI, OpenCode, Windsurf — und schreibt den Eintrag für den stdio-Server in die Konfiguration jedes einzelnen; den Rest der Datei lässt er, wie er ihn vorgefunden hat.

```sh
genoffice mcp list             # where each agent stands
genoffice mcp install cursor    # just one
genoffice mcp uninstall cursor  # take it back out
```

Einen Agent, den Sie an einem ungewöhnlichen Ort installiert haben, erreicht man mit `--dir <path>`; `--force` überschreibt einen Eintrag, der schon vorhanden ist.

## Selbst ausführen

Für einen Client auf einem anderen Rechner: Stattdessen über HTTP bereitstellen:

```sh
genoffice mcp --http 8765 --token <secret>
```

`--host <addr>` legt fest, wo er lauscht. Geben Sie demselben Client dasselbe Token.

Auch über HTTP wandern Dateien: `PUT /files/<name>` lädt eine hoch, jedes Werkzeug nimmt anstelle eines Pfads eine `http(s)`-URL entgegen, und Ausgaben kommen als Download-URLs zurück — und, wenn sie klein genug sind, als eingebettete Ressourcen. Ops, Specs und Markdown werden in beiden Fällen inline übergeben.

## Was der Agent bekommt

Jeder Befehl ist ein Werkzeug. Die interessanten:

- **`docs`, `sheet`, `slides`** — eine Datei über den Schreibpfad der App lesen und bearbeiten, einen **op** nach dem anderen. Eine neue Präsentation läuft über `deck_start`, `deck_page`, `deck_build`.
- **`render`** — eine PNG-Datei pro Seite, vom Renderer der App gesetzt, sodass der Agent eine Folie ansieht, statt zu raten.
- **`pdf`** — die Textebene einer PDF-Datei, Seite für Seite, ohne App-Prozess.
- **`info`** — Metadaten und eine Strukturübersicht, meist der billigste erste Aufruf bei einer unbekannten Datei.
- **`search`, `image`, `media`** — die in der App konfigurierten Anbieter, sodass der Agent keine eigenen Schlüssel braucht.
- **`merge`** — eine `{{key}}`-Vorlage füllen.

## Schemas und ein kleineres Budget

`apply` und `create` geben ihre Parameter `ops`, `cells` und `data` mit dem typisierten Schema pro op bekannt, erzeugt aus `genoffice guide <domain> --json`. Das ist präzise und nicht klein. Ein Client mit knappem Kontextfenster kann stattdessen einfache Arrays anfordern:

```sh
genoffice mcp --compact-schemas
GENOFFICE_MCP_COMPACT_SCHEMAS=1 genoffice mcp
```

## Der Skill

Ein Agent, der das op-Vokabular nicht kennt, wird raten. `genoffice skill` installiert einen GenOffice-Skill in die gefundenen Agenten, mit Referenz und Design-Leitfäden — demselben Material, das `genoffice guide` ausgibt.

## Was es nicht ist

Der MCP-Server liest und schreibt Dateien. Er ist nicht das Fenster: Es gibt keinen KI-Bereich, und der Aktualisierungsdialog in der App gilt hierfür nicht. Wenn ein Schritt das Fenster braucht, öffnen Sie die Datei.
