# Collegare un agente di programmazione

GenOffice parla il Model Context Protocol, così un agente di programmazione può leggere, scrivere e renderizzare i tuoi documenti attraverso gli stessi motori che usa l'app. L'agente non sta indovinando un formato di file: ottiene gli schemi tipizzati delle op dalle stesse definizioni contro cui l'executor valida.

## Registrarlo

Il caso più comune è un solo comando:

```sh
genoffice mcp install all
```

Trova gli agenti di programmazione su questa macchina — Claude Code, Codex, Cursor, Gemini CLI, Copilot CLI, OpenCode, Windsurf — e scrive la voce del server stdio nella configurazione di ciascuno, lasciando intatto tutto il resto del file.

```sh
genoffice mcp list             # where each agent stands
genoffice mcp install cursor    # just one
genoffice mcp uninstall cursor  # take it back out
```

Un agente installato in una posizione non convenzionale richiede `--dir <path>`; `--force` riscrive una voce già presente.

## Eseguirlo da soli

Per un client su un'altra macchina, pubblicalo via HTTP:

```sh
genoffice mcp --http 8765 --token <secret>
```

`--host <addr>` cambia l'indirizzo su cui è in ascolto. Passa lo stesso token al client.

Anche i file viaggiano via HTTP: `PUT /files/<name>` ne carica uno, ogni strumento accetta un URL `http(s)` al posto di un percorso e gli output tornano come URL di download — e, quando sono abbastanza piccoli, come risorse incorporate. In ogni caso, op, spec e Markdown vengono passati in linea.

## Cosa ottiene l'agente

Ogni comando è uno strumento. Quelli più interessanti:

- **`docs`, `sheet`, `slides`** — leggono e modificano un file attraverso il percorso di scrittura dell'app, una **op** alla volta. Un nuovo deck segue `deck_start`, `deck_page`, `deck_build`.
- **`render`** — un PNG per pagina, impaginato dal renderer dell'app, così l'agente può guardare una slide invece di indovinarla.
- **`pdf`** — il livello di testo di un PDF, pagina per pagina, senza processi dell'app.
- **`info`** — metadati e un riepilogo della struttura, di solito la chiamata più economica su un file non noto.
- **`search`, `image`, `media`** — i provider configurati nell'app, così l'agente non ha bisogno delle proprie chiavi.
- **`merge`** — compila un modello `{{key}}`.

## Schemi e un budget più piccolo

`apply` e `create` annunciano i loro parametri `ops`, `cells` e `data` con lo schema tipizzato per ogni op, generato da `genoffice guide <domain> --json`. È preciso e non è piccolo. Un client con una finestra di contesto ridotta può chiedere invece semplici array:

```sh
genoffice mcp --compact-schemas
GENOFFICE_MCP_COMPACT_SCHEMAS=1 genoffice mcp
```

## La skill

Un agente che non conosce il vocabolario delle op indovinerà. `genoffice skill` installa una skill GenOffice negli agenti che trova, portando con sé il riferimento e le guide di progettazione — lo stesso materiale che stampa `genoffice guide`.

## Che cos'è

Il server MCP legge e scrive file. Non è la finestra: non c'è il pannello IA e la finestra di aggiornamento dell'app non si applica. Se un passaggio richiede la finestra, apri il file.
