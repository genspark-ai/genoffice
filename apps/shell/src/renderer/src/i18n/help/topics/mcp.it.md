# Collegare un agente di programmazione

GenOffice parla il Model Context Protocol, così un agente di programmazione può leggere, scrivere e renderizzare i tuoi documenti attraverso gli stessi motori che usa l'app. L'agente non sta indovinando un formato di file: ottiene gli schemi tipizzati delle op dalle stesse definizioni contro cui l'executor valida.

## Registrarlo dall'app

È in **Impostazioni ▸ Integrazioni** che si fa questa cosa. Il pannello ha due metà, e puoi usare una sola o entrambe.

**La skill.** Una riga per ogni agente di programmazione trovato su questa macchina — Claude Code, Codex, Cursor, Gemini CLI, Copilot CLI, OpenCode, Windsurf —, ognuna con **Installa**, **Aggiorna** e **Disinstalla**, più **Installa in un'altra cartella…**, **Scarica skill (zip)** e **Copia percorso**. Se il tuo assistente non è in elenco, indica a GenOffice una cartella da cui legge `SKILL.md`, oppure salva lo zip e lascia che lo installi l'assistente. Skill e MCP possono stare fianco a fianco: l'assistente sceglie uno dei due, e fanno esattamente le stesse cose.

**MCP.** Sono offerti due percorsi: **Avviato dall'assistente (consigliato)**, dove aggiungi la configurazione mostrata al tuo client e l'assistente avvia da solo il server, e il **Server HTTP locale**, che l'app esegue al posto tuo. In entrambi i casi l'assistente finisce per parlare con GenOffice e tu non digiti mai un comando.

## Registrarlo dalla riga di comando

La stessa cosa da un terminale: questo è il percorso avanzato, quello da usare quando l'agente si trova dove il pannello non può cercarlo:

```sh
genoffice mcp install all
```

Trova gli agenti di programmazione su questa macchina e scrive la voce del server stdio nella configurazione di ciascuno, lasciando intatto tutto il resto del file.

```sh
genoffice mcp list             # where each agent stands
genoffice mcp install cursor    # just one
genoffice mcp uninstall cursor  # take it back out
```

Un agente installato in una posizione non convenzionale richiede `--dir <path>`; `--force` riscrive una voce già presente.

Tutto ciò che il server accetta sta su una sola schermata: le forme per installare, rimuovere e elencare, l'erogazione via HTTP e le due opzioni sugli schemi:

![L'output reale di genoffice mcp --help: le forme install, uninstall e list, con le opzioni --http, --host, --token, --compact-schemas, --dir e --force](img/mcp.png)

## Eseguirlo senza un assistente

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

## Perché la skill

Un agente che non conosce il vocabolario delle op indovinerà. La skill porta con sé il riferimento e le guide di progettazione — lo stesso materiale che stampa `genoffice guide` —, così l'assistente scrive op di cui ha davvero letto la specifica. Installala dal pannello qui sopra, oppure con `genoffice skill` da un terminale.

## Che cosa raggiunge nell'app

Il server non è limitato ai file su disco. Finché GenOffice è in esecuzione, l'agente può lavorare anche attraverso la finestra:

- **`open_in_genoffice`** apre un file in una scheda e la porta in primo piano.
- **`open_documents`** elenca ogni documento che hai aperto — id, tipo, percorso e se ha modifiche non salvate — e poi legge il contenuto attuale di uno o lo chiude, salvando prima a meno che tu non gli chieda di scartare.
- **Gli strumenti di contenuto** prendono quell'id (o il percorso) come argomento `document`, così una modifica finisce nella scheda che hai già aperto e la finestra passa a mostrarla.

Restano fuori portata due cose: non c'è il pannello IA e la finestra di aggiornamento dell'app non si applica.

## Il pannello Server HTTP locale

Sotto **Server HTTP locale**, l'app esegue il server per conto suo invece di lasciarlo all'assistente: un interruttore di attivazione, un campo **Porta**, un indicatore **In esecuzione / Non in esecuzione** e la **Generazione in background** (scrivere i documenti direttamente in un percorso senza aprire l'interfaccia) e l'**Esempio di configurazione client** da copiare. Aprire **Avanzato** aggiunge i due URL di connessione — Streamable HTTP e quello vecchio di SSE —, un URL di **Controllo di integrità** e un interruttore di **Registrazione** che annota l'attività del server e degli strumenti in un file locale che puoi **Aprire**, **Aggiornare** o **Svuotare** da lì. Ascolta solo su localhost.
