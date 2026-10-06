# Riga di comando e agenti

Ogni installazione include il comando `genoffice` che usa gli stessi motori della finestra — gli stessi parser, lo stesso writer, lo stesso renderer. Un file che l'app salva e un file che il comando scrive sono lo stesso file, e un controllo che il pannello IA dell'app supera è un controllo che il comando supera.

```sh
genoffice --help          # every command
genoffice guide slides    # the op reference, for writing your own
```

## Ottenere il comando

macOS e Windows lo includono dentro il bundle dell'app. Per usarlo per nome, esegui una volta `genoffice install-cli`: crea un collegamento simbolico del binario incluso in `/usr/local/bin`, o nel `PATH` dell'utente su Windows.

## I comandi da conoscere

| Comando | Cosa fa |
| --- | --- |
| `open` | Apre un documento nell'app; avvia l'app se non è in esecuzione. |
| `convert` | Converte tra formati usando i motori dell'app. |
| `create` | Crea un documento da contenuto strutturato. |
| `render` | Un PNG per pagina, così come il renderer impagina. |
| `pdf` | Legge il livello di testo di un PDF pagina per pagina, senza processi dell'app. |
| `info` | Metadati e un riepilogo della struttura di un documento. |
| `search` | Ricerca web o di immagini tramite il provider configurato nell'app. |
| `image` / `media` | Genera un'immagine, oppure descrivi e fai domande su un file immagine, video o audio. |
| `merge` | Compila i segnaposto `{{key}}` in un modello `.docx`, `.pptx` o `.xlsx`. |
| `capabilities` | Riporta quali funzioni cloud sono configurate su questa macchina. |
| `guide` | Il riferimento delle op e le guide di progettazione, generati dalle stesse definizioni contro cui l'executor valida — così non può divergere da ciò che `apply` accetta. `--json` lo restituisce con lo schema di ogni op. |
| `install-cli` | Mette `genoffice` sul `PATH`. |
| `skill` | Elenca gli agenti di programmazione trovati su questa macchina e installa o aggiorna in essi la skill GenOffice. |
| `mcp` | Espone ogni comando come strumento Model Context Protocol. Vedi **Collegare un agente di programmazione**. |

## Modifica: Docs, Sheets, Slides

`genoffice docs`, `genoffice sheet` e `genoffice slides` leggono e modificano attraverso lo stesso percorso di scrittura che usa l'app, e condividono un vocabolario: un'**op** è una singola modifica e una **spec** è un elenco di op applicate in ordine.

```sh
genoffice sheet read report.xlsx
genoffice slides apply deck.pptx --ops ops.json --dry-run
genoffice docs apply brief.docx --ops ops.json
```

`--dry-run` riporta cosa farebbe un batch e non scrive nulla, il modo più rapido per controllare una spec prima di applicarla. Il pannello IA dell'app funziona esattamente su queste op, quindi tutto ciò che puoi chiedergi puoi anche scriptarlo.

## Model Context Protocol

`genoffice mcp` espone ogni comando come strumento MCP e `genoffice mcp install <agent|all>` lo registra nella configurazione di un agente di programmazione. Per quel lato, vedi **Collegare un agente di programmazione**.

## Che cosa non fa il comando

Legge e scrive il file. Non è l'app: non c'è finestra e la finestra di aggiornamento dell'app non si applica. Tutto ciò che richiede la finestra — il pannello IA, il controllo di qualità su una slide renderizzata — deve aspettare che tu apra il file. `genoffice render` ti dà i pixel senza di essa e `genoffice slides` verifica il layout di un deck da solo.
