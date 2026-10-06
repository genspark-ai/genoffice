# Impostazioni, lingua, tema e integrazioni MCP

## Apertura delle impostazioni

La riga dell'account in basso a sinistra in Home apre il pannello delle impostazioni (quando non hai effettuato l'accesso riporta Accedi). Ha sei sezioni: Account, Modello IA, Media e ricerca IA, Generale, Integrazioni e Informazioni.

![Impostazioni ▸ Generale, dove si trovano la lingua, il tema, il salvataggio automatico e l'interruttore delle statistiche di utilizzo](img/settings-general.png)

La configurazione dei modelli ha un articolo tutto suo; in **Media e ricerca IA** attivi, per ogni provider, generazione di immagini, analisi delle immagini, analisi video, ricerca web e ricerca file locali.

## Lingua

- Le impostazioni offrono **21 lingue per l'interfaccia**: inglese, cinese semplificato, giapponese, coreano, francese, tedesco, spagnolo, thai, indonesiano, russo, arabo, portoghese, italiano, polacco, ceco, olandese, malese, ebraico, hindi, cinese tradizionale, vietnamita.
- Il cambio ha effetto immediato e viene ricordato; la barra dei menu nativa viene ricostruita nella nuova lingua.

## Tema

Chiaro / Scuro / Segui il sistema. L'opzione Segui il sistema segue l'aspetto del sistema operativo e gli editor si adattano in sincronia senza sfarfallii.

## Generale

- **Invia statistiche di utilizzo anonime** — attivo per impostazione predefinita. Utilizza Google Analytics 4 e invia il tuo IP pubblico e i metadati di trasporto; i contenuti dei documenti e i nomi dei file non vengono mai raccolti e ogni evento porta solo un tipo come «aperto un .docx». Puoi disattivarlo qui in qualsiasi momento.
- **Posizione della barra laterale IA** (sinistra o destra), **Dimensione del testo del pannello IA** e **Controllo ortografico nella chat IA**.
- **Apri il pannello IA nei nuovi documenti** — disattivato, un nuovo documento si apre con il pannello richiuso, a un clic di distanza.
- **Salva automaticamente tutti i documenti** attiva il salvataggio automatico per impostazione predefinita in ogni editor; puoi comunque disattivarlo per una singola finestra.
- **Posizione di salvataggio** con il pulsante **Cambia** e **App predefinita per i documenti Office** per far gestire a GenOffice .docx / .xlsx / .pptx.

## Media e ricerca IA

Non sono interruttori: ogni capacità sceglie il fornitore che la serve e la chiave e l'URL di base di un fornitore si inseriscono una volta sola e si condividono:

- **Ricerca web**, **Generazione di immagini**, **Analisi di immagini** e **Analisi video**, ciascuna con un provider, un modello, una chiave e un URL di base.
- **Ricerca file locali** viene eseguita su questa macchina. Al di sotto c'è il **Riordino con Jev**, **disattivato per impostazione predefinita**. Attivalo e gli estratti dei 20 risultati locali migliori — fino a 1.200 caratteri per documento, più i nomi dei file e delle cartelle — vengono inviati al modello Jev di TypeSafe per essere riordinati per pertinenza. Se resta disattivato, nulla lascia il dispositivo.

## Informazioni

- **Versione**, il collegamento GitHub del progetto e un pulsante **Metti una stella su GitHub**.
- **Canale di aggiornamento**: Stabile o Beta. Modificarlo ha effetto immediato e controlla la presenza di un aggiornamento; non riporta un'installazione Beta alla versione Stabile.

## Associazioni come app predefinita

Le impostazioni possono registrare GenOffice come gestore dei file .docx / .xlsx / .pptx / .pdf e simili (registrazione dell'app predefinita a livello di sistema; conferma quando richiesto).

## Note sul software di terze parti e aggiornamenti

- Aiuto ▸ Note sul software di terze parti: l'elenco completo delle licenze open source incluso nell'app.
- Aiuto ▸ Controlla aggiornamenti…: avvia un controllo manuale; se c'è una versione più recente ti propone di installarla.

## Accesso a Genspark

- Il punto d'ingresso per l'accesso (nelle impostazioni o nell'elenco dei progetti cloud) usa un flusso con **codice dispositivo**: GenOffice mostra un codice e apre la pagina di accesso nel browser; una volta completato, l'operazione prosegue automaticamente.
- L'accesso serve soltanto per: l'elenco dei progetti cloud e i modelli ospitati da Genspark. Senza di esso ogni funzione locale e i modelli personalizzati continuano a funzionare.
- La disconnessione si fa con un clic nelle impostazioni.

## Integrazione MCP (per utenti avanzati / client IA)

**Integrazioni** è il riquadro che collega GenOffice a un agente di programmazione, e ha un articolo tutto suo: Collegare un agente di programmazione. In breve — scegli una strada (la riga di comando, oppure MCP), segui quella sezione, poi apri una nuova conversazione e fai la domanda.

![Impostazioni ▸ Integrazioni: i tre passaggi, poi le righe delle skill e le opzioni MCP](img/settings-integrations.png)

Sotto **Server HTTP locale** l'app può anche eseguire il server per conto proprio — un interruttore di attivazione e una porta — e **Avanzato** aggiunge l'URL di controllo integrità e il file di registro, invece di lasciarlo all'assistente. Ascolta solo su localhost.

## Riepilogo rapido dei comandi

| Comando            | Cosa fa                      |
| ------------------ | ---------------------------- |
| `genoffice <file>` | apre un file                 |
| `genoffice mcp`    | avvia il server MCP locale   |
| `genoffice --help` | tutti i comandi e le opzioni |
