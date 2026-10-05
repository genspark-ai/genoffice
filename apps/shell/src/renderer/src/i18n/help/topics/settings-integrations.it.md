# Impostazioni, lingua, tema e integrazioni MCP

## Apertura delle impostazioni

La riga dell'account in basso a sinistra in Home apre il pannello delle impostazioni (quando non hai effettuato l'accesso riporta Accedi); le opzioni relative all'IA si trovano nella sua sezione Modello IA

![La finestra Impostazioni](img/settings-integrations.png) — la configurazione dei modelli è descritta in Modelli IA e impostazioni.

## Lingua

- Le impostazioni offrono **21 lingue per l'interfaccia**: inglese, cinese semplificato, giapponese, coreano, francese, tedesco, spagnolo, thai, indonesiano, russo, arabo, portoghese, italiano, polacco, ceco, olandese, malese, ebraico, hindi, cinese tradizionale, vietnamita.
- Il cambio ha effetto immediato e viene ricordato; la barra dei menu nativa viene ricostruita nella nuova lingua.

## Tema

Chiaro / Scuro / Segui il sistema. L'opzione Segui il sistema segue l'aspetto del sistema operativo e gli editor si adattano in sincronia senza sfarfallii.

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

GenOffice incorpora un **server MCP locale**, così i client IA esterni (Claude Desktop, Cursor, ...) possono leggere e scrivere direttamente i tuoi documenti:

- Avvio: `genoffice mcp` sulla riga di comando (porta e token di autenticazione configurabili; solo loopback per impostazione predefinita).
- Capacità: creare/aprire/modificare file docx, xlsx e pptx, leggere i contenuti, convertire i formati, esportare in PDF e altro — lo stesso insieme di strumenti usato dalle app desktop.
- Sicurezza: l'autenticazione con token è facoltativa ma consigliata; il listener resta sulla macchina locale per impostazione predefinita; vedi `genoffice mcp --help`.

## Riepilogo rapido dei comandi

| Comando            | Cosa fa                    |
| ------------------ | -------------------------- |
| `genoffice <file>` | apre un file               |
| `genoffice mcp`    | avvia il server MCP locale |
| `genoffice --help` | tutti i comandi e le opzioni |
