# Modelli IA e impostazioni

## Provider e modelli

Modelli e chiavi si configurano nelle impostazioni (il pulsante a ingranaggio su Home):

![La finestra Impostazioni](img/settings-integrations.png)

- **Genspark ospitato**: effettua l'accesso (flusso con codice dispositivo) e usalo — nessuna configurazione richiesta.
- **Endpoint personalizzati (BYOK)**: Impostazioni ▸ IA accetta un URL di base e una chiave API per ogni protocollo — compatibile con OpenAI, Anthropic, Gemini, DeepSeek, DashScope (qwen) e altro. Le chiavi vivono solo nelle intestazioni delle richieste — mai su disco, nei registri o nell'ambiente dei processi figli.
- Puoi scegliere un modello diverso per ciascuna capacità: chat/generazione, generazione di immagini, analisi delle immagini.
- **Prova connessione**: verifica che l'endpoint sia raggiungibile e che il modello sia visibile prima di salvare.
- Gli URL di base possono contenere un percorso e una stringa di query (stile gateway); i percorsi degli endpoint vengono aggiunti correttamente.

## Integrazione con la CLI (classe Codex)

- Le impostazioni accettano il percorso di un programma CLI locale (le directory home non ASCII e il prefisso ~ funzionano; ~ viene espanso automaticamente); Rileva modelli interroga la CLI sui modelli disponibili.
- La convalida verifica solo l'esistenza — nessuna restrizione sul set di caratteri.

## Quando hanno effetto le modifiche

- Le modifiche a modello ed endpoint hanno effetto immediato; una conversazione in corso mantiene la vecchia configurazione fino al turno successivo.
