# AI-modellen en instellingen

## Providers en modellen

Modellen en sleutels stel je in Instellingen in (de tandwielknop op Start):

![Het venster Instellingen](img/settings-integrations.png)

- **Genspark gehost**: meld je aan (flow met apparaatcode) en je kunt ermee aan de slag — nul configuratie.
- **Eigen eindpunten (BYOK)**: Instellingen ▸ AI-model vraagt een basis-URL en een API-sleutel per protocol — OpenAI-compatibel, Anthropic, Gemini, DeepSeek, DashScope (qwen) en meer. Sleutels staan alleen in aanvraagheaders — nooit op schijf, in logbestanden of in de omgeving van onderliggende processen.
- Per functie kun je een ander model kiezen: chat/genereren, afbeeldingen genereren, afbeeldingsanalyse.
- **Verbinding testen**: controleert vóór het opslaan of het eindpunt bereikbaar is en het model zichtbaar is.
- Basis-URL's mogen een pad en een querystring bevatten (gateway-stijl); eindpuntpaden worden correct toegevoegd.

## Integratie met de opdrachtregel (Codex-klasse)

- Instellingen accepteren het pad naar een lokaal CLI-programma (niet-ASCII-thuismappen en een ~-voorvoegsel werken; ~ wordt automatisch uitgebreid); Detecteer modellen vraagt de CLI welke modellen beschikbaar zijn.
- De validatie controleert alleen of het pad bestaat — geen beperkingen op de tekenset.

## Wanneer wijzigingen ingaan

- Wijzigingen aan model en eindpunt gaan meteen in; een lopend gesprek houdt de oude configuratie aan tot de volgende beurt.
