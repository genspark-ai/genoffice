# Verbinden met een coding agent

GenOffice spreekt het Model Context Protocol, zodat een coding agent je documenten kan lezen, schrijven en renderen via dezelfde engines die de app gebruikt. De agent hoeft niet te gokken naar een bestandsformaat: hij krijgt de getypte op-schemas uit dezelfde definities waarop de executor valideert.

## Registreren vanuit de app

**Instellingen ▸ Integraties** is de plek waar dit gebeurt. Het paneel heeft twee helften, en je kunt een van beide gebruiken of allebei.

**De skill.** Eén rij per coding agent die op deze computer is gevonden — Claude Code, Codex, Cursor, Gemini CLI, Copilot CLI, OpenCode, Windsurf —, elk met **Installeren**, **Bijwerken** en **Verwijderen**, plus **In een andere map installeren…**, **Skill downloaden (zip)** en **Pad kopiëren**. Staat je assistent er niet bij, wijs GenOffice dan een map aan waaruit hij `SKILL.md` leest, of sla de zip op en laat de assistent hem installeren. Skill en MCP kunnen naast elkaar staan: de assistent kiest er één, en ze doen precies hetzelfde.

**MCP.** Er worden twee routes aangeboden: **Gestart door de assistent (aanbevolen)**, waarbij je de getoonde configuratie aan je client toevoegt en de assistent de server zelf start, en de **Lokale HTTP-server**, die de app voor je draait. Hoe dan ook eindigt de assistent tegen GenOffice te praten en typ je nooit een commando.

## Registreren vanaf de opdrachtregel

Hetzelfde vanaf een terminal — dit is de geavanceerde route, en degene om naar te grijpen wanneer de agent ergens staat waar het paneel hem niet kan vinden:

```sh
genoffice mcp install all
```

Het vindt de coding agents op deze computer en schrijft de stdio-serververmelding in elk van hun eigen configuraties, en laat de rest van dat bestand met rust.

```sh
genoffice mcp list             # where each agent stands
genoffice mcp install cursor    # just one
genoffice mcp uninstall cursor  # take it back out
```

Een agent die je op een ongebruikelijke plek hebt geïnstalleerd, heeft `--dir <path>` nodig; `--force` herschrijft een vermelding die er al staat.

Alles wat de server accepteert staat op één scherm — de vormen om te registreren, te verwijderen en op te sommen, het aanbieden via HTTP, en de twee schema-opties:

![De echte uitvoer van genoffice mcp --help: de vormen install, uninstall en list, met de opties --http, --host, --token, --compact-schemas, --dir en --force](img/mcp.png)

## Zonder assistent uitvoeren

Voor een client op een andere computer bied je die in plaats daarvan via HTTP aan:

```sh
genoffice mcp --http 8765 --token <secret>
```

`--host <addr>` bepaalt waar hij luistert. Geef de client hetzelfde token.

Over HTTP gaan er ook bestanden: `PUT /files/<name>` upload er een, elke tool neemt een `http(s)`-adres in plaats van een pad, en uitvoer komt terug als downloadadressen — en, als ze klein genoeg zijn, als ingesloten bronnen. Ops, specs en Markdown worden hoe dan ook inline meegegeven.

## Wat de agent krijgt

Elke opdracht is een tool. De interessantste:

- **`docs`, `sheet`, `slides`** — lezen en bewerken een bestand via het eigen schrijfpad van de app, één **op** per keer. Een nieuwe deck gaat `deck_start`, `deck_page`, `deck_build`.
- **`render`** — één PNG per pagina, opgemaakt door de renderer van de app, zodat de agent naar een slide kan kijken in plaats van ernaar te gokken.
- **`pdf`** — de tekstlaag van een PDF, pagina per pagina, zonder app-proces.
- **`info`** — metadata en een structuuroverzicht, meestal de goedkoopste eerste aanroep op een onbekend bestand.
- **`search`, `image`, `media`** — de providers die in de app zijn ingesteld, zodat de agent geen eigen sleutels nodig heeft.
- **`merge`** — vult een `{{key}}`-sjabloon.

## Schema's en een kleinere budget

`apply` en `create` geven hun parameters `ops`, `cells` en `data` aan met het getypte schema per op, gegenereerd uit `genoffice guide <domain> --json`. Dat is nauwkeurig en niet klein. Een client met een krap contextvenster kan in plaats daarvan om gewone arrays vragen:

```sh
genoffice mcp --compact-schemas
GENOFFICE_MCP_COMPACT_SCHEMAS=1 genoffice mcp
```

## Waarom de skill

Een agent die de op-woordenlijst niet kent, gaat gokken. De skill neemt de referentie en de ontwerphoogedelen mee — hetzelfde materiaal als `genoffice guide` afdrukt —, zodat de assistent ops schrijft waarvan hij de specificatie echt heeft gelezen. Installeer hem vanuit het paneel hierboven, of met `genoffice skill` vanaf een terminal.

## Wat hij in de app kan bereiken

De server is niet beperkt tot bestanden op schijf. Zolang GenOffice draait, kan de agent ook via het venster werken:

- **`open_in_genoffice`** opent een bestand in een tabblad en brengt dat tabblad op de voorgrond.
- **`open_documents`** somt elk document op dat je open hebt — id, type, pad en of er onopgeslagen wijzigingen zijn —, en leest daarna de actuele inhoud van een document of sluit het, waarbij eerst wordt opgeslagen tenzij je hem vraagt het weg te gooien.
- **De inhoudsgereedschappen** nemen die id (of dat pad) als hun `document`-argument, zodat een bewerking terechtkomt in het tabblad dat je al open hebt en het venster ernaar toe overschakelt.

Twee dingen blijven buiten bereik: er is geen AI-paneel, en het updatevenster in de app geldt hier niet.

## Het paneel Lokale HTTP-server

Onder **Lokale HTTP-server** draait de app de server zelf in plaats van hem aan de assistent over te laten: een schakelaar om in te schakelen, een veld **Poort**, een indicator **Actief / Gestopt** en **Generatie op achtergrond** (documenten rechtstreeks naar een pad schrijven zonder de interface te openen), plus een **Voorbeeldclientconfiguratie** om te kopiëren. **Geavanceerd** voegt de twee verbindings-URL's toe — Streamable HTTP en de oudere SSE-URL —, een URL voor **Statuscontrole** en een schakelaar voor **Loggen** die server- en toolactiviteit naar een lokaal bestand schrijft dat je daar **Openen**, **Vernieuwen** of **Wissen** kunt. Hij luistert alleen op localhost.
