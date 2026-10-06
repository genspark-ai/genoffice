# Verbinden met een coding agent

GenOffice spreekt het Model Context Protocol, zodat een coding agent je documenten kan lezen, schrijven en renderen via dezelfde engines die de app gebruikt. De agent hoeft niet te gokken naar een bestandsformaat: hij krijgt de getypte op-schemas uit dezelfde definities waarop de executor valideert.

## Registreren

De gebruikelijke situatie is één commando:

```sh
genoffice mcp install all
```

Het vindt de coding agents op deze computer — Claude Code, Codex, Cursor, Gemini CLI, Copilot CLI, OpenCode, Windsurf — en schrijft de stdio-serververmelding in elk van hun eigen configuraties, en laat de rest van dat bestand met rust.

```sh
genoffice mcp list             # where each agent stands
genoffice mcp install cursor    # just one
genoffice mcp uninstall cursor  # take it back out
```

Een agent die je op een ongebruikelijke plek hebt geïnstalleerd, heeft `--dir <path>` nodig; `--force` herschrijft een vermelding die er al staat.

## Zelf uitvoeren

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

## De skill

Een agent die de op-woordenlijst niet kent, gaat gokken. `genoffice skill` installeert een GenOffice-skill in de agents die hij vindt, en neemt de referentie en de ontwerphoogedelen mee — hetzelfde materiaal als `genoffice guide` afdrukt.

## Wat het niet is

De MCP-server leest en schrijft bestanden. Het is niet het venster: er is geen AI-paneel, en het updatevenster in de app geldt hier niet. Heeft een stap het venster nodig, open dan het bestand.
