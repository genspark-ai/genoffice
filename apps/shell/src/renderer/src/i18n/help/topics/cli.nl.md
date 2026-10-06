# Opdrachtregel en agents

Elke installatie bevat een `genoffice`-commando dat dezelfde engines aanstuurt als het venster — dezelfde parsers, dezelfde writer, dezelfde renderer. Een bestand dat de app opslaat en een bestand dat de opdracht schrijft zijn hetzelfde bestand, en een controle die het AI-paneel van de app doorstaat, doorstaat de opdracht ook.

```sh
genoffice --help          # every command
genoffice guide slides    # the op reference, for writing your own
```

Dat is het hele oppervlak op één scherm — elke opdracht met een regel die zegt wat hij doet, en daarna de globale opties en de exitcodes:

![De echte uitvoer van genoffice --help: de versiebanner, elke opdracht met een beschrijving op één regel, en de globale opties en de exitcodes](img/cli.png)

## De opdracht bemachtigen

macOS en Windows leveren hem mee in de app-bundel. Wil je hem bij naam gebruiken, voer dan eenmalig `genoffice install-cli` uit: die maakt een symbolische koppeling naar het meegeleverde binaire bestand in `/usr/local/bin`, of in je gebruikers-`PATH` op Windows.

## De opdrachten die je moet kennen

| Opdracht          | Wat het doet                                                                                                                                                                                                                                                                  |
| ----------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `open`            | Opent een document in de app; start de app als die niet draait.                                                                                                                                                                                                               |
| `selection`       | Wat de gebruiker in dat bestand op dit moment heeft geselecteerd, zolang de app draait — de aanwijzer voor "deze" / "hier". Geeft een blokbereik, een bereik in het werkblad, elementen op een slide of een pagina terug, afhankelijk van de editor waarin het bestand staat. |
| `convert`         | Converteert tussen formaten met de eigen engines van de app.                                                                                                                                                                                                                  |
| `create`          | Maakt een document van gestructureerde inhoud.                                                                                                                                                                                                                                |
| `render`          | Eén PNG per pagina, zoals de renderer het opmaakt.                                                                                                                                                                                                                            |
| `pdf`             | Leest de tekstlaag van een PDF pagina per pagina, zonder app-proces.                                                                                                                                                                                                          |
| `info`            | Metadata en een structuuroverzicht van een document.                                                                                                                                                                                                                          |
| `search`          | Zoeken op het web of in afbeeldingen via de provider die in de app is ingesteld.                                                                                                                                                                                              |
| `image` / `media` | Genereert een afbeelding, of beschrijft een afbeeldings-, video- of audiobestand en beantwoordt vragen erover.                                                                                                                                                                |
| `merge`           | Vult `{{key}}`-placeholders in een `.docx`-, `.pptx`- of `.xlsx`-sjabloon.                                                                                                                                                                                                    |
| `capabilities`    | Meldt welke cloudfuncties op deze computer zijn ingesteld.                                                                                                                                                                                                                    |
| `guide`           | De op-referentie en de ontwerphoogedelen, gegenereerd uit dezelfde definities waarop de executor valideert — dus kan niet afwijken van wat `apply` accepteert. `--json` geeft ze met het schema van elke op.                                                                  |
| `install-cli`     | Zet `genoffice` in de `PATH`.                                                                                                                                                                                                                                                 |
| `skill`           | Toont de coding agents op deze computer en installeert of werkt de GenOffice-skill daarin bij.                                                                                                                                                                                |
| `mcp`             | Biedt elke opdracht aan als Model Context Protocol-tool. Zie **Verbinden met een coding agent**.                                                                                                                                                                              |

## Bewerken: docs, sheets, slides

`genoffice docs`, `genoffice sheet` en `genoffice slides` lezen en bewerken via hetzelfde schrijfpad als de app, en ze delen één woordenlijst: een **op** is één bewerking, en een **spec** is een lijst van ops die in volgorde worden toegepast.

```sh
genoffice sheet read report.xlsx
genoffice slides apply deck.pptx --ops ops.json --dry-run
genoffice docs apply brief.docx --ops ops.json
```

`--dry-run` meldt wat een batch zou doen en schrijft niets, wat de goedkoopste manier is om een spec te controleren voordat die landt. Het AI-paneel in de app draait op precies deze ops, dus alles wat je het kunt vragen kun je ook scripten.

## Model Context Protocol

`genoffice mcp` biedt elke opdracht aan als MCP-tool, en `genoffice mcp install <agent|all>` registreert die in de eigen configuratie van een coding agent. Die kant staat in **Verbinden met een coding agent**.

## Wat de opdracht niet doet

Het leest en schrijft het bestand. Het is niet de app: er is geen venster, en het updatevenster in de app geldt hier niet. Alles wat het venster nodig heeft — het AI-paneel, de kwaliteitscontrole over een gerenderde slide — moet wachten tot je het bestand opent. `genoffice render` levert je de pixels zonder venster, en `genoffice slides` controleert de opmaak van een deck zelf.
