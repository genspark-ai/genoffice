# De Markdown-editor

De Markdown-editor opent .md / .markdown met een ervaring van bron plus weergegeven voorbeeld.

- **Openen**: vanaf Start of Bestand ▸ Openen…; de opdrachtregel werkt ook.
- **Bewerken**: bewerken in platte tekst; GFM-uitbreidingen (tabellen, takenlijsten, doorhaling, automatische koppelingen) worden in het voorbeeld weergegeven.
- **Voorbeeld**: live; relatieve bronnen zoals afbeeldingen worden naast het document opgezocht.
- **Opslaan**: bytegetrouw — BOM, CRLF en de aanwezigheid van een afsluitende lege regel blijven behouden; opslaan zonder wijzigingen schrijft het bestand niet opnieuw.
- **Zoeken en vervangen**: ctrl+F zoekt in de bron; Alles vervangen schrijft het resultaat terug.
- **AI**: voorinstellingsknoppen laten de assistent het document herschrijven, uitbreiden of vertalen.

## De werkbalk

Eén rij knoppen boven de editor (zweef voor tooltips):

![De Markdown-werkbalk](img/md-toolbar.png)

- **Bestand en geschiedenis**: Opslaan, Opslaan als…, Ongedaan maken, Opnieuw, Zoeken; de schakelaar **Automatisch opslaan** rechts schrijft wijzigingen op een timer naar schijf.
- **AI-knop**: opent het AI-paneel; ernaast staan de voorinstellingen herschrijven / uitbreiden / vertalen.
- **Alineastijl** (keuzelijst): wisselen tussen bodytekst en de verschillende kopniveaus.
- **Tekstopmaak**: **vet**, _cursief_, ~~doorhaling~~, `inline code`, koppeling.
- **Lijsten**: opsomming, genummerd, takenlijst.
- **Invoegen**: tabel, afbeelding, horizontale lijn.
- **Eigenschappen**: voegt het YAML front matter-blok bovenaan het bestand in of springt erheen.
- **Overzicht**: springen op basis van de kopniveaus.
- **Spellingcontrole**: de spellingcontrole voor dit document aan- of uitzetten.

Drie snelle voorbeelden:

- **Kop**: zet de cursor op de regel ▸ keuzelijst met alineastijl ▸ "Kop 1".
- **Tabel**: klik op **Tabel invoegen** ▸ sleep het aantal rijen en kolommen ▸ typ in de cellen; het voorbeeld geeft het meteen weer.
- **Takenlijst**: selecteer een paar regels ▸ klik op **Takenlijst** ▸ elke regel wordt `- [ ]`, in het voorbeeld weergegeven als selectievakjes.

## Exporteren

Menu Bestand, alles lokaal en alles vraagt waar het resultaat heen moet:

- **Exporteren als Word…** en **Exporteren als PDF…** schrijven een echte .docx of .pdf.
- **Exporteren als afbeeldingen…** schrijft één PNG per pagina naar een map die u kiest.
- **Converteren en openen in Docs** zet om naar .docx en opent het in het ingebouwde Docs-tabblad hier in de app — het is geen overdracht naar iets in de cloud, en de omgezette kopie staat in een cachemap die na ongeveer een week wordt opgeschoond.

## Bronweergave

Het lint bevat een schakelaar **Bron** (vertalen mee met de app). Zet hem aan en de editor wordt vervangen door de ruwe Markdown: precies de tekst die een opslaactie wegschrijft, niets is opgeschoond, niets wordt eronder genormaliseerd.

- **Bewerken is bytegetrouw.** Een opslaactie vanuit de bronweergave levert dezelfde bytes als een opslaactie vanuit de editor — BOM, CRLF en de aanwezigheid van een afsluitende newline blijven allemaal behouden.
- **Het is hetzelfde document.** Schakel rustig heen en weer; de bron is de eigen tekst van de editor, geen kopie die je zou moeten samenvoegen.
- **De opmaakwerkbalk is niet beschikbaar** zolang de weergave open is, omdat de meeste van die knoppen editorconstructies invoegen die alleen aan de weergegeven kant betekenis hebben. Hij komt terug zodra je de weergave sluit.
- **JSON en andere bestanden in bronmodus** openen hier direct: er valt niets weer te geven, dus de bron _is_ het document.
