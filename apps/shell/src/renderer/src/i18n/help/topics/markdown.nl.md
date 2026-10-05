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
