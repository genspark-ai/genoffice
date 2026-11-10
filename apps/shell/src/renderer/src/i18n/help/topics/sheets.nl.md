# Sheets: spreadsheets

Sheets is de Excel-achtige editor; de berekeningen draaien in een apart proces met een engine in Rust (een crash daar brengt de app nooit naar beneden). Het opent en slaat echte .xlsx-bestanden op; .csv en .tsv openen als tabel.

## De interface

- **Lint**: acht tabs, die hieronder één voor één worden behandeld.
- **Formulebalk**: toont en bewerkt de formule van de actieve cel; gangbare functies worden ondersteund.
- **Tabbladen van werkbladen** (onderaan): werkbladen toevoegen / hernoemen / verwijderen / verplaatsen.
- **Cels bewerken**: dubbelklik of gewoon typen; Enter bevestigt en gaat omlaag, Tab gaat naar rechts, Esc annuleert (Excel-gewoonten).
- **Sneltoetsen**: afgestemd op de Excel-familie (ctrl+C/V/X, ctrl+Z/Y, ctrl+F, ...).

## Linttabs

- **Start**: lettertype, opvulling, randen, getalnotaties (valuta/procent/duizendtallen, decimalen omhoog/omlaag), uitlijning, samenvoegen, rijen en kolommen invoegen en hun grootte, voorwaardelijke opmaak, opmaken als tabel, celstijlen, klembord en opmaakpenseel, sorteren en filteren.
- **Invoegen**: afbeeldingen, vormen, tekstvakken, koppelingen, opmerkingen, selectievakje, koptekst en voettekst, symbolen, vergelijkingen; **neven grafiektypen** (kolom, staaf, lijn, vlak, cirkel, spreiding, radar, ring, en kolom + lijn) plus **Aanbevolen grafieken**, die de typen voor de selectie rangschikt en in het voorbeeld toont; **Draaigrafiek** vanuit de cel waarin u staat; **Sparklines** (lijn, kolom, winst/verlies); **Slicer** en **Tijdlijn** om een draaitabel te filteren.
  - Grafieken, draaigrafieken en sparklines zijn echte objecten in de opgeslagen werkmap.
  - **Slicer** en **Tijdlijn** zijn sessiebesturingen: de filtering die ze uitvoeren wordt opgeslagen en Excel toont dezelfde gefilterde draaitabel, maar de slicerknop zelf maakt geen deel uit van het bestand.
- **Pagina-indeling**: themakleuren en -lettertypen, schakelaars voor het afdrukken van rasterlijnen en koppen, voorbeeld van pagina-einden.
- **Formules**: Autosom en functies invoegen, namen definiëren (ook vanuit een selectie), voorwaarden en gevolgen traceren, controlevenster, werkblad of werkmap opnieuw berekenen.
- **Gegevens**: sorteren en filteren (inclusief geavanceerd filter en filter wissen), tekst naar kolommen, werkmap samenvoegen, alles vernieuwen.
- **Controleren**: opmerkingen doorbladeren (tonen, vorige/volgende), vertalen en de groep **Beveiligen** — **Blad beveiligen**, **Werkmap beveiligen** en **Bewerken van bereiken toestaan**.
  - De beveiliging wordt in de .xlsx geschreven en op wat deze app toepast zit geen wachtwoord, dus dezelfde knop wordt **Beveiliging opheffen…** en maakt het ongedaan. Beveiliging uit een ander programma dat _wél_ een wachtwoord heeft, kan hier niet worden verwijderd.
  - De twee beveiligingsknoppen werken meteen — er is geen dialoog om te annuleren, alleen een aantekening in de statusbalk dat het bij het opslaan wordt weggeschreven.
  - **Bewerken van bereiken toestaan** markeert de cellen die bewerkbaar blijven terwijl de rest van het blad vergrendeld is.
- **Beeld**: schakelaars voor rasterlijnen, **Formulebalk**, koppen en het markeren van de actieve rij en kolom; zoom; **Normaal** en **Pagina-eindevoorbeeld**. Rasterlijnen en koppen worden met het werkblad opgeslagen; het markeren is een eigen voorkeur.
- **Grafiekontwerp**: verschijnt als er een grafiek geselecteerd is — grafiektype, stijlen en kleuren, gegevensbereik bewerken.

Het tabblad Gegevens, knop voor knop (van links naar rechts op de afbeelding):

![Het tabblad Gegevens](img/sheets-data.png)

- **Draaitabel**: bouwt een draaitabel uit het huidige bereik; sleep velden om te aggregeren.
- **Vernieuwen**: berekent de gegevens van de huidige draaitabel opnieuw.
- **Uit tekst/CSV**: importeert een .csv/.txt als nieuw werkblad, gesplitst op het scheidingsteken.
- **Werkmap samenvoegen**: haalt werkbladen uit andere .xlsx-bestanden in dit bestand op.
- **Alles vernieuwen**: berekent alle draaitabellen en externe gegevensbronnen opnieuw.
- **Sorteren** (keuzelijst): oplopend / aflopend / aangepast sorteren (regels over meerdere kolommen).
- **Filteren**: voegt ▼-keuzelijsten toe aan de koprij; vink de waarden aan die je wilt behouden.
- De kleine gestapelde knoppen ernaast: **Wissen** (alle rijen terug), **Opnieuw toepassen** (het huidige filter opnieuw uitvoeren), **Geavanceerd** (filteren met een criteriebereik).
- **Tekst naar kolommen** (keuzelijst): splitst één kolom op in meerdere op een scheidingsteken of vaste breedte.
- **Snel invullen**: geef één voorbeeld en de rest van de kolom vult zich daarnaar; (ctrl+E).
- **Dubbele waarden verwijderen**: laat dubbele rijen vallen op basis van de geselecteerde kolommen.
- **Gegevensvalidatie** (keuzelijst): invoerregels voor de selectie (keuzelijsten, getalbereiken...).
- **Consolideren**: voegt meerdere bereiken op één plek samen per categorie.
- **What-if-analyse** (keuzelijst): doel zoeken — lost één invoercel op zodat een formulecel de doelwaarde bereikt.
- **Groeperen / Groepering opheffen** (keuzelijst): rij- of kolomgroepen met in- en uitklappen.
- **Subtotaal**: voegt subtotalen per categorie in.

Het tabblad Formules, knop voor knop:

![Het tabblad Formules](img/sheets-formulas.png)

- **Functie invoegen** (fx): zoek functies met een argumentwizard.
- **Autosom** (keuzelijst): SOM met één klik, plus gemiddelde/aantal/max/min.
- **Onlangs gebruikt / Financieel / Logisch / Tekst / Datum en tijd / Zoeken en verwijzen / Wiskundig en trigonometrisch / Meer**: blader door functies per categorie en voeg ze in.
- **Naambeheer**: bekijk, maak en verwijder benoemde bereiken.
- **Naam definiëren** (keuzelijst): geeft de selectie een naam; **In formule gebruiken** voegt een bestaande naam in; **Maken op basis van selectie** benoemt bereiken op basis van hun koprij of -kolom.
- **Voorwaarden traceren / Gevolgen traceren**: blauwe pijlen die laten zien waar de gegevens van een formule vandaan komen en waar ze naartoe gaan; **Pijlen verwijderen** wist ze.
- **Formules tonen**: cellen tonen de formule zelf in plaats van het resultaat.
- **Foutcontrole**: lokaliseert en verklaart formulefouten.
- **Controlevenster**: zet cellen vast die je volgt en bekijk hun live waarden.
- **Berekeningsopties** (keuzelijst): automatisch of handmatig herberekenen; in de handmatige stand start **Nu berekenen / Werkblad berekenen** het zelf.

## Getallen en opmaak

- Getalnotaties: algemeen, getal, valuta, procent, datum/tijd, breuk, wetenschappelijk en meer.
- Uitlijning, afbreken, samengevoegde cellen, randen en opvullingen.
- Rijhoogtes en kolombreedtes door slepen; dubbelklik op een grens past automatisch aan.

## Gegevens

**Sorteren en filteren** (bijvoorbeeld aflopend op één kolom):

1. Klik op **een willekeurige cel in die kolom** (je hoeft niet de hele kolom te selecteren).
2. Tabblad Start ▸ **Sorteren en filteren** ▸ **Aflopend**; hele rijen verschuiven mee (het gebied wordt als één geheel gesorteerd).
3. Voor eigen regels (meerdere kolommen, op kleur): dezelfde route, **Aangepast sorteren**.
4. Filteren: selecteer de koprij en klik op **Sorteren en filteren ▸ Filteren** — elke kop krijgt een ▼-keuzelijst waarin je de te behouden waarden aanvinkt; het filter wissen haalt alles terug.

- Sorteren en filteren.
- Vensters bevriezen.
- .csv / .tsv: worden direct als tabel geopend (een door tabs gescheiden tsv wordt als één tabel gelezen); bij opslaan wordt het oorspronkelijke formaat teruggeschreven.

## Contextmenu's

- **In het raster**: het eigen menu van de editor (Univer) — knippen/kopiëren/plakken, rijen en kolommen invoegen en verwijderen, verbergen, cellen samenvoegen, vensters bevriezen en andere alledaagse items.
- **Op de statusbalk onderaan**: kies welke statistieken de statusbalk toont (gemiddelde / aantal / som, ...); de keuze wordt onthouden.
- **Op een werkbladtab onderaan**: werkbladen toevoegen / hernoemen / verwijderen / kleuren / verbergen (het tabbladmenu van Univer).
- Het contextmenu van de bovenste tabbladenbalk staat in [Tabbladen en vensterbeheer](help://tabs-and-windows).

## AI

- Het AI-paneel aan de zijkant: selecteer een bereik en geef instructies in gewone taal (opnieuw opmaken, gegevens genereren, formules schrijven).
- Voeg bestanden toe aan een prompt met de 📎-knop, of sleep ze op het paneel; ze reizen mee met de vraag en afbeeldingen komen terug als miniaturen.
- Een antwoord kan naar een cel verwijzen — klik op de verwijzing en het raster springt er naartoe.
- AI-wijzigingen kun je vanuit het paneel terugdraaien.

## Opslaan en exporteren

- Opslaan als .xlsx (formules en opmaak blijven behouden); Opslaan als…; exporteren naar pdf volgt de afdrukpaginering.
- Automatisch opslaan volgt de algemene regel (gaat aan na de eerste handmatige keer opslaan).

## Stabiliteit

- De rekenkern in Rust is procesmatig gescheiden van de interface: als extreme gegevens hem beëindigen, krijg je een melding en een poging de sessie te herstellen — geen app-crash.
