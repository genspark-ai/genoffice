# Docs: tekstverwerking

Docs is de Word-achtige tekstverwerker: leest en schrijft echte .docx-bestanden met echte WYSIWYG-paginering.

## Het lint

Tabs: **Start / Invoegen / Tekenen / Indeling / Ontwerpen / Verwijzingen / Controleren / Beeld**, plus contextafhankelijke tabs voor het geselecteerde object (Tabelontwerp en indelingsindeling, Afbeeldingsopmaak, kop- en voettekst).

- **Start**: klembord; lettertype (inclusief CJK-maten en nadruktekens) met **Alle opmaak wissen** en een schakelaar voor **Opmaakmarkeringen weergeven/verbergen**; alinea (uitlijning/inspringen/regelafstand/lijsten) plus **Nieuw opsommingsteken definiëren / Nieuwe nummernotatie definiëren / Lijst met meerdere niveaus**, die je eigen lijststijlen in het document opslaan; stijlen (Kop 1-6/Normaal/Citatie, aanpasbaar) met een **Deelvenster Stijlen** voor de volledige lijst.
- **Invoegen**: pagina- en sectie-einden; tabellen (een raster van rijen × kolommen, of **Tabel invoegen…** voor een exacte afmeting); afbeeldingen, vormen, tekstvakken; **Voorblad** en **Lege pagina** uit een kant-en-klare galerij; **Grafiek**; **Decoratiehoofdletter**; **WordArt**; velden (datum, tijd, paginanummer, totaal aantal pagina's, bestandsnaam); hyperlinks, **Bladwijzer** en kruisverwijzingen; opmerkingen; kop- en voettekst en paginanummers; symbolen en vergelijkingen.
  - **Grafiek** voegt een echt grafiekobject met eigen gegevens in — staaf, lijn of cirkel — en geen afbeelding. _Gegevens bewerken_ van Word opent de cijfers erachter.
- **Tekenen**: inkt op de pagina, in de groep **Tekenhulpmiddelen** — **Selecteren** brengt je terug bij het bewerken van tekst, daarna **Pen**, **Markeerstift** en **Gum** (een klik of een veegbeweging haalt de hele streek weg). Daarnaast is **Penstijl** / **Markeerstiftstijl** één besturingselement met kleurstaalmonsters en een rij breedtes; het label volgt het actieve gereedschap. De inkt wordt in het document bewaard als een aantekening die boven de tekst zweeft, dus overleeft hij opslaan en heropenen, en **Alles wissen** in de volgende groep haalt hem helemaal weg.
- **Indeling**: marges, afdrukrichting en papierformaat, kolommen, inspringen en afstand van alinea's.
- **Ontwerpen**: thema's, kleurensets, watermerk, paginarand.
- **Verwijzingen**: inhoudsopgave (actualiseerbaar), voet- en eindnoten, bijschriften, kruisverwijzingen.

  ![Het tabblad Verwijzingen](img/docs-references.png)

- **Controleren**: **Editor** controleert het hele document op spelling, grammatica en interpunctie; **Vertalen**; spellingcontrole; opmerkingen (**AI-opmerkingen** werkt de openstaande door); wijzigingen bijhouden met de weergaven Alle markeringen / Eenvoudige markeringen, accepteren/weigeren en **AI-revisieoverzicht**; woordentelling; **Vergelijken** met een ander bestand; **Document beveiligen**.
- **Beeld**: vijf manieren om het bestand te bekijken — **Afdrukweergave**, **Webindeling**, **Overzicht**, **Leesmodus** en **Paginavoorbeeld**; uitzoomen/innen/100 %/paginabreedte/één pagina; **AI-deelvenster**; **Donkere modus**; liniaal, rasterlijnen en het navigatiedeelvenster; **Nieuw tabblad**, **Splitsen** en **Tabbladen wisselen**; het doorzoekbare dialoogvenster met **sneltoetsen**.
  - **Donkere modus** verduistert de pagina en het tekenvlak eromheen, nooit het lint — Words splitsing tussen een donker bewerkingsvlak en een donker venster. De keuze wordt onthouden en wint hoe dan ook boven het app-thema.
  - **Splitsen** opent daaronder een tweede deelvenster dat onafhankelijk scrolt en het eerste spiegelt; sluit het met de × op de rand.

## Het navigatievenster

**Beeld ▸ Navigatiedeelvenster** opent een zijvenster met de koppenstructuur van het
document, een zoekvak over het hele document en een miniatuur per pagina. Of het
openstaat, wordt onthouden tussen sessies, zodat een document waarin je via de
koppenstructuur navigeert navigeerbaar blijft.

**De koppenstructuur** is de koppenboom. Klik met de rechtermuisknop op een kop in de
koppenstructuur om die te vouwen of te herstructureren in plaats van alleen te
navigeren:

- **Samenvouwen / Uitvouwen** bij een kop vouwt de hele deelboom van die kop:
  het hoofdstuk verdwijnt, de tekst blijft in het document staan.
- **Alles samenvouwen / Alles uitvouwen** vouwt of ontvouwt
  alles in één keer. Bij een lang rapport is dat het verschil tussen een leesbare
  koppenstructuur en een muur van tekst.
- **Kopniveaus weergeven** filtert de boom op de dieptes die je wilt, zodat
  _Kop 1 weergeven_ je een inhoudsopgave overlaat die je echt kunt doorlopen.
- **Niveau verhogen / Niveau verlagen** wijzigt het niveau van de kop en daarmee het
  niveau dat elke kop eronder overneemt: zo wordt een hoofdstuk een sectie.
- **Nieuwe kop ervoor / erna** voegt er een in op de plek van de cursor, zonder het
  venster te verlaten.
- **Verwijderen** haalt de kop _en alles eronder_ weg, en dat is de knop om op te letten:
  hij verwijdert een deelboom, geen regel.
- **Kop en inhoud selecteren** selecteert van de kop tot het einde van de deelboom,
  klaar om een hele sectie te bewerken.

## Contextmenu

Rechtermuisklik ergens in de tekst — het menu volgt waarop je klikte. Hoofdgroepen:

- **Klembord**: knippen / kopiëren / plakken / **plakken als platte tekst**.
- **Lettertype, alinea**: lettertype en maat wijzigen, vet/curatief/onderstreept, uitlijning/inspringen/regelafstand zonder naar het lint te gaan.
- **Synoniemen**: toont synoniemen voor het geselecteerde woord; klik er een aan om het te vervangen.
- **Vertalen** (AI): vertaalt de selectie naar de doeltaal (Engels, vereenvoudigd Chinees, Japans, Koreaans, Frans, Duits, Spaans, ...) via het AI-paneel.
- **Nieuwe opmerking**: koppelt een opmerking aan de selectie.
- **Spelling** (bij een verkeerd gespeld woord): voorgestelde vervangingen, alles negeren, aan woordenboek toevoegen, proefleestaal instellen.
- **Hyperlink**: openen / bewerken / link kopiëren / hyperlink verwijderen.
- **Afbeelding**: afbeelding bekijken, afbeelding opslaan als, **tekstterugloop** (in lijn / rond links en rechts / boven en onder / achter tekst / voor tekst), rangschikking.
- **Velden** (inhoudsopgave, paginanummers): veld bijwerken / veldcodes tonen / veld bewerken.
- **Lijstnummering** (in een lijst): nummering opnieuw starten / nummering voortzetten / lijsniveau wijzigen / startwaarde instellen.
- **Tabel** (cursor in een tabel): rijen/kolommen invoegen, cellen samenvoegen / splitsen, tabel splitsen, automatisch aanpassen, celuitlijning, rijen/kolommen verdelen, tabeleigenschappen, verwijdermenu, selecteren.

## Bewerken

- Zoeken en vervangen (ctrl+F / ctrl+H): hoofdlettergevoelig, heel woord, reguliere expressies.
- Opmaakpenseel; onbeperkt ongedaan maken en opnieuw; plakopties.
- Tabellen: cellen samenvoegen/splitsen, rijen en kolomen bewerken, randen en arcering, sorteren, formules.
- Afbeeldingen: tekstterugloop, bijsnijden, comprimeren; tekenvlak.

## CJK-typografie

- Tekstcompressie en regelafbreekregels (kinsoku) sluiten aan op Word; conversie tussen volle en halve breedte.
- De lettertypenlijst dekt de gangbare CJK-familienamen van Windows en macOS.

## AI

- AI-knop op het lint en het zijpaneel: herschrijven, uitbreiden, vertalen, samenvatten, kant-en-klare tabellen invoegen, plus vrije instructies.
- Elke AI-beurt maakt eerst een momentopname; je kunt terug naar elke versie uit de versielijst, en die terugstap is zelf ongedaan te maken.

## Opslaan en exporteren

- Opslaan als .docx herschrijft alleen de gewijzigde alinea's — ongewijzigde inhoud blijft byte voor byte hetzelfde.
- Exporteren naar pdf (met de huidige paginering) en naar afbeeldingen per pagina.

## Afdrukken

ctrl+P via het systeemvenster, met WYSIWYG-pagina's.
