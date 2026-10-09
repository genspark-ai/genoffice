# PDF: lezen, annoteren en zwart maken

De pdf-editor heeft vijf linttabs: **Start / Annoteren / Bewerken / Pagina's / Beeld**. Hij leest en schrijft: tekst is bewerkbaar, inhoud kan zwart worden gemaakt en ondertekend en formulieren kunnen worden ingevuld.

## Lezen en navigeren

- Linkerzijbalk: **miniaturen** (klik om te springen, het zichtbare bereik is gemarkeerd) of **overzicht** (bladwijzers, indien aanwezig).
- Zoom: de verhoudingsknop rechtsonder; ctrl+scrollwiel verandert de zoom in stappen.
- Draaien: per pagina of voor alle pagina's via het menu Pagina's; draaiingen worden bij opslaan teruggeschreven.
- Zoeken: ctrl+F op volledige tekst met alle treffers gemarkeerd.
- Versleutelde pdf's: een wachtwoordvraag (in een eigen klein venster) opent ze; het wachtwoord wordt alleen voor deze sessie gebruikt.

## Tekst selecteren en markeren (Annoteren)

Probeer het op een willekeurige alinea:

1. **Sleep de muis over een zin** — bij loslaten verschijnt boven de tekst een annotatiebalk:

![De annotatiebalk na het selecteren van tekst](img/pdf-highlight.png)

2. Kies **markeren** (het gele monster opent een kleurenpalet), **onderstrepen** of **doorhalen**; **Vraag AI** stuurt de selectie met je vraag naar het AI-paneel.
3. Wil je een annotatie ongedaan maken, selecteer dan hetzelfde fragment opnieuw met slepen en klik op de al actieve knop op de balk (een schakelaar in Word-stijl), of selecteer het en druk op Verwijderen.

Verwacht gedrag:

- Slepen over tekst toont een balk: **markeren / onderstrepen / doorhalen / kopiëren / Vraag AI**.
- De kleuren komen uit het palet; **dezelfde markering opnieuw toepassen op een al gemarkeerd bereik verwijdert hem** (schakelaar in Word-stijl).
- Markeringen die al in het bestand staan kunnen worden geselecteerd en verwijderd (⋯-menu of Verwijderen).
- **Let op**: zolang een tekengereedschap actief is, is de tekstlaag niet selecteerbaar — het gereedschap schakelt zichzelf na elke plaatsing uit, zodat je voor de volgende actie weer in selectiemodus bent.

## Tekengereedschappen (Annoteren)

Zes gereedschappen: **inkt, rechthoek, ellips, pijl, notitie**, plus de **zwartmaakvlak** op het tabblad Annoteren.

- Elk gereedschap is een schakelaar: klik om het te activeren; **het schakelt zichzelf uit zodra er een vorm is geplaatst** (klik nogmaals op het gereedschap om door te gaan); klikken op het actieve gereedschap schakelt het ook uit.
- Inkt volgt de lijndikte; rechthoek, ellips en pijl sleep je uit; de kleuren komen uit het tekenpalet.
- Geplaatste vormen kun je selecteren, verwijderen, slepen en (rechthoek en ellips) van grootte veranderen.
- **Zwart maken, de hele werkstroom** (een tekstregel verbergen):

  1. Tabblad Annoteren ▸ klik op **Gebied zwart maken** (het gereedschap wordt actief).
  2. **Sleep een vlak over de inhoud** — die wordt bedekt met een gearceerd merkteken en op de werkbalk verschijnen de knoppen **markeringen wissen / zwart maken toepassen**:

  ![De pagina na het markeren voor zwart maken](img/pdf-redact.png)

  3. Klik op **Zwart maken toepassen** en bevestig — het resultaat is een werkexemplaar waarin de bedekte tekst en afbeeldingen fysiek zijn verwijderd (niet alleen bedekt) en dit kan niet ongedaan worden gemaakt; het oorspronkelijke document blijft ongewijzigd.

  Een fout gemaakt? De knop om markeringen te wissen haalt de huidige markeringen weg, zodat je opnieuw kunt tekenen.

## Plaknotities en reactiedraden

- Het **notitiegereedschap** zet een speld en opent een kaart in de marge voor de tekst (de auteursnaam is instelbaar); bevestigen slaat het op als een standaard pdf-tekstannotatie.
- Klik op een speld om de draad te openen: **beantwoorden** (platte draden in WPS/Acrobat-stijl), je opmerking **bewerken**, één opmerking of een hele draad **verwijderen**.
- Lopende bewerkingen blijven staan totdat het opslaan de nieuwe tekst in dezelfde annotatie in het bestand terugschrijft, waardoor reactieketens intact blijven.

## Pdf-inhoud bewerken (Bewerken)

- **Tekst bewerken**: klik op tekst om die blok voor blok te bewerken (pdfium-engine; lettertypen worden zo goed mogelijk benaderd).
- **Tekst invoegen**: plaats doorzoekbare tekst met zelfgekozen lettertype, maat en kleur.
- **Afbeelding / stempel invoegen**.
- Formulieren: AcroForm-velden vul je direct in; de waarden worden bij opslaan weggeschreven.

## Ondertekeningen

- **Inkthandtekening**: teken hem; hij kan aan een handtekeningveld van een formulier worden gekoppeld.
- **Afbeeldingshandtekening**: plaats een afbeelding als handtekening.
- Opgeslagen handtekeningen zijn herbruikbaar.

## Paginabewerkingen (Pagina's)

- **Draaien / verwijderen / herschikken**: sleep miniaturen om te herschikken; verwijderen vraagt bevestiging.
- **Pagina's importeren**: haal pagina's op uit een andere pdf. **Lege pagina invoegen** voegt een lege pagina toe.
- **Pagina's vervangen** wisselt een bereik in voor pagina's uit een ander bestand; **Pagina's bijsnijden** snijdt de randen bij, met een optie om dit op alle pagina's toe te passen.
- **Paginaformaat** schaalt elke pagina naar één en hetzelfde papierformaat; **Volgorde omkeren** draait het document van einde tot begin.
- **Pagina's extraheren**: exporteer geselecteerde pagina's naar een nieuwe pdf.
- **PDF splitsen**: twee varianten — op basis van bereiken in meerdere bestanden, of elke pagina in een raster van kleinere pagina's snijden.
- **PDF's samenvoegen**: twee varianten — andere pdf's toevoegen, of meerdere pagina's op één vel combineren. De groottes worden opgeteld **voordat er iets gelezen wordt** en in totaal meer dan **1 GiB wordt geweigerd** met een leesbare foutmelding (om het geheugengebruik begrensd te houden).
- Wijzigingen op paginaniveau worden bij de volgende keer opslaan teruggeschreven; Opslaan als… laat het origineel ongemoeid.

## Exporteren en afdrukken

- **Exporteren als Word… / PowerPoint… / Excel…** in het menu Bestand, of dezelfde drie via **PDF converteren** in het lintmenu — alles lokaal, geen upload. De .pptx komt eruit met één dia per pagina en de .xlsx met één werkblad per pagina. Elk ervan vraagt waar het opgeslagen moet worden.
- **Afdrukken**: huidige volgorde en draaiingen via het systeemvenster; paginabereiken worden ondersteund.

## Opslaan

- Gewoon opslaan, handmatig of automatisch, schrijft annotaties en bewerkingen terug in het bestand (atomair).
- **Zwart maken loopt via de eigen werkstroom Toepassen** en levert een kopie op, waardoor het origineel ongemoeid blijft en gevoelige inhoud er niet in achterblijft.
