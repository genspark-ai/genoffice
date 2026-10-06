# Het startscherm: waar je bestanden staan

Start is de beginpagina van GenOffice: een navigatiebalk links, bestandslijsten en snelmakenkaarten rechts.

![Het startscherm](img/home-screen.png)

## Navigatie in de zijbalk

- **Recent**: bestanden die je onlangs hebt geopend. Elke rij draagt zijn moment — vandaag, gisteren of de datum.
- **Favorieten**: bestanden die je hebt gemarkeerd. Beweeg over een bestandsrij en klik op de ster om het toe te voegen of te verwijderen.
- **Gebruikershandleiding**: opent deze handleiding.
- **Genspark Projects**: na het aanmelden bij je Genspark-account zie je de projecten die je op het web met Genspark AI hebt gemaakt; klik er een aan om in de browser verder te bewerken. Zoeken, sorteren op tijd, vernieuwen en meer laden worden ondersteund.
- **Mappen**: de mappen die je met **Map toevoegen…** aan de zijbalk toevoegt, of er vanuit je bestandsverkenner naartoe sleep. Elke map wordt een hoofdmap die je kunt openen, er submappen in kunt aanmaken, hernoemen en verwijderen; een map die offline gaat wordt als niet beschikbaar getoond en kan van de lijst worden gehaald. **Nieuwe map** maakt er nog een.

Er is hier geen Prullenbak-item. Verwijderde bestanden belanden in de prullenbak van het systeem, en ze herstellen is zaak van het besturingssysteem.

## De bestandslijst

Elke rij toont een pictogram, de bestandsnaam, de wijzigingstijd en meer. Het **⋯-menu** in de rij biedt:

- **Openen**, en **Tonen in map** om het bestand in je bestandsverkenner te vinden.
- **Pad kopiëren**.
- **Verplaatsen naar map…**: opent een mapkiezer en verplaatst het bestand echt; staat er op de bestemming al een bestand met dezelfde naam, dan kunt u overslaan, overschrijven of hernoemen.
- **Naam wijzigen**: ter plekke hernoemen, de extensie blijft automatisch behouden.
- **Aan favorieten toevoegen / Uit favorieten verwijderen** — favorieten blijven over een herstart heen staan en volgen het bestand als je het hernoemt.
- **Dupliceren**: maakt een kopie in dezelfde map.
- **Verwijderen**: verplaatst het bestand naar de prullenbak van het systeem — dit is geen definitieve verwijdering.
- **Uit lijst verwijderen**, in het bovenste **Recent**-overzicht, om een item te laten verdwijnen zonder het bestand aan te raken.

### Meerdere bestanden tegelijk

Vink het vakje bij een rij aan, of ⌘/ctrl-klik, om een selectie op te bouwen; het vakje in de kop selecteert alles wat er op dit moment staat, en een balk boven de lijst meldt (**{n} geselecteerd**) hoeveel er geselecteerd zijn en biedt **Verplaatsen naar map…** en **Bestanden verwijderen** voor de hele set. Je kunt een meervoudige selectie ook naar een map in de zijbalk slepen.

## Zoeken

Het zoekvak bovenaan filtert op twee dingen tegelijk:

- **Bestandsnamen**: snel filteren op naam.
- **Bestandsinhoud**: GenOffice indexeert je bestanden op de achtergrond (de tekst in docx/xlsx/pptx/pdf/md/html, met OCR als terugval voor gescande pdf's), dus zoeken naar hoofdtekst vindt ook bestanden. Het bereik en de schakelaars staan in de zoekinstellingen.

## Snelstartkaarten

De kaarten boven de lijsten maken met één stap een nieuw document. Klikken op een kaart maakt een bestand van dat type en opent de bijbehorende editor — je kunt meteen gaan typen of de AI een eerste opzet laten maken (elke editor heeft een **AI-knop** op het lint en **Vraag AI** in het contextmenu van de selectie).

Het nieuwe bestand komt in de map die op dat moment in de zijbalk geselecteerd is; is er geen selectie, dan gaat het naar de standaardmap.

Wat elke kaart doet:

- **AI Docs** (.docx): een leeg tekstdocument in de Docs-editor. Het bestand wordt pas bij de **eerste keer opslaan** naar schijf geschreven; nieuwe documenten openen met het AI-paneel uitgeklapt (zet dat uit in Instellingen → "AI-paneel openen in nieuwe documenten").
- **AI Sheets** (.xlsx): een leeg spreadsheet in de Sheets-editor. Zolang je niet opslaat bestaat er geen bestand op schijf — de naam is gereserveerd voor de eerste keer opslaan; na de eerste AI-generatie kan het bestand ook automatisch op basis van de inhoud worden hernoemd.
- **AI Slides** (.pptx): een lege presentatie in de Slides-editor.
- **AI Markdown** (.md): een leeg Markdown-document in de Markdown-editor.
- **AI HTML** (.html): een lege webpagina in de HTML-editor.
- **AI PDF** (.pdf): anders dan de rest — maakt **onmiddellijk** een echt leeg pdf-bestand van één pagina in de doelmap en opent het als gewoon bestand (de pdf-editor werkt met echte bestanden). Handig om aantekeningen te maken, gebieden zwart te maken of tekst toe te voegen; bij de eerste keer opslaan kan het bestand automatisch op basis van de inhoud worden hernoemd.
- **Lokaal bestand openen**: een systeemkiezer voor Word (.docx/.doc), Excel (.xlsx/.xlsm/.xls/.csv/.tsv), PowerPoint (.pptx/.ppt), pdf, Markdown (.md/.markdown) en webpagina's (.html/.htm). Je kunt meerdere bestanden selecteren; elk bestand krijgt een eigen tabblad.

> Tip: Bestand ▸ Nieuw in de menubalk maakt dezelfde soort documenten (⌘N/Ctrl+N maakt standaard een tekstdocument); een bestand naar het venster slepen opent het.

## Cloudprojecten (Genspark Projects)

- Voor het eerste gebruik moet je je aanmelden bij je Genspark-account (flow met apparaatcode: GenOffice toont een code en jij rondt het aanmelden in de browser af).
- De projectlijst synchroniseert met het web; Openen in browser springt daarheen om verder te gaan.
- Niet aanmelden heeft geen invloed op een enkele lokale functie.
