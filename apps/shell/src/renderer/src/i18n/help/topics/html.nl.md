# De HTML-editor

De HTML-editor opent .html / .htm met twee modi: **voorbeeld** (de weergegeven pagina) en **bron**.

- **Voorbeeld**: echte weergave; relatieve stylesheets en afbeeldingen worden naast het bestand geladen.
- **Voorbeeldinspecteur**: klik om een element te selecteren, dubbelklik om de tekst ter plekke te bewerken, verwijderen via de werkbalk en Vraag AI over de selectie.
- **Bronmodus**: de HTML bewerken; ctrl+F zoeken, Alles vervangen slaat de herschreven markup op.
- **Opslaan**: bytegetrouw (BOM/CRLF/afsluitende regel blijven behouden); opslaan zonder wijzigingen schrijft niets opnieuw.
- **Zoom**: ctrl+scrollwiel / knijpen schaalt het voorbeeld; ctrl+Z in het voorbeeld maakt de laatste bewerking ongedaan.

## De werkbalk

Klik op een willekeurig element in het voorbeeld en er verschijnt een werkbalk boven:

![De zwevende werkbalk boven een geselecteerd element](img/html-toolbar.png)

- **Bestand en geschiedenis**: Opslaan, Opslaan als…, Ongedaan maken, Opnieuw, Zoeken; de schakelaar **Automatisch opslaan** schrijft wijzigingen op een timer.
- Schakelaar **Voorbeeld / Bron**; **Presenteren** toont de pagina op volledig scherm.
- **Opmaak**: vet, cursief, tekengrootte omhoog/omlaag; het **stijlpaneel** voor het geselecteerde element (kleuren en meer).
- **Invoegen**: kop, alinea, tabel, afbeelding (via koppeling), meer.
- **Afbeeldingsacties** (met een afbeelding geselecteerd): bijsnijden, **achtergrond verwijderen**, vervangen, verhouding vergrendelen.
- **Elementacties** (met een element geselecteerd in de voorbeeldinspecteur): verwijderen, dupliceren, omhoog/omlaag verplaatsen.
- **AI-knop**: opent het AI-paneel; stel je vraag direct over het geselecteerde element.

## Skeleton invoegen

Voor een lege pagina schrijft **Invoegen ▸ Skeleton invoegen** een minimaal document in de standaardmodus:

```html
<!DOCTYPE html>
<html lang="en">
  <head>
    <meta charset="UTF-8">
    <title></title>
  </head>
  <body></body>
</html>
```

Elk onderdeel zit er om een reden, en daarom is het een opdracht en niet iets wat je typt:

- de **doctype**, anders werkt het voorbeeld in de quirksmodus, waar de boxafmetingen en de tabelopmaak andere regels volgen dan je verwacht;
- de **`lang`**, anders heeft een schermlezer geen taal om de pagina in te lezen en kiest de browser een lettertype en een spellingscontrole voor de verkeerde taal;
- de **`charset`**, anders kan een pagina met niet-Latijns tekst als mojibake worden gelezen.

Een viewport-meta is bewust weggelaten: dit tekent zich af in een bureaubladpaneel zonder mobiele viewport waarop het van invloed zou zijn.

De `lang` volgt de taal van de gebruikersinterface van de app, dus het skeleton dat je invoegt is degene waar je gereedschap al op is ingesteld. Daarna kun je het vrij bewerken.

De optie verschijnt alleen in de bewerkingsmodus, en alleen zolang het document leeg is — zodra er inhoud is, is er niets meer om een skeleton *in* in te voegen.

