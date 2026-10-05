# Docs: tekstverwerking

Docs is de Word-achtige tekstverwerker: leest en schrijft echte .docx-bestanden met echte WYSIWYG-paginering.

## Het lint

Tabs: **Start / Invoegen / Indeling / Ontwerpen / Verwijzingen / Controleren / Beeld**, plus contextafhankelijke tabs voor het geselecteerde object (Tabelontwerp, Afbeeldingsopmaak).

- **Start**: klembord; lettertype (inclusief CJK-maten en nadruktekens); alinea (uitlijning/inspringen/regelafstand/lijsten); stijlen (Kop 1-6/Normaal/Citatie, aanpasbaar).
- **Invoegen**: pagina- en sectie-einden, tabellen (inclusief snelle tabellen), afbeeldingen, vormen, hyperlinks, kop- en voettekst, paginanummers, datum, tekstvakken.
- **Indeling**: marges, afdrukrichting en papierformaat, kolommen, inspringen en afstand van alinea's.
- **Ontwerpen**: thema's, kleurensets, watermerk, paginarand.
- **Verwijzingen**: inhoudsopgave (actualiseerbaar), voet- en eindnoten, bijschriften, kruisverwijzingen.

  ![Het tabblad Verwijzingen](img/docs-references.png)

- **Controleren**: spellingcontrole, opmerkingen, wijzigingen bijhouden (weergaven Alle wijzigingen/Enkele markeringen), woordentelling.
- **Beeld**: liniaal, rasterlijnen, navigatievenster, zoom en het doorzoekbare dialoogvenster met **sneltoetsen**.

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
