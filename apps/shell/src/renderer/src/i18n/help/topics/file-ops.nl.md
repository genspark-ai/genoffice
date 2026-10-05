# Bestandsbewerkingen: hernoemen, verwijderen, exporteren

Dit hoofdstuk behandelt de bestandsbewerkingen die alle editors delen; de eigen exportopties van elke editor staan in het bijbehorende hoofdstuk.

Het **⋯-menu** in de rij (verschijnt als je over een bestandsrij zweeft) bundelt deze acties:

![Het ⋯-menu van een bestandsrij](img/file-ops.png)

## Hernoemen

Twee instappen, één set controles:

- **⋯ ▸ Naam wijzigen** in een rij op Start.
- **Dubbelklik op een bestandstabblad** om ter plekke te hernoemen (zie [Tabbladen en vensterbeheer](help://tabs-and-windows)).

Regels: de extensie blijft automatisch behouden; ongeldige tekens, afsluitende punten en gereserveerde namen (CON/NUL en dergelijke) worden met een melding geweigerd; een naamconflict in dezelfde map wordt eveneens tegengehouden. Het echte bestand op schijf wordt hernoemd, en de lijsten met recente en favoriete bestanden volgen.

## Verwijderen

- **⋯ ▸ Verwijderen** op Start: verplaatst het bestand naar de **prullenbak van het systeem**, te herstellen vanuit het besturingssysteem.
- Na het verwijderen verschijnt enkele seconden een melding met ongedaan maken — ongedaan maken zet het bestand terug waar het was.

## Dupliceren

**⋯ ▸ Dupliceren** maakt een kopie van <naam> in dezelfde map en opent die als nieuw tabblad; bij een botsing wordt automatisch een teller toegevoegd.

## Opslaan en Opslaan als

- **⌘S / ctrl+S** slaat het huidige bestand op; een naamloos bestand vraagt eerst om locatie en naam.
- **Opslaan als…** schrijft een nieuw bestand en laat het origineel ongemoeid; latere bewerkingen gaan naar het nieuwe bestand.
- Elk opslaan is atomair (tijdelijk bestand + hernoemen); tussentijds afsluiten kan het bestand niet beschadigen.
- Automatisch opslaan komt pas op gang na de eerste handmatige keer opslaan (zie [Snel starten](help://getting-started)).

## Exporteren naar pdf

- **Docs**: Bestand ▸ Exporteren als pdf… (of de knop op het lint), gepagineerd zoals het eruitziet.
- **Slides**: exporteren rastert pagina voor pagina, met voortgang bij grote presentaties.
- **Sheets**: exporteren volgt de afdrukpaginering.
- Exports worden in een verborgen venster weergegeven en belanden waar jij kiest.

## Exporteren naar Word of afbeeldingen

- **Exporteren als Word…** in pdf: zet de pdf om naar een .docx-bestand (lokale conversie; ingewikkelde opmaak wordt zo goed mogelijk benaderd).
- **Docs** kan pagina's exporteren als afbeeldingen (png per pagina).

## Afdrukken

Bestand ▸ Afdrukken… in elke editor (⌘P/ctrl+P) opent het systeemvenster voor afdrukken; pdf's worden afgedrukt met de huidige paginavolgorde en draaiingen.

## Waar naamloze bestanden staan

De locatie die je bij de eerste keer opslaan kiest is hun thuisbasis; daarvoor bestaat het document alleen in het geheugen. Automatisch opslaan neemt het pas over na die eerste keer opslaan.
