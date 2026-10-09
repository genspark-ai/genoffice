# Snel starten: de interface en de basis

GenOffice is een kantoorsuite die volledig op je eigen computer draait: één venster, één rij tabbladen, met zes editors — Docs (tekstverwerking), Sheets (spreadsheets), Slides (presentaties), PDF, Markdown en HTML. De bestanden zijn echte .docx / .xlsx / .pptx / .pdf-bestanden, volledig uitwisselbaar met Word, Excel en PowerPoint. Er is geen netwerk nodig.

## Overzicht van de interface

![Het startscherm](img/home-screen.png)

Het venster bestaat uit drie delen:

- **Tabbladenbalk (boven)**: elk geopend bestand is een tabblad. Het meest linkse Start-tabblad is er altijd en kan niet worden gesloten; de andere tabbladen zijn je documenten. Dubbelklik op een tabblad om de bestandsnaam ter plekke te hernoemen.
- **Inhoudsgebied**: de editor (of Start) die bij het actieve tabblad hoort.
- **Menubalk**: op macOS in de systeemmenubalk, op Windows/Linux bovenaan het venster. De menu's Bestand/Bewerken/Beeld passen zich aan aan de actieve editor.

## Een document maken

Een van de volgende:

- Klik op een snelmakenkaart in het onderdeel [Snel starten](help://getting-started) van **Start** (AI Docs, AI Sheets, AI Slides, ...).
- Menu **Bestand ▸ Nieuw**: Docs (⌘N/ctrl+N), Sheets, Slides, Markdown, HTML of PDF.
- Sleep een bestand naar het venster, of dubbelklik erop in je bestandsverkenner (als GenOffice de standaardapp is).

Een nieuw document opent zonder titel; het bestand op schijf wordt pas aangemaakt bij de eerste keer opslaan.

## Bestanden openen

- Menu **Bestand ▸ Openen…** (⌘O/ctrl+O) opent de systeemkiezer: .docx / .doc / .xlsx / .xls / .csv / .tsv / .pptx / .ppt / .pdf / .md / .html.
- Klik op een willekeurig item in de lijst **Recent** op Start.
- Ook `genoffice <bestand>` in een terminal opent bestanden.

## Het opslagmodel

- **Handmatig opslaan**: ⌘S/ctrl+S, of Bestand ▸ Opslaan / Opslaan als… Bij de eerste keer opslaan wordt om locatie en naam gevraagd.
- **Automatisch opslaan** gaat pas aan nadat je het bestand minstens één keer handmatig hebt opgeslagen — een PDF dat je alleen hebt gelezen wordt nooit stilletjes overschreven. Automatisch opslaan treedt kort op nadat de inhoud is gewijzigd.
- Een tabblad met onopgeslagen wijzigingen sluiten vraagt eerst om Opslaan / Weggooien / Annuleren.
- Elke schrijfactie is atomair (tijdelijk bestand + hernoemen), zodat een stroomstoring nooit een half bestand achterlaat.

## Veelgebruikte sneltoetsen

| Actie                   | macOS | Windows / Linux |
| ----------------------- | ----- | --------------- |
| Nieuw document          | ⌘N    | ctrl+N          |
| Openen                  | ⌘O    | ctrl+O          |
| Opslaan                 | ⌘S    | ctrl+S          |
| Tabblad sluiten         | ⌘W    | ctrl+W          |
| Deze handleiding openen | F1    | F1              |
| Het lint samenvouwen    | ⌥⌘R   | Ctrl+F1         |

**Het lint samenvouwen** werkt in elke editor. De tabbalk blijft staan en de opdrachtbalk eronder verdwijnt; het geselecteerde tabblad dient tegelijk als het besturingselement om het lint te vouwen, dus zolang het lint samengevouwen is, is er geen tabblad geselecteerd en drukken op een willekeurig tabblad brengt de balk terug. Een dubbelklik op een tabblad doet hetzelfde. Hoe u het hebt gelaten, wordt per editor onthouden.

Sneltoetsen binnen elke editor (opmaakpenseel, zoeken en vervangen, tabelbewerkingen, ...) staan in de bijbehorende hoofdstukken; Docs heeft bovendien een doorzoekbaar dialoogvenster met sneltoetsen (**⌘/**) (zie dat hoofdstuk).

## De Option+Command-sneltoetsen

Option+Command is de laag die Word reserveert voor gestructureerde sprongen, en GenOffice vult die op dezelfde manier. Docs neemt daar het grootste deel van in, Sheets neemt er twee van zichzelf om met Excel gelijk te lopen; één sneltoets werkt overal.

**Docs**

| Sneltoets (macOS) | Wat het doet      | Windows / Linux    |
| ----------------- | ----------------- | ------------------ |
| ⌥⌘1 / ⌥⌘2 / ⌥⌘3   | Kop 1 / 2 / 3     | Ctrl+Alt+1 / 2 / 3 |
| ⌥⌘0               | Standaard         | Ctrl+Alt+0         |
| ⌥⌘M               | Alinea            | Ctrl+Alt+M         |
| ⌥⌘A               | Nieuwe opmerking  | Ctrl+Alt+A         |
| ⌥⌘F               | Voetnoot invoegen | Ctrl+Alt+F         |
| ⌥⌘E               | Eindnoot invoegen | Ctrl+Alt+D         |
| ⌥⌘G               | Ga naar           | Ctrl+G             |

Twee daarvan veranderen op Windows, om dezelfde reden waarom Word ze opsplitst. **macOS bezet ⌥⌘D** — die toont en verbergt het Dock —, waardoor de eindnoot ⌥⌘E is op de Mac en Ctrl+Alt+D overal elders. En **Ga naar** laat Alt vallen: Ctrl+G, waar de Mac-sneltoets hem wel meeneemt.

**Sheets**, zolang het raster de focus heeft

| Sneltoets (macOS) | Wat het doet | Windows / Linux |
| ----------------- | ------------ | --------------- |
| ⌥⌘0               | Buitenranden | Ctrl+Shift+7    |
| ⌥⌘−               | Geen rand    | Ctrl+Shift+−    |

Windows is geen herschrijving van het Mac-paar. Excel voor Mac geeft Sheets **allebei** — ⌘⇧7 en ⌥⌘0 zijn twee toetsen voor dezelfde buitenranden —, dus op Windows houdt de opdracht de Ctrl+Shift-plek die het al had, en de Option-laag is er simpelweg niet.

Let op: **⌥⌘0 betekent Standaard in Docs en Buitenranden in Sheets**. Ze komen nooit in dezelfde editor voor, dus in gebruik botst er niets, maar ⌥⌘0 is al bezet en is niet beschikbaar als globale sneltoets.

**Elke editor**: **⌥⌘R / Ctrl+F1** vouwt het lint samen, zoals hierboven beschreven.

Daarmee blijft ⌥⌘D vrij voor GenOffice om op macOS te gebruiken, als een toekomstige opdracht hem nodig heeft.

## Waar nu heen

- Waar je bestanden staan: [Het startscherm](help://home-screen).
- Veel geopende bestanden beheren: [Tabbladen en vensterbeheer](help://tabs-and-windows).
- Het werk door de AI laten doen: [Het AI-assistentpaneel](help://ai-panel).
- Taal, thema, standaardapps: [Instellingen, taal, thema en MCP-integraties](help://settings-integrations).
