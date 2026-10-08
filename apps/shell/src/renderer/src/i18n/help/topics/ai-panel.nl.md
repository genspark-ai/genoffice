# Het AI-assistentpaneel

Elke editor kan het AI-paneel oproepen: selecteer iets, geef een instructie en zie het gestreamde resultaat.

## Openen en gebruiken

![Het AI-paneel in Docs](img/ai-panel.png)

- Instappen: de **AI-knop** op het lint van elke editor, **Vraag AI** in contextmenu's, of Vraag AI op de markeringsbalk.
- Beschrijf de taak in gewone taal (herschrijf dit / maak deze kolom percentages / herbereken deze pagina...) en druk op Enter.
- Antwoorden worden **gestreamd** weergegeven; heeft de AI gereedschap nodig (het document lezen, het bewerken), dan voert het dat uit en gaat het door tot het klaar is.
- **Stop**: onderbreek de huidige beurt op elk moment.

## Wat het kan

- **Docs**: herschrijven/uitbreiden/vertalen/samenvatten, tabellen en afbeeldingen invoegen, opmaak aanpassen; elke beurt maakt eerst een momentopname.
- **Sheets**: formules, gegevens invullen, bulktransformaties, opmaak.
- **Slides**: de hele presentatie genereren, indeling bijstellen, teksten herschrijven.
- **PDF**: vragen en antwoorden en samenvattingen op basis van geselecteerde tekst of pagina's.
- **Markdown / HTML**: herschrijven, uitbreiden, vertalen.

## Terugdraaien en veiligheid

- Het paneel in Docs houdt een **versielijst** bij: één momentopname per beurt, je kunt naar elke terug en die terugstap is zelf met Ctrl+Z ongedaan te maken. Momentopnamen blijven bestaan als je het document opnieuw opent.
- AI-bewerkingen lopen via dezelfde bewerkingsroute als handmatige bewerkingen (ongedaan te maken, pas bij opslaan vastgelegd) — niets slaat je opslagbevestiging over.

## Privacy

- Instructies en de relevante documentinhoud gaan naar de **modelservice die je hebt ingesteld** (Genspark gehost of een eigen eindpunt, volgend hoofdstuk); zonder instelling wordt er niets verstuurd.
- Lokale bestanden worden nergens anders heen geüpload; BYOK-sleutels worden opgeslagen in het instellingenbestand van de app op deze machine en alleen in aanvraagheaders verzonden.
