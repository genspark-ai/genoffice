# Het AI-assistentpaneel

Elke editor kan het AI-paneel oproepen: selecteer iets, geef een instructie en zie het gestreamde resultaat.

## Openen en gebruiken

![Het AI-paneel in Docs](img/ai-panel.png)

- Instappen: de **AI-knop** op het lint van elke editor, **Vraag AI** in contextmenu's, of Vraag AI op de markeringsbalk.
- Beschrijf de taak in gewone taal (herschrijf dit / maak deze kolom percentages / herbereken deze pagina...) en druk op Enter.
- Antwoorden worden **gestreamd** weergegeven; heeft de AI gereedschap nodig (het document lezen, het bewerken), dan voert het dat uit en gaat het door tot het klaar is.
- **Stop**: onderbreek de huidige beurt op elk moment.
- **Wachtrij**: terwijl een antwoord binnenstroomt, zet Enter uw bericht in de wachtrij in plaats van het weg te gooien — een strook boven het invoerveld telt ze; open hem om te bewerken, verwijderen, herordenen, pauzeren of wissen. Het volgende bericht verstuurt zichzelf wanneer de beurt eindigt; Esc stopt nog steeds de beurt.

## Wat het kan

- **Docs**: herschrijven/uitbreiden/vertalen/samenvatten, tabellen en afbeeldingen invoegen, opmaak aanpassen; elke beurt maakt eerst een momentopname.
- **Sheets**: formules, gegevens invullen, bulktransformaties, opmaak.
- **Slides**: de hele presentatie genereren, indeling bijstellen, teksten herschrijven.
- **PDF**: vragen en antwoorden en samenvattingen op basis van geselecteerde tekst of pagina's.
- **Markdown / HTML**: herschrijven, uitbreiden, vertalen.
- **Start ▸ Ask AI** antwoordt uit uw bestanden met ster en citeert ze — het kennisbank-paneel heeft een eigen sectie in het hoofdstuk Start, en het AI-paneel van Docs heeft dezelfde vaardigheid.

## Terugdraaien en veiligheid

- Het paneel in Docs houdt een **versielijst** bij: één momentopname per beurt, je kunt naar elke terug en die terugstap is zelf met Ctrl+Z ongedaan te maken. Momentopnamen blijven bestaan als je het document opnieuw opent.
- AI-bewerkingen lopen via dezelfde bewerkingsroute als handmatige bewerkingen (ongedaan te maken, pas bij opslaan vastgelegd) — niets slaat je opslagbevestiging over.

## Een selectie voor de AI verbergen

- In Slides, HTML en Markdown kan een selectie uit het zicht van het model worden gehouden: rechtsklik ▸ **Selectie voor AI verbergen**, geef de placeholder een naam, en het model leest `{{naam}}` op zijn plaats. (Markdown activeert het item met de oog-uit-schakelaar op zijn lint; die van HTML zit in het bronpaneel.)
- Het bestand houdt de echte tekst — opslaan, exporteren en afdrukken blijven onaangetast, en de markering overleeft het opnieuw openen. Hetzelfde item maakt zichtbaar (Markdown krijgt een apart item **Selectie weer aan AI tonen**).
- HTML verbergt elementinhouden, attribuutwaarden en stringliteraalss in scripts; de gerenderde pagina verandert nooit.
- Slides verbergt ook hele afbeeldingen, video's en audio, en een dia met verborgen inhoud wordt nooit als screenshot verstuurd: de automatische lay-outcontrole slaat de pagina over en meldt dat.
- Dit geldt alleen voor de AI in de app — de CLI, MCP-tools en headless-export lezen het bestand rechtstreeks en zien de tekst nog steeds.

## Privacy

- Instructies en de relevante documentinhoud gaan naar de **modelservice die je hebt ingesteld** (Genspark gehost of een eigen eindpunt, volgend hoofdstuk); zonder instelling wordt er niets verstuurd.
- Lokale bestanden worden nergens anders heen geüpload; BYOK-sleutels worden opgeslagen in het instellingenbestand van de app op deze machine en alleen in aanvraagheaders verzonden.
