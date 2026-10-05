# Panel AI asistenta

Každý editor může vyvolat panel AI: vyberte něco, zadejte pokyn a sledujte výsledek, který se přenáší postupně.

## Otevření a používání

![Panel AI v Docs](img/ai-panel.png)

- Vstupy: **tlačítko AI** na pásku každého editoru, **Zeptat se AI** v kontextových nabídkách nebo Zeptat se AI na liště značení.
- Popište úkol běžnými slovy (přepsat to / udělat z této sloupce procenta / přelayoutovat tuto stránku...) a stiskněte Enter.
- Odpovědi se vykreslují **průběžně**; když AI potřebuje nástroje (přečíst dokument, upravit ho, spustit skript), použije je a pokračuje až do dokončení.
- **Zastavit**: kdykoli přeruší právě probíhající krok.

## Co umí

- **Docs**: přepsat/rozšířit/přeložit/shrnout, vkládat tabulky a obrázky, měnit formátování; každý krok nejprve vytvoří snímek stavu.
- **Sheets**: vzorce, vyplňování dat, hromadné přeměny, formátování.
- **Slides**: vytvoření celé prezentace, úprava rozvržení, přepis textů.
- **PDF**: otázky, odpovědi a shrnutí nad vybraným textem nebo stránkami.
- **Markdown / HTML**: přepsat, rozšířit, přeložit.

## Vrácení zpět a bezpečnost

- Panel v Docs vede **seznam verzí**: jeden snímek na krok, můžete se vrátit k libovolnému a samotný návrat lze zase vrátit pomocí Ctrl+Z. Snímy přežijí i znovu otevření dokumentu.
- Změny od AI procházejí stejnou cestou úprav jako ruční změny (lze je vrátit, vyžadují uložení) — nic neobchází vaše potvrzení uložení.

## Soukromí

- Pokyny a příslušný obsah dokumentu putují do **služby modelu, kterou jste nastavili** (Genspark v cloudu nebo vlastní koncový bod, viz další kapitola); bez nastavení se nic neodesílá.
- Místní soubory se nikam jinam neodesílají; klíče BYOK zůstávají pouze v hlavičkách požadavků — nikdy na disku ani v protokolech.
