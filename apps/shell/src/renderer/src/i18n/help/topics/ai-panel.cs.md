# Panel AI asistenta

Každý editor může vyvolat panel AI: vyberte něco, zadejte pokyn a sledujte výsledek, který se přenáší postupně.

## Otevření a používání

![Panel AI v Docs](img/ai-panel.png)

- Vstupy: **tlačítko AI** na pásku každého editoru, **Zeptat se AI** v kontextových nabídkách nebo Zeptat se AI na liště značení.
- Popište úkol běžnými slovy (přepsat to / udělat z této sloupce procenta / přelayoutovat tuto stránku...) a stiskněte Enter.
- Odpovědi se vykreslují **průběžně**; když AI potřebuje nástroje (přečíst dokument, upravit ho), použije je a pokračuje až do dokončení.
- **Zastavit**: kdykoli přeruší právě probíhající krok.
- **Fronta**: zatímco odpověď běží, Enter zařadí zprávu do fronty, místo aby ji zahodil — pruh nad polem je počítá; rozbalte ho pro úpravu, odebrání, přeřazení, pauzu nebo vymazání. Další zpráva se pošle sama koncem tahu; Esc tahu i nadále ukončuje.

## Co umí

- **Docs**: přepsat/rozšířit/přeložit/shrnout, vkládat tabulky a obrázky, měnit formátování; každý krok nejprve vytvoří snímek stavu.
- **Sheets**: vzorce, vyplňování dat, hromadné přeměny, formátování.
- **Slides**: vytvoření celé prezentace, úprava rozvržení, přepis textů.
- **PDF**: otázky, odpovědi a shrnutí nad vybraným textem nebo stránkami.
- **Markdown / HTML**: přepsat, rozšířit, přeložit.
- **Domů ▸ Ask AI** odpovídá z vašich souborů s hvězdičkou a cituje je — panel znalostní báze má vlastní sekci v kapitole Domů a AI panel Docs nese tutéž dovednost.

## Vrácení zpět a bezpečnost

- Panel v Docs vede **seznam verzí**: jeden snímek na krok, můžete se vrátit k libovolnému a samotný návrat lze zase vrátit pomocí Ctrl+Z. Snímy přežijí i znovu otevření dokumentu.
- Změny od AI procházejí stejnou cestou úprav jako ruční změny (lze je vrátit, vyžadují uložení) — nic neobchází vaše potvrzení uložení.

## Skrýt výběr před AI

- V Slides, HTML a Markdown lze výběr schovat před očima modelu: pravý klik ▸ **Skrýt výběr před AI**, pojmenujte zástupný text a model čte `{{název}}` místo něj. (Markdown položku zapíná přepínačem přeškrtnutého oka na pásu; ta HTML žije v panelu zdroje.)
- Soubor si ponechá skutečný text — uložení, export i tisk zůstávají nedotčené a značka přežije znovuotevření. Tatáž položka skrytí zruší (Markdown má navíc **Znovu ukázat výběr AI**).
- HTML skrývá obsah prvků, hodnoty atributů a řetězcové literály ve skriptech; vykreslená stránka se nikdy nemění.
- Slides skrývá i celé obrázky, videa a audio a snímek se skrytým obsahem nikdy neodchází jako snímek obrazovky: automatická kontrola rozvržení stránku přeskočí a oznámí to.
- Platí to jen pro AI v aplikaci — CLI, nástroje MCP a bezhlavý export čtou soubor přímo a text stále vidí.

## Soukromí

- Pokyny a příslušný obsah dokumentu putují do **služby modelu, kterou jste nastavili** (Genspark v cloudu nebo vlastní koncový bod, viz další kapitola); bez nastavení se nic neodesílá.
- Místní soubory se nikam jinam neodesílají; klíče BYOK se ukládají do souboru nastavení aplikace na tomto počítači a odesílají pouze v hlavičkách požadavků.
