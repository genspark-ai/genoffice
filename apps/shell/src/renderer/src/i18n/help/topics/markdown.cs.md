# Editor Markdown

Editor Markdown otevírá soubory .md / .markdown v režimu zdrojový text plus vykreslený náhled.

- **Otevřít**: z domovské stránky nebo Soubor ▸ Otevřít…; funguje i z příkazové řádky.
- **Úpravy**: úpravy v čistém textu; rozšíření GFM (tabulky, kontrolní seznamy, přeškrtnutí, automatické odkazy) se vykreslují v náhledu.
- **Náhled**: živý; relativní zdroje, například obrázky, se načítají vedle dokumentu.
- **Ukládání**: věrné bajt po bajtu — BOM, znaky CRLF i přítomnost prázdného řádku na konci se zachovávají; uložení bez změn soubor nepřepíše.
- **Najít a nahradit**: ctrl+F hledá ve zdroji; Nahradit vše zapíše výsledek.
- **AI**: předvolená tlačítka nechají asistenta dokument přepsat, rozšířit nebo přeložit.

## Pruh nástrojů

Jeden řádek tlačítek nad editorem (nápovědy zobrazíte po najetí):

![Pruh nástrojů Markdown](img/md-toolbar.png)

- **Soubor a historie**: Uložit, Uložit jako…, Zpět, Znovu, Najít; přepínač **Automatické ukládání** vpravo zapisuje změny na disk v pravidelných intervalech.
- **Tlačítko AI**: otevře panel AI; vedle něj jsou předvolby přepsat / rozšířit / přeložit.
- **Styl odstavce** (rozbalovací seznam): přechod mezi základním textem a úrovněmi nadpisů.
- **Formátování v textu**: **tučné**, _kurzíva_, ~~přeškrtnutí~~, `kód v textu`, odkaz.
- **Seznamy**: odrážkový, číslovaný, kontrolní.
- **Vložit**: tabulka, obrázek, vodorovná čára.
- **Vlastnosti**: vloží nebo přeskočí na blok YAML front matter na začátku souboru.
- **Osnova**: přeskakování podle hierarchie nadpisů.
- **Kontrola pravopisu**: zapne nebo vypne kontrolu pravopisu pro tento dokument.

Tři rychlé příklady:

- **Nadpis**: umístěte kurzor na řádek ▸ rozbalovací seznam stylu odstavce ▸ „Nadpis 1“.
- **Tabulka**: klikněte na **Vložit tabulku** ▸ přetáhněte pro počet řádků a sloupců ▸ napište obsah buněk; náhled ji vykreslí ihned.
- **Kontrolní seznam**: vyberte několik řádků ▸ klikněte na **Kontrolní seznam** ▸ každý řádek se změní na `- [ ]`, což se v náhledu zobrazí jako zaškrtávací pole.
