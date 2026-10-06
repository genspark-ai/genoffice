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

## Export

Nabídka Soubor, všechno místní a všechno se ptá, kam výsledek uložit:

- **Exportovat jako Word…** a **Exportovat jako PDF…** zapíšou skutečný .docx nebo .pdf.
- **Exportovat jako obrázky…** zapíše jeden PNG na každou stránku do složky, kterou vyberete.
- **Převést a otevřít v Docs** převede na .docx a otevře ho ve vestavěné kartě Docs tady v aplikaci — nejde o předání do cloudu a převedená kopie žije v cache složce, která se po asi týdnu uklidí.

## Zobrazení zdroje

Pásek obsahuje přepínač **Zdroj** (překládá se spolu s aplikací). Když jej zapnete, editor nahradí nezpracovaný Markdown: přesně ten text, který uložení zapisuje, nic není zkrášlené, nic není pod ním znormalizované.

- **Úpravy jsou věrné bajtům.** Uložení ze zobrazení zdroje vytvoří stejné bajty jako uložení z editoru — BOM, CRLF i přítomnost koncového znaku nového řádku zůstanou zachovány.
- **Je to stejný dokument.** Klidně mezi nimi přepínejte; zdroj je vlastní text editoru, ne kopie, kterou by bylo nutno slučovat.
- **Pruh nástrojů pro formátování není dostupný**, dokud je zobrazení otevřené, protože většina těchto tlačítek vkládá konstrukce editoru, které dávají smysl jen na vykreslené straně. Když zobrazení zavřete, vrátí se.
- **Soubory JSON a další v režimu zdroje** se otevírají přímo zde: není co vykreslovat, takže zdroj _je_ dokument.
