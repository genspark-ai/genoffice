# Slides: prezentace

Slides je editor podobný PowerPointu: čte a zapisuje skutečné soubory.pptx.

## Rozhraní

- **Pásek**: karta Domů obsahuje vkládání a formátování (textové pole/tvar/obrázek/tabulka/graf), písmo a odstavec, zarovnání a uspořádání (vrstvy, zarovnání, rozložení).
- **Pruh miniatur** (vlevo): kliknutím přepnete, tažením změníte pořadí, kontextová nabídka pro nový/duplikát/odstranění.
- **Plátno**: úpravy WYSIWYG; tažení, úchopy pro změnu velikosti, vodicí čáry.
- **Poznámky**: poznámky přednášejícího ke každému snímku, zachovávané při exportu.

## Karty pásku

Pruh karet (na macOS začíná na Domů; Windows přidává kartu Soubor):

- **Domů**: vkládání a formátování — textová pole, tvary, obrázky, tabulky, grafy; písmo a odstavec; zarovnání a uspořádání (vrstvy/zarovnání/rozložení); rozvržení.
- **Vložení**: textové pole, tabulka (s volitelným počtem řádků a sloupců), obrázky, číslo snímku, tlačítko pro přesun (stiskněte během přehrávání, chcete-li přejít na konkrétní snímek) a další.
- **Kreslení**: **pero** (kreslení od ruky na snímek, ukládá se jako stopa pera na stránce) a **zvýrazňovač** (průsvitné, silnější tahy), s tloušťkou pera; opětovné stisknutí nástroje jej zruší.
- **Návrh**: motivy, barevná schéma a pozadí; předlohy a rozvržení.
- **Přechody**: vyberte přechod pro aktuální snímek (účinkuje v prezentátoru PowerPointu), s použitím na všechny; Žádný jej odstraní.

![Karta Přechody](img/slides-transitions.png)

- **Animace**: efekty vstupu a zvýraznění pro vybraný tvar, **pohybové cesty** (pohyb podél cesty); **náhled** přehraje animace snímku na plátně; Žádný je odstraní.

![Karta Animace](img/slides-animations.png)

Vyzkoušejte: vyberte textové pole názvu ▸ karta Animace ▸ zvolte vstupní efekt ▸ **Náhled** jej přehraje na plátně.

- **Prezentace**: spuštění od začátku nebo od aktuálního snímku a nastavení prezentace.
- **Revize**: **nový komentář** k aktuálnímu snímku (zapíše se do souboru.pptx a uvidíte jej v PowerPointu).
- **Zobrazení**: **Normální** (miniatury a plátno), **Zobrazení osnovy** (prohlížení a přeskakování podle textu), **Řazení snímků** (mřížkový přehled, dvojklik otevře úpravy), **Zobrazení pro čtení** (celá obrazovka, stránka po stránce; Esc ukončí); navíc **Předloha snímků**, **Zobrazení prezentujícího**, **Vlastní prezentace** a **Skrýt snímek** a přepínače **Pravítko**, **Mřížka**, **Vodítka**, **Poznámky** a **Podokno miniatur**. Oddálit/přiblížit/100 % a **Přizpůsobit oknu** jsou tady také.
  - Tyto přepínače platí pro jednu relaci — po opětovném otevření aplikace se vrátí na výchozí hodnotu. Výjimkou jsou ty, o které byste nechtěli přijít: pozice **vodicích čar** se ukládají s prezentací a podokno miniatur si pamatuje svou šířku.
  - **Vlastní prezentace** určuje, které snímky přehraje běh v režimu kiosku, a ukládá se s prezentací.

## Kontextové nabídky

Nabídka závisí na místě, kde jste klikli pravým tlačítkem:

- **Na prázdném plátně**: nový snímek (s volbou rozvržení), vystřihnout/kopírovat/vložit snímek (také **vložit jako obrázek** a **vložit zachovávající formátování zdroje**), duplikovat/odstranit/skrýt snímek, **přidat oddíl** (před), přejmenovat oddíl/posunout nahoru a dolů/odstranit/sbalit vše/rozbalit vše, **formátovat pozadí** / **změnit obrázek pozadí**, obnovit rozvržení snímku.
- **Na vybraném prvku**: úprava textu, hypertextový odkaz, velikost a poloha, **zarovnat** (doleva/vodorovně na střed/doprava/nahoru/svisle na střed/dolů) a rozložit vodorovně i svisle, přenést dopředu / odeslat dozadu, seskupit/rozsksupit/znovu seskupit, převrátit vodorovně/svisle, oříznout obrázek / nahradit obrázek / uložit jako obrázek, změnit tvar, **nastavit jako výchozí tvar**, upravit body; textové prvky dostanou navíc tučné/kurzíva/podtržené/odrážky/číslování.
- **Na vybrané tabulce**: vložit řádky/sloupce (nad/pod/vlevo/vpravo), odstranit řádky/sloupce, sloučit / sloučit doprava / sloučit dolů, rozdělit buňku, **stínování buňky**, ukotvení obsahu buňky (nahoru/uprostřed/dolů).

## Obsah

- Textová pole, tvary (výplň/obrys/stín), obrázky (oříznutí/náhrada), tabulky, grafy (sloupcový/pákový/čárový/kruhový..., s editovatelnými daty).
- Předlohy a rozvržení: sjednocují písma, zástupné symboly a pozadí; nové snímky dědí zvolené rozvržení.
- Barvy a písma motivu se řídí motivem.

## Generování pomocí AI

- Karta AI Slides na domovské stránce: zadejte téma nebo osnovu a AI sestaví prezentaci; pokud selže generování v cloudu, použije se místní generování.
- Dál můžete upravovat panelem AI.

## Prezentace a export

- **Exportovat jako PDF**: rastrování stránka po stránce ve skrytém okně, s ukazatelem průběhu a časovým limitem pro velké prezentace.
- Export obrázků: PNG na stránku.
- Prezentace na celou obrazovku tam, kde ji verze poskytuje.

## Ukládání

Úplný přechod tam a zpět ve formátu.pptx: nedotčené stránky zůstávají beze změny bajt po bajtu; poznámky, předlohy a poznámky v okraji se zachovávají.
