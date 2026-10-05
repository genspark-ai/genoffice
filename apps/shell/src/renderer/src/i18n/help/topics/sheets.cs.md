# Sheets: tabulky

Sheets je editor podobný Excelu; výpočty zajišťuje samostatný proces s enginem napsaným v Rustu (jeho pád nikdy neukončí aplikaci). Otevírá a ukládá skutečné soubory .xlsx; soubory .csv a .tsv se otevírají jako tabulky.

## Rozhraní

- **Pásek**: osm karet, postupně popsaných níže.
- **Řádek vzorců**: zobrazuje a upravuje vzorec aktivní buňky; podporovány jsou běžné funkce.
- **Karty listů** (dole): přidat / přejmenovat / odstranit / přesunout list.
- **Úprava buňky**: dvojklikem nebo rovnou psáním; Enter potvrdí a posune dolů, Tab doprava, Esc zruší (zvyky z Excelu).
- **Zkratky**: ve stylu řady Excelu (ctrl+C/V/X, ctrl+Z/Y, ctrl+F, ...).

## Karty pásku

- **Domů**: písmo, výplň, rámečky, číselné formáty (měna/procenta/tisíce, zvýšení a snížení desetinných míst), zarovnání, sloučení, vkládání a velikost řádků a sloupců, podmíněné formátování, formátovat jako tabulka, styly buněk, schránka a formátovací štětec, řazení a filtrování.
- **Vložení**: tvary, ikony, symboly, rovnice, snímky obrazovky a další.
- **Rozložení stránky**: barvy a písma motivu, přepínače tisku mřížky a záhlaví, náhled zalomení stránek.
- **Vzorce**: automatický součet a vkládání funkcí, definice názvů (také z výběru), sledování předchůzců a závislostí, sledovací okno, přepočítání listu nebo sešitu.
- **Data**: řazení a filtrování (včetně pokročilého filtru a zrušení filtru), text do sloupců, sloučení sešitů, aktualizovat vše.
- **Revize**: procházení komentářů (zobrazit, předchozí/ další), přeložit.
- **Zobrazení**: přepínače mřížky a záhlaví, přiblížení, zobrazení Normální / náhled zalomení stránek.
- **Návrh grafu**: zobrazí se po výběru grafu — typ grafu, styly a barvy, úprava rozsahu dat.

Karta Data, tlačítko po tlačítku (zleva doprava na obrázku):

![Karta Data](img/sheets-data.png)

- **Kontingenční tabulka**: sestaví kontingenční tabulku z aktuální oblasti; přetažením polí vytvoříte agregaci.
- **Aktualizovat**: přepočítá data aktuální kontingenční tabulky.
- **Z textu/CSV**: importuje soubor .csv/.txt jako nový list, rozdělený podle oddělovače.
- **Sloučit sešity**: přenese listy z jiných souborů .xlsx do tohoto souboru.
- **Aktualizovat vše**: přepočítá všechny kontingenční tabulky i externí zdroje dat.
- **Seřadit** (rozbalovací seznam): vzestupně / sestupně / vlastní řazení (pravidla pro více sloupců).
- **Filtrovat**: přidá rozbalovací seznamy ▼ do řádku záhlaví; zaškrtněte hodnoty, které chcete ponechat.
- Malá tlačítka nad sebou vedle něj: **Vymazat** (vrátí všechny řádky), **Použít znovu** (spustí aktuální filtr znovu), **Pokročilé** (filtruje pomocí oblasti kritérií).
- **Text do sloupců** (rozbalovací seznam): rozdělí jeden sloupec na několik podle oddělovače nebo pevné šířky.
- **Rychlé vyplnění**: zadejte jeden příklad a zbytek sloupce se podle něj doplní (ctrl+E).
- **Odebrat duplicity**: odstraní duplicitní řádky podle vybraných sloupců.
- **Ověření dat** (rozbalovací seznam): pravidla zadávání pro výběr (rozbalovací seznamy, číselné rozsahy...).
- **Konsolidovat**: sloučí několik oblastí na jednom místě podle kategorie.
- **Analýza hypotéz** (rozbalovací seznam): hledání řešení / tabulky dat.
- **Skupit / Zrušit seskupení** (rozbalovací seznam): skupiny řádků nebo sloupců se sbalením a rozbalením.
- **Součet**: vloží dílčí součty pro jednotlivé kategorie.

Karta Vzorce, tlačítko po tlačítku:

![Karta Vzorce](img/sheets-formulas.png)

- **Vložit funkci** (fx): vyhledá funkce s průvodcem po argumentech.
- **Automatický součet** (rozbalovací seznam): SUMA jedním kliknutím, dále průměr/počet/maximum/minimum.
- **Nedávno použité / Finanční / Logické / Textové / Datum a čas / Vyhledávání a reference / Matematické a trigonometrické / Další**: procházejte a vkládejte funkce podle kategorie.
- **Správce názvů**: zobrazit, vytvořit a odstranit pojmenované oblasti.
- **Definovat název** (rozbalovací seznam): pojmenuje výběr; **Použít ve vzorci** vloží existující název; **Vytvořit z výběru** pojmenuje oblasti podle jejich řádku či sloupce záhlaví.
- **Sledovat předchůzce / Sledovat závislosti**: modré šipky ukazují, odkud data vzorce pocházejí a kam směřují; **Odebrat šipky** je vymaže.
- **Zobrazit vzorce**: buňky zobrazují samotný vzorec místo výsledku.
- **Kontrola chyb**: vyhledá a vysvětlí chyby ve vzorcích.
- **Sledovací okno**: připne buňky, o kterých chcete vědět, a sleduje jejich živé hodnoty.
- **Možnosti výpočtu** (rozbalovací seznam): automatický nebo ruční přepočet; v ručním režimu je spustí **Přepočítat nyní / Přepočítat list**.

## Čísla a formátování

- Číselné formáty: obecný, číslo, měna, procento, datum/čas, zlomek, vědecký a další.
- Zarovnání, zvětšení řádku, sloučené buňky, rámečky a výplně.
- Výšky řádků a šířky sloupců tažením; dvojklikem na hranici se velikost přizpůsobí automaticky.

## Data

**Řazení a filtrování** (například sestupně podle jednoho sloupce):

1. Klikněte na **libovolnou buňku v tomto sloupci** (není potřeba vybrat celý sloupec).
2. Karta Domů ▸ **Seřadit a filtrovat** ▸ **Sestupně**; celé řádky se přesunou společně (oblast se řadí jako jeden celek).
3. Pro vlastní pravidla (více sloupců, podle barvy): stejná cesta, **Vlastní řazení**.
4. Filtr: vyberte řádek záhlaví a klikněte na **Seřadit a filtrovat ▸ Filtrovat** — každé záhlaví získá rozbalovací seznam ▼, kde zaškrtáte hodnoty, které chcete ponechat; zrušením filtru se vše vrátí.

- Řazení a filtrování.
- Ukotvení oblastí.
- .csv / .tsv: otevírají se přímo jako tabulka (tsv s oddělovačem tabulátorem se zpracuje jako jeden); uložením se zapíše původní formát.

## Kontextové nabídky

- **V mřížce**: vlastní nabídka editoru (Univer) — vystřihnout/kopírovat/vložit, vkládání a mazání řádků a sloupců, skrytí, sloučení buněk, ukotvení oblastí a další běžné položky.
- **Na dolním stavovém pruhu**: vyberte, které statistiky má stavový pruh zobrazovat (průměr / počet / součet, ...); volba se zapamatuje.
- **Na kartě listu dole**: přidat / přejmenovat / odstranit / obarvit / skrýt listy (nabídka karet Univer).
- Kontextová nabídka horního pruhu karet je popsána v [Karty a správa oken](help://tabs-and-windows).

## AI

- Postranní panel AI: vyberte oblast a zadejte pokud v přirozeném jazyce (přeformátovat, vygenerovat data, napsat vzorce).
- Změny provedené AI lze z panelu vrátit zpět.

## Ukládání a export

- Ukládání .xlsx (vzorce a formáty se zachovávají); Uložit jako…; export do PDF používá stránkování pro tisk.
- Automatické ukládání se řídí globálním pravidlem (zapne se po prvním ručním uložení).

## Stabilita

- Výpočtový proces v Rustu je oddělen od uživatelského rozhraní: pokud ho zabijí extrémní data, uvidíte zprávu a pokus o obnovení relace — nikoli pád aplikace.
