# Docs: textový editor

Docs je textový procesor podobný Wordu: čte a zapisuje skutečné soubory .docx se stránkováním typu WYSIWYG.

## Pásek

Karty: **Domů / Vložení / Rozložení / Návrh / Odkazy / Revize / Zobrazení**, navíc kontextové karty pro vybraný objekt (Návrh tabulky, Formát obrázku).

- **Domů**: schránka; písmo (včetně velikostí CJK a zvýrazňovacích znaků); odstavec (zarovnání/odsazení/řádkování/seznamy); styly (Nadpis 1-6/Normální/Citace, upravitelné).
- **Vložení**: zalomení stránky a oddílu, tabulky (také rychlé tabulky), obrázky, tvary, hypertextové odkazy, záhlaví a patka, čísla stránek, datum, textová pole.
- **Rozložení**: okraje, orientace a velikost papíru, sloupce, odsazení a mezery odstavců.
- **Návrh**: motivy, sady barev, vodoznak, rámeček stránky.
- **Odkazy**: obsah (s možností aktualizace), poznámky pod čarou a v textu, popisky, křížové reference.

  ![Karta Odkazy](img/docs-references.png)

- **Revize**: kontrola pravopisu, komentáře, sledování změn (zobrazení Všechny změny/Jednoduché značení), počítání slov.
- **Zobrazení**: pravítko, mřížka, podokno procházení, přiblížení a prohledávatelný dialog **klávesových zkratek**.

## Podokno procházení

**Zobrazení ▸ Navigační podokno** otevře postranní panel s osnovou nadpisů dokumentu,
polem pro hledání v celém dokumentu a náhledem každé stránky. Zda je otevřené, si
pamatuje mezi spuštěními, takže dokument, kterým procházíte podle osnovy, zůstává
procházitelný.

**Osnova** je strom nadpisů. Klikněte pravým tlačítkem na nadpis v ní, abyste jej sbalili i přestavěli, a ne jen jím procházeli:

- **Sbalit / Rozbalit** u nadpisu sbalí celý jeho podstrom: kapitola zmizí, její text
  v dokumentu zůstane.
- **Sbalit vše / Rozbalit vše** sbalí nebo rozbalí všechno
  najednou. U dlouhé zprávy je to rozdíl mezi čitelnou osnovou a zdí textu.
- **Zobrazit úrovně nadpisů** filtruje strom na hloubky, které vás zajímají, takže
  *Zobrazit nadpis 1* vám nechá obsah, který opravdu projdete očima.
- **Zvýšit úroveň / Snížit úroveň** změní úroveň nadpisu a spolu s ní i úroveň, kterou
  zdědí každý nadpis pod ním — tak se kapitola stane oddílem.
- **Nový nadpis před / za** vloží jeden na místo kurzoru, aniž byste panel opouštěli.
- **Odstranit** smaže nadpis *i vše pod ním* — a právě tohle je ta položka, se kterou
  je třeba dávat pozor: maže podstrom, ne řádek.
- **Vybrat nadpis a obsah** označí od nadpisu až do konce jeho podstromu, připraveno k
  úpravě celé části.

## Kontextová nabídka

Klikněte pravým tlačítkem kamkoli do textu — nabídka se přizpůsobí tomu, na co jste klikli. Hlavní skupiny:

- **Schránka**: vystřihnout / kopírovat / vložit / **vložit jako čistý text**.
- **Písmo, odstavec**: změna řady a velikosti, tučné/kurzíva/podtržené, zarovnání/odsazení/řádkování bez cestování na pásek.
- **Synonyma**: vypíše synonyma vybraného slova; kliknutím na jedno ho nahradíte.
- **Přeložit** (AI): přeloží výběr do cílového jazyka (angličtina, zjednodušená čínština, japonština, korejština, francouzština, němčina, španělština, ...) přes panel AI.
- **Nový komentář**: připojí komentář k výběru.
- **Pravopis** (u chybně napsaného slova): navržené náhrady, ignorovat vše, přidat do slovníku, nastavit jazyk kontroly.
- **Hypertextový odkaz**: otevřít / upravit / kopírovat odkaz / odebrat odkaz.
- **Obrázek**: zobrazit obrázek, uložit obrázek jako, **zarovnání textu kolem** (v textu / čtvercově vlevo a vpravo / nahoře a dole / za textem / před textem), pořadí vrstev.
- **Pole** (obsah, čísla stránek): aktualizovat pole / přepnout kódy polí / upravit pole.
- **Číslování seznamu** (uvnitř seznamu): restartovat číslování / pokračovat v číslování / změnit úroveň seznamu / nastavit počáteční hodnotu číslování.
- **Tabulka** (kurzor v tabulce): vložit řádky/sloupce, sloučit / rozdělit buňky, rozdělit tabulku, automaticky přizpůsobit, zarovnání buněk, rozložit řádky/sloupce, vlastnosti tabulky, nabídka mazání, vybrat.

## Úpravy

- Najít a nahradit (ctrl+F / ctrl+H): rozlišení velikosti písmen, celá slova, regulární výrazy.
- Formátovací štětec; hluboké zpět / znovu; možnosti vkládání.
- Tabulky: sloučení a rozdělení buněk, práce s řádky a sloupci, rámečky a stínování, řazení, vzorce.
- Obrázky: zarovnání textu kolem, oříznutí, komprese; kreslicí plátno.

## CJK typografie

- Komprese interpunkce a pravidla dělení řádků (kinsoku) odpovídají Wordu; převod celé a poloviční šířky.
- Nabídka písem pokrývá běžné názvy rodin CJK v Windows a macOS.

## AI

- Tlačítko AI na pásku a postranní panel: přeformulovat, rozšířit, přeložit, shrnout, vložit předvolenou tabulku a libovolné vlastní pokyny.
- Každý krok AI nejprve vytvoří snímek stavu; můžete se vrátit k libovolnému ze seznamu verzí a samotný návrat lze zase vrátit zpět.

## Ukládání a export

- Ukládání .docx přepisuje pouze změněné odstavce — nedotčený obsah zůstává beze změny bajt po bajtu.
- Export do PDF (s aktuálním stránkováním) a na obrázky po jednotlivých stránkách.

## Tisk

ctrl+P přes systémový dialog, stránky WYSIWYG.
