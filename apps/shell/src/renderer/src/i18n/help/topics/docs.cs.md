# Docs: textový editor

Docs je textový procesor podobný Wordu: čte a zapisuje skutečné soubory .docx se stránkováním typu WYSIWYG.

## Pásek

Karty: **Domů / Vložení / Kreslení / Rozložení / Návrh / Odkazy / Revize / Zobrazení**, navíc kontextové karty pro vybraný objekt (Návrh tabulky a rozložení, Formát obrázku, Záhlaví a patka).

- **Domů**: schránka; písmo (včetně velikostí CJK a zvýrazňovacích znaků) s **Vymazat veškeré formátování** a přepínačem **Zobrazit/skrýt formátovací značky**; odstavec (zarovnání/odsazení/řádkování/seznamy) plus **Definovat novou odrážku / Definovat nový formát čísla / Víceúrovňový seznam**, které uloží vaše vlastní styly seznamů do dokumentu; styly (Nadpis 1-6/Normální/Citace, upravitelné); **Podokno stylů** zobrazí úplný seznam.
- **Vložení**: zalomení stránky a oddílu; tabulky (mřížka řádků × sloupců, nebo **Vložit tabulku…** pro přesnou velikost); obrázky, tvary, textová pole; **Titulní strana** a **Prázdná stránka** z připravené galerie; **Graf**; **Iniciála**; **WordArt**; pole (datum, čas, číslo stránky, celkem stránek, název souboru); odkazy, **Záložka** a křížové odkazy; komentáře; záhlaví a zápatí a čísla stránek; symboly a rovnice.
  - **Graf** vloží skutečný grafický objekt s vlastními daty — sloupcový, spojnicový nebo výsečový — a nikoli obrázek. Příkaz _Upravit data_ ve Wordu otevře čísla, která jsou za ním.
- **Kreslení**: inkoust na stránce, ve skupině **Nástroje kreslení** — **Vybrat** se vrátí k úpravě textu, pak **Pero**, **Zvýrazňovač** a **Guma** (jedno kliknutí nebo jedno přetření odstraní celou stopu). Vedle nich je **Styl pera** / **Styl zvýrazňovače** jeden ovládací prvek s barevnými vzorky a řadou šířek; jeho popisek sleduje aktivní nástroj. Inkoust se ukládá do dokumentu jako poznámka plovoucí nad textem, takže přežije uložení i znovu otevření, a **Vymazat vše** v další skupině ho odstraní celý.
- **Rozložení**: okraje, orientace a velikost papíru, sloupce, odsazení a mezery odstavců.
- **Návrh**: motivy, sady barev, vodoznak, rámeček stránky.
- **Odkazy**: obsah (s možností aktualizace), poznámky pod čarou a v textu, popisky, křížové reference.

  ![Karta Odkazy](img/docs-references.png)

- **Revize**: **Editor** prověří celý dokument — pravopis, gramatiku i interpunkci; **Přeložit**; kontrolu pravopisu; komentáře (**Vyřešit komentáře pomocí AI** zpracuje ty otevřené); sledování změn se zobrazeními Všechny značky / Jednoduché značky, přijetí/odmítnutí a **Souhrn revizí od AI**; počítání slov; **Porovnat** s jiným souborem; **Zamknout dokument**.
- **Zobrazení**: pět způsobů, jak se na soubor podívat — **Rozložení při tisku**, **Rozložení webové stránky**, **Osnova**, **Režim čtení** a **Náhled stránky**; oddálit/přiblížit/100 %/šířka stránky/jedna stránka; **Panel AI**; **Tmavý režim**; pravítko, mřížka a navigační podokno; **Nová karta**, **Rozdělit** a **Přepnout karty**; prohledávatelný dialog **klávesových zkratek**.
  - **Tmavý režim** ztmaví stránku a plátno kolem ní, nikoli ale pás — je to Wordovo rozdělení mezi tmavou plochu úprav a tmavé okno. Volba se pamatuje a v každém případě má přednost před motivem aplikace.
  - **Rozdělit** otevře pod ním druhé podokno, které se posouvá nezávisle a zrcadlí první; zavřete ho křížkem × na jeho okraji.

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
  _Zobrazit nadpis 1_ vám nechá obsah, který opravdu projdete očima.
- **Zvýšit úroveň / Snížit úroveň** změní úroveň nadpisu a spolu s ní i úroveň, kterou
  zdědí každý nadpis pod ním — tak se kapitola stane oddílem.
- **Nový nadpis před / za** vloží jeden na místo kurzoru, aniž byste panel opouštěli.
- **Odstranit** smaže nadpis _i vše pod ním_ — a právě tohle je ta položka, se kterou
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
