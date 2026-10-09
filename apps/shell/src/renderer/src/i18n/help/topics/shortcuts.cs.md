# Klávesové zkratky

Každý editor má vlastní přiřazení kláves a funguje všude jen několik z nich. Tato strana je úplný seznam, rozdělený podle aplikace, s chordem pro macOS a pro Windows / Linux vedle sebe.

## Přečtěte si nejprve

**Tady téměř nic není globální.** Následující tři přístupové klávesy ano; vše za nimi patří aplikaci, která je před vámi, a mění se při přepnutí panelů.

| Akce                                                       | macOS | Windows / Linux |
| ---------------------------------------------------------- | ----- | --------------- |
| Otevřít uživatelskou příručku — funguje kdekoli v aplikaci | F1    | F1              |
| Složit / rozbalit pás — každý editor                       | ⌥⌘R   | Ctrl+F1         |
| Otevřít list Klávesové zkratky — pouze Word                | ⌘/    | Ctrl+/          |

- **F1 otevírá tuto příručku**, ne list zkratek. Patří do menu aplikace, takže funguje z Wordu, Sheets, Slides, PDF i panelu Home.
- **Přepínač pásu je `Ctrl+F1`, nebo `⌥⌘R` na macOS.** Samotné `F1` nic neudělá — F1 je příručka a klávesa pásu ji záměrně nesdílí. `Ctrl+F1` se na macOS také přijímá, ale `⌥⌘R` je klávesa, kterou používá Office pro Mac.
- **List Klávesové zkratky existuje jen ve Wordu** (`⌘/`). Vypisuje přesně tabulku Wordu níže. Jiný editor ho nemá.

![List Klávesové zkratky ve Wordu: akce vlevo, klávesová zkratka vpravo, seskupeno podle Soubor, Úpravy, Formátování textu, Formátování odstavce, Vkládání, Kontrola a nástroje, Zobrazení](img/word-shortcuts.png)

- **`⌘O`, `⌘S` a `⌘P` jsou v každém editoru implementované samostatně** a nechovají se úplně stejně. `⌘P` ve Wordu nejdřív sestaví náhled tisku, aby jeden vytištěný list odpovídal jedné stránce na obrazovce; v PDF otevře tiskové dialogové okno PDF; Sheets a Slides mají zase vlastní.
- **Zbytek je per aplikace.** Když je aktivní panel Wordu, `⌘B` je tučné písmo Wordu. Když je aktivní Slides, `⌘B` je tučné písmo Slides a menu, do kterého patří, není to, na které jste zvyklí. Tam, kde přiřazení funguje jen s výběrem nebo v konkrétním podokně, uvádí to sloupec Akce.

## Konvence

- Sloupec macOS používá známé symboly: `⌘` Command, `⌥` Option, `⇧` Shift, `⌃` Control.
- Sloupec Windows / Linux je vypisuje slovy: `Ctrl`, `Alt`, `Shift`.
- Tam, kde GenOffice na dvou platformách skutečně potřebuje různé klávesy, tabulka to ukazuje. Běžné případy jsou `⌃H` pro Najít a nahradit (macOS má `⌘H` pro Skrýt), `⌥⌘G` pro Přejít na a klávesa zavírání v Slides.
- Chord, který funguje jen s výběrem, je zapsaný ve sloupci Akce, ne v samostatném sloupci, aby tabulka měla pořád tři sloupce.

## Word

Word má plnou lištu nabídky, takže níže uvedené zkratky souboru a formátování fungují vždy, když je aktivní panel Wordu, ať je otevřený jakýkoli dokument. Zkratky úprav vyžadují načtený editovatelný dokument s kurzorem v textu; neselhávají, když píšete do vyhledávacího pole, do promptu AI ani do jiného pole.

### Soubor

| Akce           | macOS | Windows / Linux |
| -------------- | ----- | --------------- |
| Nový           | ⌘N    | Ctrl+N          |
| Nové okno      | ⇧⌘N   | Ctrl+Shift+N    |
| Otevřít…       | ⌘O    | Ctrl+O          |
| Uložit         | ⌘S    | Ctrl+S          |
| Uložit jako…   | ⇧⌘S   | Ctrl+Shift+S    |
| Tisk           | ⌘P    | Ctrl+P          |
| Zavřít (panel) | ⌘W    | Ctrl+W          |

⌘P není tisk prohlížeče. Word nejdřív sestaví náhled tisku, takže jeden vytištěný list je přesně jedna stránka na obrazovce, a pak otevře vlastní tiskové dialogové okno.

### Úpravy

| Akce                    | macOS | Windows / Linux |
| ----------------------- | ----- | --------------- |
| Zpět                    | ⌘Z    | Ctrl+Z          |
| Znovu                   | ⇧⌘Z   | Ctrl+Shift+Z    |
| Vyjmout                 | ⌘X    | Ctrl+X          |
| Kopírovat               | ⌘C    | Ctrl+C          |
| Vložit                  | ⌘V    | Ctrl+V          |
| Vložit jako prostý text | ⌥⇧⌘V  | Ctrl+Shift+V    |
| Vybrat vše              | ⌘A    | Ctrl+A          |
| Najít                   | ⌘F    | Ctrl+F          |
| Nahradit                | ⌃H    | Ctrl+H          |
| Přejít na               | ⌥⌘G   | Ctrl+G          |
| Klávesové zkratky       | ⌘/    | Ctrl+/          |

Nahradit je na obou platformách Ctrl+H, protože ⌘H na macOS skrývá okno a nikdy se nedostane do Wordu. Přejít na kopíruje vlastní rozdělení Wordu: ⌥⌘G na Macu, aby ⇧⌘G zůstal volný pro Počet slov.

### Formátování textu

| Akce                                          | macOS    | Windows / Linux       |
| --------------------------------------------- | -------- | --------------------- |
| Tučné                                         | ⌘B       | Ctrl+B                |
| Kurzíva                                       | ⌘I       | Ctrl+I                |
| Podtržení                                     | ⌘U       | Ctrl+U                |
| Písmo… (dialogové okno)                       | ⌘D       | Ctrl+D                |
| Zvětšit velikost písma                        | ⇧⌘.      | Ctrl+Shift+.          |
| Zmenšit velikost písma                        | ⇧⌘,      | Ctrl+Shift+,          |
| Zvětšit velikost písma o 1 pt                 | ⌘]       | Ctrl+]                |
| Zmenšit velikost písma o 1 pt                 | ⌘[       | Ctrl+[                |
| Horní index                                   | ⇧⌘= / ⌘. | Ctrl+Shift+= / Ctrl+. |
| Dolní index                                   | ⌘,       | Ctrl+,                |
| Změnit velikost písmen                        | ⇧F3      | Shift+F3              |
| Kopírovat formátování (s vybraným textem)     | ⇧⌘C      | Ctrl+Shift+C          |
| Vložit formátování (po Kopírovat formátování) | ⇧⌘V      | Ctrl+Alt+Shift+V      |
| Vymazat veškeré formátování                   | ⌃␣       | Ctrl+Space            |

Dvě zkratky malíře formátování jsou jediné místo, kde klávesa Windows není prosté přejmenování: Vložit formátování má Alt, protože Ctrl+Shift+V tam už vkládá a přebírá styl.

### Formátování odstavce

| Akce                                                | macOS | Windows / Linux |
| --------------------------------------------------- | ----- | --------------- |
| Zarovnat doleva                                     | ⌘L    | Ctrl+L          |
| Vycentrovat                                         | ⌘E    | Ctrl+E          |
| Zarovnat doprava                                    | ⌘R    | Ctrl+R          |
| Zarovnat přes šířku                                 | ⌘J    | Ctrl+J          |
| Řádkování 1,0                                       | ⌘1    | Ctrl+1          |
| Řádkování 1,5                                       | ⌘5    | Ctrl+5          |
| Řádkování 2,0                                       | ⌘2    | Ctrl+2          |
| Zvětšit odsazení                                    | ⌃M    | Ctrl+M          |
| Zmenšit odsazení                                    | ⌃⇧M   | Ctrl+Shift+M    |
| Převislé odsazení                                   | ⌘T    | Ctrl+T          |
| Odebrat převislé odsazení                           | ⇧⌘T   | Ctrl+Shift+T    |
| Vymazat formátování odstavce                        | ⌃Q    | Ctrl+Q          |
| Normální (styl)                                     | ⌥⌘0   | Ctrl+Alt+0      |
| Nadpis 1 (styl)                                     | ⌥⌘1   | Ctrl+Alt+1      |
| Nadpis 2 (styl)                                     | ⌥⌘2   | Ctrl+Alt+2      |
| Nadpis 3 (styl)                                     | ⌥⌘3   | Ctrl+Alt+3      |
| Odstavec… (dialogové okno)                          | ⌥⌘M   | Ctrl+Alt+M      |
| Přesunout odstavec nahoru                           | ⌥⇧↑   | Alt+Shift+↑     |
| Přesunout odstavec dolů                             | ⌥⇧↓   | Alt+Shift+↓     |
| Přesunout položku seznamu níže (na začátku položky) | Tab   | Tab             |
| Přesunout položku seznamu výše (na začátku položky) | ⇧Tab  | Shift+Tab       |

Odsazení, převislé odsazení a Vymazat formátování odstavce fungují na obou platformách jen přes Ctrl: ⌘M okno minimalizuje a ⌘Q ukončí aplikaci, takže chord s Command by byl tichou úpravou při ukončování. Tab změní úroveň seznamu jen tehdy, když je kurzor na začátku položky nebo přesahuje několik položek; uvnitř textu vloží tabulátorový znak a ⇧Tab vždy zvedne úroveň.

### Vkládání

| Akce                      | macOS | Windows / Linux  |
| ------------------------- | ----- | ---------------- |
| Zlom stránky              | ⌘⏎    | Ctrl+Enter       |
| Zlom sloupce              | ⇧⌘⏎   | Ctrl+Shift+Enter |
| Zlom řádku                | ⇧⏎    | Shift+Enter      |
| Odkaz…                    | ⌘K    | Ctrl+K           |
| Nový komentář             | ⌥⌘A   | Ctrl+Alt+A       |
| Vložit poznámku pod čarou | ⌥⌘F   | Ctrl+Alt+F       |
| Vložit poznámku na konec  | ⌥⌘E   | Ctrl+Alt+D       |
| Datum (pole)              | ⌥⇧D   | Alt+Shift+D      |
| Čas (pole)                | ⌥⇧T   | Alt+Shift+T      |
| Nedělitelná mezera        | ⇧⌘␣   | Ctrl+Shift+Space |
| Nedělitelná spojovník     | ⇧⌘-   | Ctrl+Shift+-     |

Vložit poznámku na konec je skutečné rozdělení platforem, ne přejmenování: Word na Macu používá ⌥⌘E, protože ⌥⌘D tam patří do Docku.

### Kontrola a nástroje

| Akce               | macOS | Windows / Linux |
| ------------------ | ----- | --------------- |
| Sledovat změny     | ⇧⌘E   | Ctrl+Shift+E    |
| Počet slov         | ⇧⌘G   | Ctrl+Shift+G    |
| Zkontrolovat       | F7    | F7              |
| Aktualizovat pole  | F9    | F9              |
| Přepnout kódy polí | ⌥F9   | Alt+F9          |

F7 spouští kontrolu pomocí AI a F9 aktualizuje pole; obojí potřebuje načtený dokument. Jakmile omezení úprav zapne Sledovat změny, už se nedá vypnout.

### Zobrazení

| Akce                              | macOS | Windows / Linux |
| --------------------------------- | ----- | --------------- |
| Přiblížit                         | ⌘=    | Ctrl+=          |
| Oddálit                           | ⌘-    | Ctrl+-          |
| Přiblížit na 100%                 | ⌘0    | Ctrl+0          |
| Zobrazit/skrýt značky formátování | ⌘8    | Ctrl+Shift+8    |

Zkratka značek formátování kopíruje vlastní rozdělení Wordu, ⌘8 proti Ctrl+Shift+8. Word řídí přiblížení jen třemi výše uvedenými klávesami menu — ⌘+ tu nic neudělá.

F1 otevírá **uživatelskou příručku**, ne tento seznam. List zkratek je ⌘/ nebo Ctrl+/, a to je položka menu Wordu, takže se objeví jen při aktivním panelu Wordu.

## Sheets

GenOffice Sheets je mřížka kompatibilní s Excelem. Většina zkratek níže jsou zkratky Excelu, které už znáte.

> **Úpravy mají přednost.** Téměř každá zkratka v Sheets se během úpravy buňky nebo psaní do **řádku vzorců** ignoruje, aby psaní, pohyb kurzoru a Backspace si zachovaly svůj obvyklý význam při úpravě textu. Pokud klávesa jako by nic nedělala, zkontrolujte, jestli právě neupravujete: nejdřív stiskněte Esc nebo Enter a pak to zkuste znovu. Příkazy působící na výběr také nic nedělají, dokud je otevřené dialogové okno. Tam, kde je chord výjimkou z tohoto pravidla, uvádí to tabulka.

Dvě konvence jsou specifické pro Sheets. `⌘` a `⌃` jsou stejná pozice modifikátoru — obě se na každé platformě přijímají, takže `⌘1` a `Ctrl+1` jsou stejný chord. **Prázdná buňka ve sloupci Windows / Linux znamená, že chord existuje jen na macOS.**

### Navigace

| Akce                                                              | macOS      | Windows / Linux |
| ----------------------------------------------------------------- | ---------- | --------------- |
| Další list                                                        | ⌘PgDn      | Ctrl+PgDn       |
| Předchozí list                                                    | ⌘PgUp      | Ctrl+PgUp       |
| Další list (v souladu s Excelem pro Mac)                          | ⌥→         |                 |
| Předchozí list (v souladu s Excelem pro Mac)                      | ⌥←         |                 |
| Přejít na první nezmrazenou buňku                                 | ⌘Home      | Ctrl+Home       |
| Přejít na poslední použitou buňku (obsah nebo formátování)        | ⌘End       | Ctrl+End        |
| Dialogové okno Přejít na                                          | ⌘G, F5     | Ctrl+G, F5      |
| Vložit nový list                                                  | ⇧F11       | Shift+F11       |
| Posunout aktivní buňku o obrazovku dolů (a posunout zobrazení)    | Page Down  | Page Down       |
| Posunout aktivní buňku o obrazovku nahoru (a posunout zobrazení)  | Page Up    | Page Up         |
| Posunout aktivní buňku o obrazovku doprava (a posunout zobrazení) | ⌥Page Down | Alt+Page Down   |
| Posunout aktivní buňku o obrazovku doleva (a posunout zobrazení)  | ⌥Page Up   | Alt+Page Up     |

### Výběr

| Akce                                                              | macOS           | Windows / Linux                                     |
| ----------------------------------------------------------------- | --------------- | --------------------------------------------------- |
| Vybrat celý sloupec aktivní buňky                                 | ⌃Space          | Ctrl+Space                                          |
| Vybrat celý řádek aktivní buňky                                   | ⇧Space          | Shift+Space                                         |
| Přejít na okraj bloku dat — kterákoli ze čtyř šipek               | ⌘↑ ⌘↓ ⌘← ⌘→     | Ctrl+↑ Ctrl+↓ Ctrl+← Ctrl+→                         |
| Roztáhnout výběr k okraji bloku dat; stisknutím znovu ho zmenšíte | ⌘⇧↑ ⌘⇧↓ ⌘⇧← ⌘⇧→ | Ctrl+Shift+↑ Ctrl+Shift+↓ Ctrl+Shift+← Ctrl+Shift+→ |
| Přejít na první sloupec aktuálního řádku                          | Home            | Home                                                |

### Úpravy

| Akce                                                                                                | macOS                      | Windows / Linux       |
| --------------------------------------------------------------------------------------------------- | -------------------------- | --------------------- |
| Vyplnit celý výběr aktuálním zadáním                                                                | ⌃⏎ nebo ⌘⏎                 | Ctrl+Enter            |
| Začít nový řádek v buňce                                                                            | ⌥⏎, nebo ⌃⌥⏎ / ⌘⌥⏎ na Macu | Alt+Enter             |
| Upravit aktivní buňku na místě                                                                      | F2                         | F2                    |
| Tučné                                                                                               | ⌘B                         | Ctrl+B                |
| Kurzíva                                                                                             | ⌘I                         | Ctrl+I                |
| Podtržení                                                                                           | ⌘U                         | Ctrl+U                |
| Najít                                                                                               | ⌘F                         | Ctrl+F                |
| Vyplnit doprava z nejlevějšího vybraného sloupce                                                    | ⌘R                         | Ctrl+R                |
| Přepnout odkaz ve vzorci mezi relativní a absolutní                                                 | F4                         | F4                    |
| Přepnout odkaz ve vzorci mezi relativní a absolutní (v souladu s Excelem pro Mac, při úpravě buňky) | ⌘T                         |                       |
| Kopírovat                                                                                           | ⌘C                         | Ctrl+C                |
| Vyjmout                                                                                             | ⌘X                         | Ctrl+X                |
| Vložit                                                                                              | ⌘V                         | Ctrl+V                |
| Vložit pouze hodnoty                                                                                | ⌘⇧V                        | Ctrl+Shift+V          |
| Rychlé vyplnění podle vzoru nad                                                                     | ⌘E                         | Ctrl+E                |
| Zpět                                                                                                | ⌘Z                         | Ctrl+Z                |
| Znovu                                                                                               | ⇧⌘Z                        | Ctrl+Shift+Z          |
| Vybrat celý list                                                                                    | ⌘A                         | Ctrl+A                |
| Vložit dnešní datum                                                                                 | ⌘;                         | Ctrl+;                |
| Vložit aktuální čas                                                                                 | ⌘⇧;                        | Ctrl+Shift+;          |
| Vymazat obsah všech vybraných buněk                                                                 | ⌫ nebo ⌦                   | Backspace nebo Delete |

AutoSum existuje jen na macOS. Stiskněte ⌘⇧T.

### Formátování

| Akce                                                              | macOS | Windows / Linux |
| ----------------------------------------------------------------- | ----- | --------------- |
| Dialogové okno Formát buněk                                       | ⌘1    | Ctrl+1          |
| Přeškrtnutí (s vybranými buňkami)                                 | ⌘5    | Ctrl+5          |
| Zvětšit velikost písma (s vybranými buňkami)                      | ⌘⇧.   | Ctrl+Shift+.    |
| Zmenšit velikost písma (s vybranými buňkami)                      | ⌘⇧,   | Ctrl+Shift+,    |
| Formát čísla: Obecný                                              | ⌘⇧`   | Ctrl+Shift+`    |
| Formát čísla: číslo se dvěma desetinnými místy                    | ⌘⇧1   | Ctrl+Shift+1    |
| Formát čísla: čas `h:mm AM/PM`                                    | ⌘⇧2   | Ctrl+Shift+2    |
| Formát čísla: datum `d-mmm-yy`                                    | ⌘⇧3   | Ctrl+Shift+3    |
| Formát čísla: měna                                                | ⌘⇧4   | Ctrl+Shift+4    |
| Formát čísla: procenta                                            | ⌘⇧5   | Ctrl+Shift+5    |
| Formát čísla: vědecký                                             | ⌘⇧6   | Ctrl+Shift+6    |
| Všechny okraje (s vybranými buňkami)                              | ⌘⇧7   | Ctrl+Shift+7    |
| Bez okrajů (s vybranými buňkami)                                  | ⌘⇧-   | Ctrl+Shift+-    |
| Zrušit malíře formátování (jen když je malíř formátování zapnutý) | Esc   | Esc             |

### Vkládání a mazání

| Akce                                                                                                       | macOS                                  | Windows / Linux                           |
| ---------------------------------------------------------------------------------------------------------- | -------------------------------------- | ----------------------------------------- |
| Vložit buňky — celé řádky nebo sloupce se vloží přímo, cokoli jiného otevře dialogové okno Vložit          | ⌘⇧=                                    | Ctrl+Shift+=                              |
| Vložit buňky — `+` na číselné klávesnici, stejné dialogové okno                                            | ⌘ společně s `+` na číselné klávesnici | Ctrl společně s `+` na číselné klávesnici |
| Odstranit buňky — celé řádky nebo sloupce se odstraní přímo, cokoli jiného otevře dialogové okno Odstranit | ⌘-                                     | Ctrl+-                                    |
| Odstranit buňky — `-` na číselné klávesnici, stejné dialogové okno                                         | ⌘ společně s `-` na číselné klávesnici | Ctrl společně s `-` na číselné klávesnici |
| Skrýt vybrané řádky                                                                                        | ⌘9                                     | Ctrl+9                                    |
| Zobrazit skryté řádky ve výběru                                                                            | ⌘⇧9                                    | Ctrl+Shift+9                              |
| Skrýt vybrané sloupce                                                                                      | ⌘0                                     | Ctrl+0                                    |
| Zobrazit skryté sloupce ve výběru                                                                          | ⌘⇧0                                    | Ctrl+Shift+0                              |

### Data a kontrola

| Akce                                              | macOS | Windows / Linux |
| ------------------------------------------------- | ----- | --------------- |
| Dialogové okno Vložit funkci                      | ⇧F3   | Shift+F3        |
| Správce názvů                                     | ⌃F3   | Ctrl+F3         |
| Přidat nebo upravit poznámku k aktivní buňce      | ⇧F2   | Shift+F2        |
| Vložit hypertextový odkaz                         | ⌘K    | Ctrl+K          |
| sledovat předchůdce (buňka musí obsahovat vzorec) | ⌘[    | Ctrl+[          |
| Sledovat následující                              | ⌘]    | Ctrl+]          |
| Skupovat vybrané řádky                            | ⌥⇧→   | Alt+Shift+→     |
| Zrušit skupování vybraných řádků                  | ⌥⇧←   | Alt+Shift+←     |
| Přepočítat celý sešit                             | F9    | F9              |
| Přepočítat jen aktivní list                       | ⇧F9   | Shift+F9        |
| Zapnout nebo vypnout filtr v označeném rozsahu    | ⌘⇧F   |                 |

Zkratka filtru Excelu, Ctrl+Shift+L, na Windows ani Linuxu nedělá nic. Na macOS stiskněte ⌘⇧F, nebo na jakékoli platformě použijte příkaz filtru na pásu.

### Okno

| Akce                                                      | macOS         | Windows / Linux  |
| --------------------------------------------------------- | ------------- | ---------------- |
| Uložit sešit                                              | ⌘S            | Ctrl+S           |
| Uložit jako                                               | ⇧⌘S           | Ctrl+Shift+S     |
| Otevřít sešit                                             | ⌘O            | Ctrl+O           |
| Tisk                                                      | ⌘P            | Ctrl+P           |
| Zavřít sešit (macOS) / Ukončit aplikaci (Windows a Linux) | ⌘W            | Ctrl+Q           |
| Otevřít uživatelskou příručku                             | F1            | F1               |
| Přiblížit                                                 | ⌘=            | Ctrl+=           |
| Přiblížení kolečkem myši                                  | ⌘ + posouvání | Ctrl + posouvání |
| Zobrazit nebo skrýt vzorce v buňkách                      | ⌘`            | Ctrl+`           |
| Rozbalit nebo sbalit řádek vzorců                         | ⌘⇧U           | Ctrl+Shift+U     |

## Slides

GenOffice Slides je editor prezentací ve stylu PowerPointu. Většina kláves dělá všude to samé, ale pár z nich mění význam podle fokusu — uvnitř textového pole, nebo když má fokus klávesnice **panel náhledů**, **panel osnovy** nebo **třídění snímků**. Tyto případy jsou ve sloupci Akce označené.

### Soubor a aplikace

| Akce                       | macOS | Windows / Linux |
| -------------------------- | ----- | --------------- |
| Uložit                     | ⌘S    | Ctrl+S          |
| Uložit jako                | ⇧⌘S   | Ctrl+Shift+S    |
| Otevřít                    | ⌘O    | Ctrl+O          |
| Zavřít aktuální panel      | ⌘W    | Ctrl+Q          |
| Tisk                       | ⌘P    | Ctrl+P          |
| Zpět                       | ⌘Z    | Ctrl+Z          |
| Znovu                      | ⇧⌘Z   | Ctrl+Shift+Z    |
| Znovu (alternativní chord) | ⌘Y    | Ctrl+Y          |

Zavírání panelu je jediné místo, kde se dvě platformy zcela rozdělí: ⌘W na macOS a na Windows a Linuxu Ctrl+Q, který ukončí celou aplikaci, ne panel.

### Úpravy

| Akce                                         | macOS  | Windows / Linux |
| -------------------------------------------- | ------ | --------------- |
| Nový snímek                                  | ⇧⌘N    | Ctrl+M          |
| Najít (ne při úpravě textu v tvaru)          | ⌘F     | Ctrl+F          |
| Opatrnit výběr pomocí AI (s vybranými tvary) | ⌘K     | Ctrl+K          |
| Přiblížit                                    | ⌘=     | Ctrl+=          |
| Oddálit                                      | ⌘-     | Ctrl+-          |
| Skutečná velikost (100%)                     | ⌘0     | Ctrl+0          |
| Vybrat vše                                   | ⌘A     | Ctrl+A          |
| Přecházet výběrem tvarů vpřed / vzad         | ⇥ / ⇧⇥ | Tab / Shift+Tab |

Nový snímek je druhé rozdělení platforem. Každá platforma má vlastní klávesu a druhá nic nedělá: ⇧⌘N na macOS, Ctrl+M na Windows a Linuxu. ⌘M to na macOS nemůže udělat, protože je to systémová klávesa Minimalizovat.

Vybrat vše sleduje fokus: s fokusem na panel náhledů, panel osnovy nebo třídění snímků vybere všechny snímky, všude jinde všechny prvky aktuálního snímku. Tab prochází tvary na snímku, a to jen když má fokus samotné plátno snímku.

### Formátování textu

Ty působí na text vybraného tvaru, nebo na odstavec s kurzorem, když jste uvnitř textového pole.

| Akce                                  | macOS | Windows / Linux |
| ------------------------------------- | ----- | --------------- |
| Tučné                                 | ⌘B    | Ctrl+B          |
| Kurzíva                               | ⌘I    | Ctrl+I          |
| Podtržení                             | ⌘U    | Ctrl+U          |
| Zarovnat doleva                       | ⌘L    | Ctrl+L          |
| Vycentrovat                           | ⌘E    | Ctrl+E          |
| Zarovnat doprava                      | ⌘R    | Ctrl+R          |
| Zarovnat přes šířku                   | ⌘J    | Ctrl+J          |
| Zvětšit velikost písma o 1 pt         | ⌘]    | Ctrl+]          |
| Zmenšit velikost písma o 1 pt         | ⌘[    | Ctrl+[          |
| Písmo na další velikost v žebříku     | ⇧⌘>   | Ctrl+Shift+>    |
| Písmo na předchozí velikost v žebříku | ⇧⌘<   | Ctrl+Shift+<    |

Klávesy žebříku písma berou také klávesy . a , s drženým Shift, pro rozvržení klávesnice, které je hlásí místo kláves s interpunkcí.

### Schránka a objekty

| Akce                                            | macOS       | Windows / Linux    |
| ----------------------------------------------- | ----------- | ------------------ |
| Kopírovat                                       | ⌘C          | Ctrl+C             |
| Vyjmout                                         | ⌘X          | Ctrl+X             |
| Vložit                                          | ⌘V          | Ctrl+V             |
| Vložit pouze formát                             | ⇧⌘V         | Ctrl+Shift+V       |
| Kopírovat pouze formát (s vybranými tvary)      | ⇧⌘C         | Ctrl+Shift+C       |
| Duplikovat na místě (s vybranými tvary)         | ⌘D          | Ctrl+D             |
| Skupovat výběr (s vybranými tvary)              | ⌘G          | Ctrl+G             |
| Zrušit skupování výběru (s vybranými tvary)     | ⇧⌘G         | Ctrl+Shift+G       |
| Odstranit výběr (s vybranými tvary)             | ⌫ / Delete  | Backspace / Delete |
| Odstranit vybraný vrchol (v režimu Úprava bodů) | ⌫ / Delete  | Backspace / Delete |
| Posunout výběr (s vybranými tvary)              | ↑ ↓ ← →     | Arrow keys         |
| Posunout o jeden pixel obrazovky                | ⌘↑ ⌘↓ ⌘← ⌘→ | Ctrl+Arrow keys    |
| Změnit velikost kolem středu                    | ⇧↑ ⇧↓ ⇧← ⇧→ | Shift+Arrow keys   |

Uvnitř textového pole fungují kopírovat, vyjmout a vložit na text, ne na tvar. Shift se šipkami mění velikost každého tvaru kolem jeho vlastního středu, Spojovatele nechává být a vícenásobný výběr potvrzuje jako jeden krok Zpět.

### Snímky — fokus na panelu náhledů, panelu osnovy nebo třídění

Příkazy na úrovni snímku typu PowerPointu fungují jen tehdy, když má fokus klávesnice některé ze těchto tří podoken.

| Akce                                        | macOS                          | Windows / Linux                                |
| ------------------------------------------- | ------------------------------ | ---------------------------------------------- |
| Předchozí / další snímek                    | ↑ ↓ (také Page Up / Page Down) | ArrowUp / ArrowDown (také Page Up / Page Down) |
| První / poslední snímek                     | Home / End                     | Home / End                                     |
| Roztáhnout výběr na první / poslední snímek | ⇧Home / ⇧End                   | Shift+Home / Shift+End                         |
| Vybrat všechny snímky                       | ⌘A                             | Ctrl+A                                         |
| Odstranit vybrané snímky                    | ⌫ / Delete                     | Backspace / Delete                             |
| Kopírovat snímky                            | ⌘C                             | Ctrl+C                                         |
| Vyjmout snímky                              | ⌘X                             | Ctrl+X                                         |

Pohyb mezi snímky šipkami je výjimka — funguje i z plátna, bez jakékoli vybrané položky. Odstraňování snímků neplatí, když je aktivní pero ani v **režimu čtení**.

### Uvnitř textového pole

| Akce                                                   | macOS                                 | Windows / Linux                           |
| ------------------------------------------------------ | ------------------------------------- | ----------------------------------------- |
| Začít úpravu textu ve vybraném textovém tvaru          | libovolný tisknutelný znak, ⏎ nebo F2 | libovolný tisknutelný znak, Enter nebo F2 |
| Potvrdit text a znovu vybrat tvar                      | Esc nebo F2                           | Esc nebo F2                               |
| Potvrdit text a znovu vybrat tvar (alternativní chord) | ⌘⏎                                    | Ctrl+⏎                                    |
| Přejít na další zástupný symbol                        | ⌃⏎                                    | Ctrl+⏎                                    |
| Přesunout úroveň seznamu níže / výše                   | ⇥ / ⇧⇥                                | Tab / Shift+Tab                           |
| Další / předchozí buňka tabulky                        | ⇥ / ⇧⇥                                | Tab / Shift+Tab                           |
| Měkký zlom řádku                                       | ⇧⏎                                    | Shift+⏎                                   |

Tab znamená dvě různé věci: další buňku v tabulce a přesunutí úrovně seznamu níže v textovém tvaru.

### Prezentace

| Akce                          | macOS             | Windows / Linux                      |
| ----------------------------- | ----------------- | ------------------------------------ |
| Spustit od začátku            | F5                | F5                                   |
| Spustit od aktuálního snímku  | ⇧F5 nebo ⌘⏎       | ⇧F5                                  |
| Další snímek                  | → ↓ ␣ ⏎ Page Down | Right, Down, Space, Enter, Page Down |
| Předchozí snímek              | ← ↑ Page Up       | Left, Up, Page Up                    |
| První snímek                  | Home              | Home                                 |
| Poslední snímek               | End               | End                                  |
| Přejít na snímek N            | napište N, pak ⏎  | napište N, pak Enter                 |
| Černá obrazovka               | B nebo .          | B nebo .                             |
| Bílá obrazovka                | W nebo ,          | W nebo ,                             |
| Obnovit po zhasnutí obrazovky | libovolná klávesa | libovolná klávesa                    |
| Ukončit prezentaci            | Esc               | Esc                                  |

Spuštění od aktuálního snímku bere na macOS ⌘⏎ i ⇧F5. Ve Windows zůstává stejná klávesa vyhrazena pro pohyb mezi zástupnými symboly.

### Režim čtení

| Akce                | macOS             | Windows / Linux                      |
| ------------------- | ----------------- | ------------------------------------ |
| Další snímek        | → ↓ ␣ ⏎ Page Down | Right, Down, Space, Enter, Page Down |
| Předchozí snímek    | ← ↑ Page Up       | Left, Up, Page Up                    |
| Ukončit režim čtení | Esc               | Esc                                  |

Stiskněte **Esc**, abyste se dostali ven z čehokoli, v čem právě jste — výřezu, tvaru, který kreslíte, **Úpravy bodů**, skupiny, pera, zvýrazňovače nebo gumy, přehrávače nebo **režimu čtení** — a abyste zrušili výběr.

Dvě další jsou modifikátory, ne chody: při kreslení držte **Shift**, abyste zachovali proporce nového tvaru, a při otáčení držte **Shift**, abyste přitahovali po krocích 15°.

## PDF

PDF nemá vlastní menu, takže každý chord níže sleduje aktivní panel a funguje jen tehdy, když je vpředu panel PDF.

### Soubor a úpravy

| Akce                    | macOS | Windows / Linux |
| ----------------------- | ----- | --------------- |
| Uložit                  | ⌘S    | Ctrl+S          |
| Zpět                    | ⌘Z    | Ctrl+Z          |
| Znovu                   | ⇧⌘Z   | Ctrl+Shift+Z    |
| Najít (otevřít hledání) | ⌘F    | Ctrl+F          |
| Tisk                    | ⌘P    | Ctrl+P          |

Zpět a Znovu procházejí historií dokumentu; když je kurzor v textovém poli, nechají ⌘Z tomu poli. ⌘P otevírá stejné dialogové okno tisku v aplikaci jako tlačítko Tisk na panelu nástrojů.

### Zobrazení

| Akce                 | macOS            | Windows / Linux     |
| -------------------- | ---------------- | ------------------- |
| Přiblížit            | ⌘=               | Ctrl+=              |
| Oddálit              | ⌘-               | Ctrl+-              |
| Přizpůsobit šířce    | ⌘0               | Ctrl+0              |
| Plynulé přiblížování | ⌘ + kolečko myši | Ctrl + kolečko myši |

Přiblížit reagují také ⌘+ a Ctrl++.

### Pohyb po stránkách

| Akce                        | macOS               | Windows / Linux     |
| --------------------------- | ------------------- | ------------------- |
| Další strana                | →                   | →                   |
| Předchozí strana            | ←                   | ←                   |
| Posunout o obrazovku dolů   | PageDown nebo Space | PageDown nebo Space |
| Posunout o obrazovku nahoru | PageUp              | PageUp              |
| Posunout dolů               | ↓                   | ↓                   |
| Posunout nahoru             | ↑                   | ↑                   |
| Přejít na první stránku     | Home                | Home                |
| Přejít na poslední stránku  | End                 | End                 |

Ty nemají modifikátor, takže obě platformy vypadají stejně. ↓ a ↑ jsou jediný řádek, který se mění s fokusem: v dokumentu posouvají stránku, ale s fokusem v postranním panelu náhledů přeskakují o celou stránku. Nic nedělají, když píšete do pole.

### Výběr a úpravy

| Akce                        | macOS                 | Windows / Linux       |
| --------------------------- | --------------------- | --------------------- |
| Odstranit vybranou poznámku | Delete nebo Backspace | Delete nebo Backspace |

Funguje jen tehdy, když je vybrána poznámka, razítko, redakce nebo obrázek.

### Escape a dialogová okna

| Akce                             | macOS  | Windows / Linux |
| -------------------------------- | ------ | --------------- |
| Zrušit nejvyšší stav             | Escape | Escape          |
| Potvrdit otevřené dialogové okno | Enter  | Enter           |

Escape rozvinuje jednu vrstvu po druhé — nejdřív otevřené dialogové okno, pak překryv AI, koncept textu, panel hledání a nakonec samotný výběr. Když je kurzor ve vyhledávacím poli, Enter a ⇧Enter přecházejí na další a předchozí shodu a Escape zavírá panel; když je kurzor v nové textové poznámce, ⌘⏎ ji potvrzuje.

## Markdown

GenOffice Markdown upravuje Markdown s živým náhledem vedle sebe.

### Soubor a zobrazení

| Akce                     | macOS      | Windows / Linux    |
| ------------------------ | ---------- | ------------------ |
| Uložit                   | ⌘S         | Ctrl+S             |
| Uložit jako              | ⇧⌘S        | Ctrl+Shift+S       |
| Tisk                     | ⌘P         | Ctrl+P             |
| Najít                    | ⌘F         | Ctrl+F             |
| Přepnout zdroj / náhled  | ⌘E         | Ctrl+E             |
| Přiblížit                | ⌘= nebo ⌘+ | Ctrl+= nebo Ctrl++ |
| Oddálit                  | ⌘-         | Ctrl+-             |
| Skutečná velikost (100%) | ⌘0         | Ctrl+0             |

Oddálit bere také ⌘_ a Ctrl+_. Přibližuje také sevření trackpadu, nebo držení ⌘ nebo Ctrl a posouvání nad dokumentem.

### Bloky

| Akce                                                   | macOS | Windows / Linux      |
| ------------------------------------------------------ | ----- | -------------------- |
| Duplikovat vybrané bloky                               | ⌘D    | Ctrl+D               |
| Odstranit vybrané bloky                                | ⇧⌘⌫   | Ctrl+Shift+Backspace |
| Posunout vybraný blok nahoru                           | ⇧⌘↑   | Ctrl+Shift+ArrowUp   |
| Posunout vybraný blok dolů                             | ⇧⌘↓   | Ctrl+Shift+ArrowDown |
| Vložit odkaz na výběr                                  | ⌘K    | Ctrl+K               |
| Odebrat odkaz (stiskněte znovu, když je odkaz aktivní) | ⌘K    | Ctrl+K               |

### Uvnitř matematického bloku

Ty platí jen, když je otevřen plovoucí editor LaTeX.

| Akce                      | macOS | Windows / Linux |
| ------------------------- | ----- | --------------- |
| Použít vzorec             | ⌘⏎    | Ctrl+⏎          |
| Zavřít editor bez použití | Esc   | Esc             |

Použití prázdného vzorce místo toho smaže blok. Editor se také zavře při jakékoli změně dokumentu a kliknutím mimo se zavře.

## HTML

GenOffice HTML zobrazuje editor kódu vedle živého náhledu. Klávesy, které působí na prvek stránky, fungují jen tehdy, když má samotný náhled fokus klávesnice.

### Soubor a zobrazení zdroje

| Akce                                             | macOS      | Windows / Linux    |
| ------------------------------------------------ | ---------- | ------------------ |
| Uložit                                           | ⌘S         | Ctrl+S             |
| Uložit jako                                      | ⇧⌘S        | Ctrl+Shift+S       |
| Přepínat zobrazení (úpravy / rozdělení / náhled) | ⌘\         | Ctrl+\             |
| Najít                                            | ⌘F         | Ctrl+F             |
| Přiblížit                                        | ⌘= nebo ⌘+ | Ctrl+= nebo Ctrl++ |
| Oddálit                                          | ⌘-         | Ctrl+-             |
| Skutečná velikost (100%)                         | ⌘0         | Ctrl+0             |

Oddálit bere také ⌘_ a Ctrl+_.

### Formátování a historie

| Akce                       | macOS | Windows / Linux |
| -------------------------- | ----- | --------------- |
| Tučné                      | ⌘B    | Ctrl+B          |
| Kurzíva                    | ⌘I    | Ctrl+I          |
| Zpět                       | ⌘Z    | Ctrl+Z          |
| Znovu                      | ⇧⌘Z   | Ctrl+Shift+Z    |
| Znovu (alternativní chord) | ⌘Y    | Ctrl+Y          |

Uvnitř editoru kódu nebo formulářového pole přebírají vlastní klávesy úprav daného pole.

### V rámci náhledu, s vybraným prvkem

| Akce                           | macOS      | Windows / Linux    |
| ------------------------------ | ---------- | ------------------ |
| Zeptej se AI na prvek          | ⌘K         | Ctrl+K             |
| Odstranit prvek                | ⌫ / Delete | Backspace / Delete |
| Předchozí sourozenec           | ↑          | ArrowUp            |
| Následující sourozenec         | ↓          | ArrowDown          |
| Vybrat nadřazený prvek         | ←          | ArrowLeft          |
| Vybrat prvního potomka         | →          | ArrowRight         |
| Posunout prvek výš v dokumentu | ⌥↑         | Alt+ArrowUp        |
| Posunout prvek níž v dokumentu | ⌥↓         | Alt+ArrowDown      |
| Zrušit výběr                   | Esc        | Esc                |

Každá z nich potřebuje už vybraný prvek; když nic není vybráno, klávesa nedělá nic.

### V rámci náhledu, při úpravě textu na místě

| Akce                                 | macOS | Windows / Linux |
| ------------------------------------ | ----- | --------------- |
| Potvrdit úpravu                      | Esc   | Esc             |
| Potvrdit úpravu (alternativní chord) | ⌘⏎    | Ctrl+⏎          |
| Ztučnit vybraný rozsah               | ⌘B    | Ctrl+B          |
| Nastavit kurzívu vybranému rozsahu   | ⌘I    | Ctrl+I          |

Potvrzení zachová to, co jste napsali, a vrátí prvek do vybrané polohy.
