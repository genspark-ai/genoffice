# Billentyűparancsok

Minden szerkesztőnek megvannak a saját billentyűhozzárendelései, és csak néhány közülük működik mindenhol. Ez az oldal a teljes lista, alkalmazásonként csoportosítva, a macOS és a Windows / Linux billentyűkombinációkkal egymás mellett.

## Ezt olvasd el először

**Itt szinte semmi sem globális.** Az alábbi három billentyű az; minden, ami ezután szerepel, az éppen előtted lévő alkalmazáshoz tartozik, és lapváltáskor változik.

| Művelet                                                              | macOS | Windows / Linux |
| -------------------------------------------------------------------- | ----- | --------------- |
| A Felhasználói útmutató megnyitása — az alkalmazásban bárhol működik | F1    | F1              |
| Menüszalag összecsukása / kibontása — minden szerkesztőben           | ⌥⌘R   | Ctrl+F1         |
| Menüszalag kompakt módra váltása — csak a Docsban                    | ⌥⌘K   | Ctrl+Alt+K      |
| A Billentyűparancsok panel megnyitása — csak a Wordben               | ⌘/    | Ctrl+/          |

- **Az F1 ezt az útmutatót nyitja meg**, nem a billentyűparancsok paneljét. Ez az alkalmazásmenü része, ezért Word, Sheets, Slides, PDF, sőt a Kezdőlap lap felől is ugyanúgy működik.
- **A menüszalag megjelenítését a `Ctrl+F1` kapcsolja, macOS-en az `⌥⌘R`.** Az önmagában lenyomott `F1` nem csinál semmit — az F1 az útmutatóé, és a menüszalag billentyűje szándékosan nem osztozik rajta. A `Ctrl+F1` macOS-en is elfogadott, bár az `⌥⌘R` az a billentyű, amelyet az Office for Mac használ.
- **A Billentyűparancsok panel kizárólag a Wordé** (`⌘/`). Pontosan a lenti Word-táblázatot sorolja fel. Egyetlen más szerkesztőben sincs ilyen.

![A Word saját Billentyűparancsok panelje: balra a művelet, jobbra a billentyűkombináció, Fájl, Szerkesztés, Szöveg, Bekezdés, Beszúrás, Véleményezés és Nézet szerint csoportosítva](img/word-shortcuts.png)

- **A `⌘O`, a `⌘S` és a `⌘P` minden szerkesztő mögött külön megvalósítású**, és nem pontosan ugyanúgy viselkednek. A Word `⌘P` parancsa először oldalrendezést készít a nyomtatási előnézethez, hogy egy nyomtatott lap egy képernyőn látható oldalnak feleljen meg; a PDF-é a PDF nyomtatási párbeszédpanelét nyitja meg; a Sheets és a Slides megint csak a sajátját használja.
- **A többi alkalmazásonkénti.** Amikor a Word lap aktív, a `⌘B` a Word Félkövér parancsa. Amikor a Slides aktív, a `⌘B` a Slides Félkövér parancsa, és a menü, amelyhez tartozik, nem az, amelyhez szoktál. Ahol a hozzárendelés csak valaminek a kijelölésekor vagy egy adott panelen működik, azt a Művelet oszlop jelzi.

## Jelölések

- A macOS oszlop a megszokott szimbólumokat használja: `⌘` Command, `⌥` Option, `⇧` Shift, `⌃` Control.
- A Windows / Linux oszlop szövegesen írja őket ki: `Ctrl`, `Alt`, `Shift`.
- Ahol a GenOffice a két platformon valóban eltérő billentyűt igényel, azt a táblázat mutatja. A jellemző esetek: `⌃H` a Keresés és csere parancshoz (macOS-en a `⌘H` az Elrejtés funkcióé), `⌥⌘G` az Ugráshoz, valamint a Slides bezáró billentyűje.
- Az olyan kombinációt, amely csak kijelölés mellett működik, a Művelet oszlopban jelöljük, külön oszlop helyett, így a táblázat három oszlopos marad.

## Word

A Word teljes menüvel rendelkezik, ezért az alábbi fájl- és formázási billentyűparancsok akkor is működnek, ha egy Word-lap van fókuszban, függetlenül attól, melyik dokumentum van megnyitva. A szerkesztési billentyűparancsokhoz olyan, betöltött és szerkeszthető dokumentum kell, amelyben a kurzor a szövegben áll; nem működnek, amíg a Keresés mezőbe, az AI-promptba vagy bármely más mezőbe gépelsz.

### Fájl

| Művelet         | macOS | Windows / Linux |
| --------------- | ----- | --------------- |
| Új              | ⌘N    | Ctrl+N          |
| Új ablak        | ⇧⌘N   | Ctrl+Shift+N    |
| Megnyitás…      | ⌘O    | Ctrl+O          |
| Mentés          | ⌘S    | Ctrl+S          |
| Mentés másként… | ⇧⌘S   | Ctrl+Shift+S    |
| Nyomtatás       | ⌘P    | Ctrl+P          |
| Bezárás (lap)   | ⌘W    | Ctrl+W          |

A ⌘P nem a böngésző nyomtatása. A Word előbb nyomtatási előnézetet jelenít meg, így egy kinyomtatott lap pontosan egy képernyőoldalnak felel meg, majd megnyitja a saját nyomtatási párbeszédpaneljét.

### Szerkesztés

| Művelet                     | macOS | Windows / Linux |
| --------------------------- | ----- | --------------- |
| Visszavonás                 | ⌘Z    | Ctrl+Z          |
| Újra                        | ⇧⌘Z   | Ctrl+Shift+Z    |
| Kivágás                     | ⌘X    | Ctrl+X          |
| Másolás                     | ⌘C    | Ctrl+C          |
| Beillesztés                 | ⌘V    | Ctrl+V          |
| Beillesztés csak szövegként | ⌥⇧⌘V  | Ctrl+Shift+V    |
| Az összes kijelölése        | ⌘A    | Ctrl+A          |
| Keresés                     | ⌘F    | Ctrl+F          |
| Csere                       | ⌃H    | Ctrl+H          |
| Ugrás                       | ⌥⌘G   | Ctrl+G          |
| Billentyűparancsok          | ⌘/    | Ctrl+/          |

A Csere mindkét platformon Ctrl+H, mert a ⌘H macOS alatt elrejti az ablakot, és soha nem jut el a Wordig. Az Ugrás a Word saját felosztását követi: macOS-en ⌥⌘G, így a ⇧⌘G szabadon marad a Szavak száma számára.

### Szövegformázás

| Művelet                                        | macOS    | Windows / Linux       |
| ---------------------------------------------- | -------- | --------------------- |
| Félkövér                                       | ⌘B       | Ctrl+B                |
| Dőlt                                           | ⌘I       | Ctrl+I                |
| Aláhúzott                                      | ⌘U       | Ctrl+U                |
| Betűtípus… (párbeszédpanel)                    | ⌘D       | Ctrl+D                |
| Betűméret növelése                             | ⇧⌘.      | Ctrl+Shift+.          |
| Betűméret csökkentése                          | ⇧⌘,      | Ctrl+Shift+,          |
| Betűméret növelése 1 pt-tal                    | ⌘]       | Ctrl+]                |
| Betűméret csökkentése 1 pt-tal                 | ⌘[       | Ctrl+[                |
| Felső index                                    | ⇧⌘= / ⌘. | Ctrl+Shift+= / Ctrl+. |
| Alsó index                                     | ⌘,       | Ctrl+,                |
| Kis- és nagybetűk                              | ⇧F3      | Shift+F3              |
| Formázás másolása (kijelölt szöveggel)         | ⇧⌘C      | Ctrl+Shift+C          |
| Formázás beillesztése (Formázás másolása után) | ⇧⌘V      | Ctrl+Alt+Shift+V      |
| Összes formázás törlése                        | ⌃␣       | Ctrl+Space            |

A két formázás-másoló billentyűparancs az egyetlen hely, ahol a Windows-billentyű nem puszta átnevezés: a Formázás beillesztése Alt-ot kap, mert ott a Ctrl+Shift+V már a beillesztést és a stílus illesztését végzi.

### Bekezdésformázás

| Művelet                              | macOS | Windows / Linux |
| ------------------------------------ | ----- | --------------- |
| Balra igazítás                       | ⌘L    | Ctrl+L          |
| Középre igazítás                     | ⌘E    | Ctrl+E          |
| Jobbra igazítás                      | ⌘R    | Ctrl+R          |
| Sorkizárt                            | ⌘J    | Ctrl+J          |
| Sorköz 1,0                           | ⌘1    | Ctrl+1          |
| Sorköz 1,5                           | ⌘5    | Ctrl+5          |
| Sorköz 2,0                           | ⌘2    | Ctrl+2          |
| Behúzás növelése                     | ⌃M    | Ctrl+M          |
| Behúzás csökkentése                  | ⌃⇧M   | Ctrl+Shift+M    |
| Függő behúzás                        | ⌘T    | Ctrl+T          |
| Függő behúzás eltávolítása           | ⇧⌘T   | Ctrl+Shift+T    |
| Bekezdés formázásának törlése        | ⌃Q    | Ctrl+Q          |
| Normál (stílus)                      | ⌥⌘0   | Ctrl+Alt+0      |
| Címsor 1 (stílus)                    | ⌥⌘1   | Ctrl+Alt+1      |
| Címsor 2 (stílus)                    | ⌥⌘2   | Ctrl+Alt+2      |
| Címsor 3 (stílus)                    | ⌥⌘3   | Ctrl+Alt+3      |
| Bekezdés… (párbeszédpanel)           | ⌥⌘M   | Ctrl+Alt+M      |
| Bekezdés mozgatása felfelé           | ⌥⇧↑   | Alt+Shift+↑     |
| Bekezdés mozgatása lefelé            | ⌥⇧↓   | Alt+Shift+↓     |
| Listaelem lefokozása (elem elején)   | Tab   | Tab             |
| Listaelem előléptetése (elem elején) | ⇧Tab  | Shift+Tab       |

A Behúzás, a függő behúzás és a Bekezdés formázásának törlése mindkét platformon csak Ctrl-parancs: a ⌘M minimalizálja az ablakot, a ⌘Q pedig kilép, így egy Command-parancs észrevétlenül szerkesztené a dokumentumot kilépés közben. A Tab csak akkor vált listaszintet, ha a kurzor egy elem elején áll, vagy a kijelölés több elemre terjed ki; a szövegen belül tabulátorkaraktert szúr be, a ⇧Tab pedig mindig előléptet.

### Beszúrás

| Művelet              | macOS | Windows / Linux  |
| -------------------- | ----- | ---------------- |
| Oldaltörés           | ⌘⏎    | Ctrl+Enter       |
| Hasábtörés           | ⇧⌘⏎   | Ctrl+Shift+Enter |
| Sortörés             | ⇧⏎    | Shift+Enter      |
| Hivatkozás…          | ⌘K    | Ctrl+K           |
| Új megjegyzés        | ⌥⌘A   | Ctrl+Alt+A       |
| Lábjegyzet beszúrása | ⌥⌘F   | Ctrl+Alt+F       |
| Végjegyzet beszúrása | ⌥⌘E   | Ctrl+Alt+D       |
| Dátum (mező)         | ⌥⇧D   | Alt+Shift+D      |
| Idő (mező)           | ⌥⇧T   | Alt+Shift+T      |
| Nemtörő szóköz       | ⇧⌘␣   | Ctrl+Shift+Space |
| Nemtörő kötőjel      | ⇧⌘-   | Ctrl+Shift+-     |

A Végjegyzet beszúrása valódi platformeltérés, nem átnevezés: a Word Macen ⌥⌘E-t használ, mert ott az ⌥⌘D a Dockhoz tartozik.

### Véleményezés és eszközök

| Művelet                    | macOS | Windows / Linux |
| -------------------------- | ----- | --------------- |
| Változáskövetés            | ⇧⌘E   | Ctrl+Shift+E    |
| Szavak száma               | ⇧⌘G   | Ctrl+Shift+G    |
| Nyelvi ellenőrzés          | F7    | F7              |
| Mező frissítése            | F9    | F9              |
| Mezőkódok be-/kikapcsolása | ⌥F9   | Alt+F9          |

Az F7 lefuttatja az AI nyelvi ellenőrzését, az F9 pedig frissíti a mezőket; mindkettőhöz betöltött dokumentum kell. Ha egy szerkesztési korlátozás bekapcsolta a Változáskövetést, az már nem kapcsolható ki.

### Nézet

| Művelet                               | macOS | Windows / Linux |
| ------------------------------------- | ----- | --------------- |
| Nagyítás                              | ⌘=    | Ctrl+=          |
| Kicsinyítés                           | ⌘-    | Ctrl+-          |
| Nagyítás 100%-ra                      | ⌘0    | Ctrl+0          |
| Formázásjelek megjelenítése/elrejtése | ⌘8    | Ctrl+Shift+8    |

A formázásjelek billentyűparancsa a Word saját felosztását követi: macOS-en ⌘8, máshol Ctrl+Shift+8. A Word a nagyítást kizárólag a fenti három menüparancson keresztül kezeli — a ⌘+ itt nem csinál semmit.

Az F1 a **Felhasználói útmutatót** nyitja meg, nem ezt a listát. A billentyűparancs-lista a ⌘/ vagy a Ctrl+/, és ez a Word menüjének egyik eleme, ezért csak akkor jelenik meg, amikor egy Word-lap aktív.

## Sheets

A GenOffice Sheets az Excellel kompatibilis táblázatrács. Az alábbi billentyűparancsok többsége azok az Excel-billentyűparancsok, amelyeket már ismersz.

> **A szerkesztés elsőbbséget élvez.** Szinte minden Sheets-billentyűparancs figyelmen kívül marad, amíg cellát szerkesztesz, vagy a **szerkesztőlécen** gépelsz, hogy a gépelés, a kurzor mozgatása és a Backspace megtartsa a szövegszerkesztés megszokott jelentését. Ha egy billentyű úgy tűnik, hogy nem csinál semmit, ellenőrizd, hogy éppen szerkesztés közben vagy-e: előbb nyomd meg az Esc vagy az Enter billentyűt, majd próbáld újra. A kijelölésre ható parancsok akkor sem csinálnak semmit, ha párbeszédpanel van nyitva. Ahol egy billentyűkombináció kivételt jelent ez alól a szabály alól, azt a táblázat jelzi.

Két konvenció jellemző a Sheetsre. A `⌘` és a `⌃` ugyanazt a módosítóbillentyűt jelöli — mindkettőt minden platform elfogadja, így a `⌘1` és a `Ctrl+1` ugyanaz a kombináció. **Egy üres Windows / Linux cella azt jelenti, hogy a kombináció csak macOS-en létezik.**

### Navigáció

| Művelet                                                           | macOS      | Windows / Linux |
| ----------------------------------------------------------------- | ---------- | --------------- |
| Következő munkalap                                                | ⌘PgDn      | Ctrl+PgDn       |
| Előző munkalap                                                    | ⌘PgUp      | Ctrl+PgUp       |
| Következő munkalap (az Excel for Mac szerint)                     | ⌥→         |                 |
| Előző munkalap (az Excel for Mac szerint)                         | ⌥←         |                 |
| Ugrás az első nem rögzített cellára                               | ⌘Home      | Ctrl+Home       |
| Ugrás az utoljára használt cellára (tartalom vagy formázás)       | ⌘End       | Ctrl+End        |
| Ugrás párbeszédpanel                                              | ⌘G, F5     | Ctrl+G, F5      |
| Új munkalap beszúrása                                             | ⇧F11       | Shift+F11       |
| Az aktív cella mozgatása egy képernyőnyivel lejjebb (görgetéssel) | Page Down  | Page Down       |
| Az aktív cella mozgatása egy képernyőnyivel feljebb (görgetéssel) | Page Up    | Page Up         |
| Az aktív cella mozgatása egy képernyőnyivel jobbra (görgetéssel)  | ⌥Page Down | Alt+Page Down   |
| Az aktív cella mozgatása egy képernyőnyivel balra (görgetéssel)   | ⌥Page Up   | Alt+Page Up     |

### Kijelölés

| Művelet                                                                       | macOS           | Windows / Linux                                     |
| ----------------------------------------------------------------------------- | --------------- | --------------------------------------------------- |
| Az aktív cella teljes oszlopának kijelölése                                   | ⌃Space          | Ctrl+Space                                          |
| Az aktív cella teljes sorának kijelölése                                      | ⇧Space          | Shift+Space                                         |
| Ugrás az adatterület széléhez — a négy nyíl bármelyike                        | ⌘↑ ⌘↓ ⌘← ⌘→     | Ctrl+↑ Ctrl+↓ Ctrl+← Ctrl+→                         |
| A kijelölés kiterjesztése az adatterület széléhez; újabb lenyomás visszahúzza | ⌘⇧↑ ⌘⇧↓ ⌘⇧← ⌘⇧→ | Ctrl+Shift+↑ Ctrl+Shift+↓ Ctrl+Shift+← Ctrl+Shift+→ |
| Ugrás az aktuális sor első oszlopára                                          | Home            | Home                                                |

### Szerkesztés

| Művelet                                                                                                 | macOS                       | Windows / Linux       |
| ------------------------------------------------------------------------------------------------------- | --------------------------- | --------------------- |
| A teljes kijelölés kitöltése az aktuális bejegyzéssel                                                   | ⌃⏎ vagy ⌘⏎                  | Ctrl+Enter            |
| Új sor kezdése a cellán belül                                                                           | ⌥⏎, vagy ⌃⌥⏎ / ⌘⌥⏎ macOS-en | Alt+Enter             |
| Az aktív cella szerkesztése helyben                                                                     | F2                          | F2                    |
| Félkövér                                                                                                | ⌘B                          | Ctrl+B                |
| Dőlt                                                                                                    | ⌘I                          | Ctrl+I                |
| Aláhúzott                                                                                               | ⌘U                          | Ctrl+U                |
| Keresés                                                                                                 | ⌘F                          | Ctrl+F                |
| Kitöltés jobbra a kijelölt bal szélső oszlopából                                                        | ⌘R                          | Ctrl+R                |
| Képlethivatkozás váltása relatív és abszolút között                                                     | F4                          | F4                    |
| Képlethivatkozás váltása relatív és abszolút között (az Excel for Mac szerint, cellaszerkesztés közben) | ⌘T                          |                       |
| Másolás                                                                                                 | ⌘C                          | Ctrl+C                |
| Kivágás                                                                                                 | ⌘X                          | Ctrl+X                |
| Beillesztés                                                                                             | ⌘V                          | Ctrl+V                |
| Csak az értékek beillesztése                                                                            | ⌘⇧V                         | Ctrl+Shift+V          |
| Villámkitöltés a fenti minta alapján                                                                    | ⌘E                          | Ctrl+E                |
| Visszavonás                                                                                             | ⌘Z                          | Ctrl+Z                |
| Újra                                                                                                    | ⇧⌘Z                         | Ctrl+Shift+Z          |
| A teljes munkalap kijelölése                                                                            | ⌘A                          | Ctrl+A                |
| A mai dátum beszúrása                                                                                   | ⌘;                          | Ctrl+;                |
| Az aktuális idő beszúrása                                                                               | ⌘⇧;                         | Ctrl+Shift+;          |
| Az összes kijelölt cella tartalmának törlése                                                            | ⌫ vagy ⌦                    | Backspace vagy Delete |

Az AutoSzum csak macOS-en érhető el. Nyomd meg a ⌘⇧T kombinációt.

### Formázás

| Művelet                                                                      | macOS | Windows / Linux |
| ---------------------------------------------------------------------------- | ----- | --------------- |
| Cellák formázása párbeszédpanel                                              | ⌘1    | Ctrl+1          |
| Áthúzott (kijelölt cellák esetén)                                            | ⌘5    | Ctrl+5          |
| Betűméret növelése (kijelölt cellák esetén)                                  | ⌘⇧.   | Ctrl+Shift+.    |
| Betűméret csökkentése (kijelölt cellák esetén)                               | ⌘⇧,   | Ctrl+Shift+,    |
| Számformátum: Általános                                                      | ⌘⇧`   | Ctrl+Shift+`    |
| Számformátum: két tizedesjegyű szám                                          | ⌘⇧1   | Ctrl+Shift+1    |
| Számformátum: idő `h:mm AM/PM`                                               | ⌘⇧2   | Ctrl+Shift+2    |
| Számformátum: dátum `d-mmm-yy`                                               | ⌘⇧3   | Ctrl+Shift+3    |
| Számformátum: Pénznem                                                        | ⌘⇧4   | Ctrl+Shift+4    |
| Számformátum: Százalék                                                       | ⌘⇧5   | Ctrl+Shift+5    |
| Számformátum: Tudományos                                                     | ⌘⇧6   | Ctrl+Shift+6    |
| Minden szegély (kijelölt cellák esetén)                                      | ⌘⇧7   | Ctrl+Shift+7    |
| Nincs szegély (kijelölt cellák esetén)                                       | ⌘⇧-   | Ctrl+Shift+-    |
| Formátumfestő megszakítása (csak akkor, ha a Formátumfestő be van kapcsolva) | Esc   | Esc             |

### Beszúrás és törlés

| Művelet                                                                                                                | macOS                                         | Windows / Linux                                  |
| ---------------------------------------------------------------------------------------------------------------------- | --------------------------------------------- | ------------------------------------------------ |
| Cellák beszúrása — teljes sorok vagy oszlopok közvetlenül beszúródnak, minden más megnyitja a Beszúrás párbeszédpanelt | ⌘⇧=                                           | Ctrl+Shift+=                                     |
| Cellák beszúrása — numerikus billentyűzet `+`, ugyanaz a párbeszédpanel                                                | ⌘ a numerikus billentyűzet `+` billentyűjével | Ctrl a numerikus billentyűzet `+` billentyűjével |
| Cellák törlése — teljes sorok vagy oszlopok közvetlenül törlődnek, minden más megnyitja a Törlés párbeszédpanelt       | ⌘-                                            | Ctrl+-                                           |
| Cellák törlése — numerikus billentyűzet `-`, ugyanaz a párbeszédpanel                                                  | ⌘ a numerikus billentyűzet `-` billentyűjével | Ctrl a numerikus billentyűzet `-` billentyűjével |
| A kijelölt sorok elrejtése                                                                                             | ⌘9                                            | Ctrl+9                                           |
| A kijelölésben lévő rejtett sorok felfedése                                                                            | ⌘⇧9                                           | Ctrl+Shift+9                                     |
| A kijelölt oszlopok elrejtése                                                                                          | ⌘0                                            | Ctrl+0                                           |
| A kijelölésben lévő rejtett oszlopok felfedése                                                                         | ⌘⇧0                                           | Ctrl+Shift+0                                     |

### Adatok és véleményezés

| Művelet                                                                    | macOS | Windows / Linux |
| -------------------------------------------------------------------------- | ----- | --------------- |
| Függvény beszúrása párbeszédpanel                                          | ⇧F3   | Shift+F3        |
| Névkezelő                                                                  | ⌃F3   | Ctrl+F3         |
| Jegyzet hozzáadása vagy szerkesztése az aktív cellán                       | ⇧F2   | Shift+F2        |
| Hivatkozás beszúrása                                                       | ⌘K    | Ctrl+K          |
| Hivatkozott cellák nyomon követése (a cellának képletet kell tartalmaznia) | ⌘[    | Ctrl+[          |
| Függő cellák nyomon követése                                               | ⌘]    | Ctrl+]          |
| A kijelölt sorok csoportosítása                                            | ⌥⇧→   | Alt+Shift+→     |
| A kijelölt sorok csoportbontása                                            | ⌥⇧←   | Alt+Shift+←     |
| A teljes munkafüzet újraszámítása                                          | F9    | F9              |
| Csak az aktív munkalap újraszámítása                                       | ⇧F9   | Shift+F9        |
| A szűrő be- vagy kikapcsolása a kijelölt tartományon                       | ⌘⇧F   |                 |

Az Excel szűrési billentyűparancsa, a Ctrl+Shift+L, nem csinál semmit Windows vagy Linux alatt. macOS-en nyomd meg a ⌘⇧F kombinációt, vagy bármely platformon használd a szűrőparancsot a menüszalagon.

### Ablak

| Művelet                                                                     | macOS        | Windows / Linux |
| --------------------------------------------------------------------------- | ------------ | --------------- |
| A munkafüzet mentése                                                        | ⌘S           | Ctrl+S          |
| Mentés másként                                                              | ⇧⌘S          | Ctrl+Shift+S    |
| Munkafüzet megnyitása                                                       | ⌘O           | Ctrl+O          |
| Nyomtatás                                                                   | ⌘P           | Ctrl+P          |
| A munkafüzet bezárása (macOS) / Kilépés az alkalmazásból (Windows és Linux) | ⌘W           | Ctrl+Q          |
| A Felhasználói útmutató megnyitása                                          | F1           | F1              |
| Nagyítás                                                                    | ⌘=           | Ctrl+=          |
| Nagyítás az egérgörgővel                                                    | ⌘ + görgetés | Ctrl + görgetés |
| A képletek megjelenítése vagy elrejtése a cellákban                         | ⌘`           | Ctrl+`          |
| A szerkesztőléc kibontása vagy összecsukása                                 | ⌘⇧U          | Ctrl+Shift+U    |

## Slides

A GenOffice Slides egy PowerPoint-stílusú bemutatószerkesztő. A legtöbb billentyű bárhol ugyanazt teszi, de néhányé a fókusztól függően változik — szövegdobozon belül, illetve ha a **miniatűrpanel**, a **vázlatpanel** vagy a **diarendező** kapja a billentyűfókuszt. Ezek az esetek a Művelet oszlopban vannak megjelölve.

### Fájl és alkalmazás

| Művelet                      | macOS | Windows / Linux |
| ---------------------------- | ----- | --------------- |
| Mentés                       | ⌘S    | Ctrl+S          |
| Mentés másként               | ⇧⌘S   | Ctrl+Shift+S    |
| Megnyitás                    | ⌘O    | Ctrl+O          |
| Aktuális lap bezárása        | ⌘W    | Ctrl+Q          |
| Nyomtatás                    | ⌘P    | Ctrl+P          |
| Visszavonás                  | ⌘Z    | Ctrl+Z          |
| Újra                         | ⇧⌘Z   | Ctrl+Shift+Z    |
| Újra (alternatív kombináció) | ⌘Y    | Ctrl+Y          |

A lap bezárása az egyetlen pont, ahol a két platform teljesen szétválik: macOS-en ⌘W, Windowson és Linuxon pedig Ctrl+Q, amely az egész alkalmazást zárja be, nem csak a lapot.

### Szerkesztés

| Művelet                                               | macOS  | Windows / Linux |
| ----------------------------------------------------- | ------ | --------------- |
| Új dia                                                | ⇧⌘N    | Ctrl+M          |
| Keresés (alakzat szövegének szerkesztésekor nem)      | ⌘F     | Ctrl+F          |
| AI kérdezése a kijelölésről (alakzatok kijelölésekor) | ⌘K     | Ctrl+K          |
| Nagyítás                                              | ⌘=     | Ctrl+=          |
| Kicsinyítés                                           | ⌘-     | Ctrl+-          |
| Tényleges méret (100%)                                | ⌘0     | Ctrl+0          |
| Az összes kijelölése                                  | ⌘A     | Ctrl+A          |
| Következő / előző alakzat kijelölése                  | ⇥ / ⇧⇥ | Tab / Shift+Tab |

Az új dia a másik platformfüggő eset. Mindkét platformnak megvan a saját billentyűje, a másik pedig semmit sem csinál: macOS-en ⇧⌘N, Windowson és Linuxon Ctrl+M. A ⌘M macOS-en nem láthatja el ezt a feladatot, mert az a rendszer ablakminimalizáló billentyűje.

„Az összes kijelölése” követi a fókuszt: ha a miniatűrpanel, a vázlatpanel vagy a diarendező van fókuszban, akkor az összes diát kijelöli, bárhol másutt pedig az aktuális dia összes elemét. A Tab a dia alakzatai között lépked, és csak akkor, ha maga a diavászon van fókuszban.

### Szövegformázás

Ezek a kijelölt alakzat szövegére, illetve ha éppen a szövegdobozban vagy, a kurzor bekezdésére hatnak.

| Művelet                                               | macOS | Windows / Linux |
| ----------------------------------------------------- | ----- | --------------- |
| Félkövér                                              | ⌘B    | Ctrl+B          |
| Dőlt                                                  | ⌘I    | Ctrl+I          |
| Aláhúzott                                             | ⌘U    | Ctrl+U          |
| Balra igazítás                                        | ⌘L    | Ctrl+L          |
| Középre igazítás                                      | ⌘E    | Ctrl+E          |
| Jobbra igazítás                                       | ⌘R    | Ctrl+R          |
| Sorkizárt                                             | ⌘J    | Ctrl+J          |
| Betűméret növelése 1 ponttal                          | ⌘]    | Ctrl+]          |
| Betűméret csökkentése 1 ponttal                       | ⌘[    | Ctrl+[          |
| Betűméret növelése a következő előre megadott méretre | ⇧⌘>   | Ctrl+Shift+>    |
| Betűméret csökkentése az előző előre megadott méretre | ⇧⌘<   | Ctrl+Shift+<    |

A betűméret-léptető billentyűk a . és , billentyűkkel is működnek, Shift lenyomva tartásával — azokon a billentyűzetkiosztásokon, amelyek az írásjelbillentyűk helyett ezeket jelentik.

### Vágólap és objektumok

| Művelet                                                   | macOS       | Windows / Linux      |
| --------------------------------------------------------- | ----------- | -------------------- |
| Másolás                                                   | ⌘C          | Ctrl+C               |
| Kivágás                                                   | ⌘X          | Ctrl+X               |
| Beillesztés                                               | ⌘V          | Ctrl+V               |
| Csak formázás beillesztése                                | ⇧⌘V         | Ctrl+Shift+V         |
| Csak formázás másolása (alakzatok kijelölésekor)          | ⇧⌘C         | Ctrl+Shift+C         |
| Duplikálás a helyén (alakzatok kijelölésekor)             | ⌘D          | Ctrl+D               |
| Csoportosítás (alakzatok kijelölésekor)                   | ⌘G          | Ctrl+G               |
| Csoportbontás (alakzatok kijelölésekor)                   | ⇧⌘G         | Ctrl+Shift+G         |
| Kijelölés törlése (alakzatok kijelölésekor)               | ⌫ / Delete  | Backspace / Delete   |
| A kijelölt csúcspont törlése (Pontok szerkesztése módban) | ⌫ / Delete  | Backspace / Delete   |
| Kijelölés mozgatása (alakzatok kijelölésekor)             | ↑ ↓ ← →     | Nyílbillentyűk       |
| Mozgatás egy képernyőpixellel                             | ⌘↑ ⌘↓ ⌘← ⌘→ | Ctrl+Nyílbillentyűk  |
| Átméretezés a középpont körül                             | ⇧↑ ⇧↓ ⇧← ⇧→ | Shift+Nyílbillentyűk |

Egy szövegdobozban a másolás, a kivágás és a beillesztés az alakzat helyett a szövegre hat. A Shift a nyílbillentyűkkel együtt minden alakzatot a saját középpontja körül méretez át, az összekötőket nem bántja, a többszörös kijelölést pedig egyetlen visszavonási lépésként rögzíti.

### Diák — miniatűrpanel, vázlatpanel vagy diarendező fókuszban

A PowerPoint-féle diaszintű parancsok csak akkor működnek, ha e három panel valamelyike birtokolja a billentyűfókuszt.

| Művelet                                        | macOS                              | Windows / Linux                                    |
| ---------------------------------------------- | ---------------------------------- | -------------------------------------------------- |
| Előző / következő dia                          | ↑ ↓ (valamint Page Up / Page Down) | ArrowUp / ArrowDown (valamint Page Up / Page Down) |
| Első / utolsó dia                              | Home / End                         | Home / End                                         |
| Kijelölés kiterjesztése az első / utolsó diáig | ⇧Home / ⇧End                       | Shift+Home / Shift+End                             |
| Az összes dia kijelölése                       | ⌘A                                 | Ctrl+A                                             |
| Kijelölt diák törlése                          | ⌫ / Delete                         | Backspace / Delete                                 |
| Diák másolása                                  | ⌘C                                 | Ctrl+C                                             |
| Diák kivágása                                  | ⌘X                                 | Ctrl+X                                             |

A nyílbillentyűkkel történő mozgás a diák között a kivétel — az a vászonról is működik, ha semmi nincs kijelölve. A diák törlése nem érhető el, amíg a toll eszköz aktív, vagy az **Olvasó nézetben** vagy.

### Szövegdobozon belül

| Művelet                                                                     | macOS                                     | Windows / Linux                               |
| --------------------------------------------------------------------------- | ----------------------------------------- | --------------------------------------------- |
| Szövegszerkesztés megkezdése a kijelölt szöveges alakzaton                  | bármilyen nyomtatható karakter, ⏎ vagy F2 | bármilyen nyomtatható karakter, Enter vagy F2 |
| A szöveg véglegesítése és az alakzat újrakijelölése                         | Esc vagy F2                               | Esc vagy F2                                   |
| A szöveg véglegesítése és az alakzat újrakijelölése (alternatív kombináció) | ⌘⏎                                        | Ctrl+⏎                                        |
| Ugrás a következő helyőrzőre                                                | ⌃⏎                                        | Ctrl+⏎                                        |
| Listaszint csökkentése / növelése                                           | ⇥ / ⇧⇥                                    | Tab / Shift+Tab                               |
| Következő / előző táblázatcella                                             | ⇥ / ⇧⇥                                    | Tab / Shift+Tab                               |
| Kézi sortörés                                                               | ⇧⏎                                        | Shift+⏎                                       |

A Tab két különböző dolgot jelent: táblázaton belül a következő cellára lép, szöveges alakzaton belül pedig a listaszintet csökkenti.

### Diavetítés

| Művelet                            | macOS               | Windows / Linux                      |
| ---------------------------------- | ------------------- | ------------------------------------ |
| Indítás az elejétől                | F5                  | F5                                   |
| Indítás az aktuális diától         | ⇧F5 vagy ⌘⏎         | ⇧F5                                  |
| Következő dia                      | → ↓ ␣ ⏎ Page Down   | Jobbra, Le, Szóköz, Enter, Page Down |
| Előző dia                          | ← ↑ Page Up         | Balra, Fel, Page Up                  |
| Első dia                           | Home                | Home                                 |
| Utolsó dia                         | End                 | End                                  |
| Ugrás az N. diára                  | N beírása, majd ⏎   | N beírása, majd Enter                |
| Fekete képernyő                    | B vagy .            | B vagy .                             |
| Fehér képernyő                     | W vagy ,            | W vagy ,                             |
| Visszaállítás fekete képernyő után | bármelyik billentyű | bármelyik billentyű                  |
| A vetítés befejezése               | Esc                 | Esc                                  |

macOS-en az aktuális diáról való indítás ⇧F5 mellett ⌘⏎ kombinációval is működik. Windowson ez a billentyű továbbra is a helyőrzők közötti mozgásra van fenntartva.

### Olvasó nézet

| Művelet                    | macOS             | Windows / Linux                      |
| -------------------------- | ----------------- | ------------------------------------ |
| Következő dia              | → ↓ ␣ ⏎ Page Down | Jobbra, Le, Szóköz, Enter, Page Down |
| Előző dia                  | ← ↑ Page Up       | Balra, Fel, Page Up                  |
| Kilépés az Olvasó nézetből | Esc               | Esc                                  |

Az **Esc** billentyű mindig kiléptet abból, amiben éppen vagy — a képvágásból, az éppen rajzolt alakzatból, a **Pontok szerkesztéséből**, egy csoportból, a toll, a kiemelő vagy a radír eszközből, a médialejátszóból vagy az **Olvasó nézetből** —, és megszünteti a kijelölést.

Két további billentyű nem kombináció, hanem módosítóbillentyű: rajzolás közben tartsd lenyomva a **Shift** billentyűt, hogy az új alakzat megtartsa az arányait, forgatás közben pedig ugyancsak tartsd lenyomva a **Shift** billentyűt, hogy a forgatás 15°-os lépésekre illeszkedjen.

## PDF

A PDF-nek nincs saját menüje, ezért az alábbi összes billentyűparancs az aktív laphoz tartozik, és csak akkor működik, ha elöl egy PDF-lap áll.

### Fájl és szerkesztés

| Művelet                     | macOS | Windows / Linux |
| --------------------------- | ----- | --------------- |
| Mentés                      | ⌘S    | Ctrl+S          |
| Visszavonás                 | ⌘Z    | Ctrl+Z          |
| Újra                        | ⇧⌘Z   | Ctrl+Shift+Z    |
| Keresés (kereső megnyitása) | ⌘F    | Ctrl+F          |
| Nyomtatás                   | ⌘P    | Ctrl+P          |

A Visszavonás és az Újra a dokumentum előzményeiben lépked; ha a kurzor éppen egy szövegmezőben áll, a ⌘Z a mezőé marad. A ⌘P ugyanazt az alkalmazáson belüli nyomtatási párbeszédpanelt nyitja meg, mint az eszköztár Nyomtatás gombja.

### Nézet

| Művelet               | macOS            | Windows / Linux     |
| --------------------- | ---------------- | ------------------- |
| Nagyítás              | ⌘=               | Ctrl+=              |
| Kicsinyítés           | ⌘-               | Ctrl+-              |
| Szélességhez igazítás | ⌘0               | Ctrl+0              |
| Nagyítás (folyamatos) | ⌘ + görgetőkerék | Ctrl + görgetőkerék |

A Nagyítás a ⌘+ és a Ctrl++ billentyűkre is működik.

### Lapozás

| Művelet                          | macOS               | Windows / Linux     |
| -------------------------------- | ------------------- | ------------------- |
| Következő oldal                  | →                   | →                   |
| Előző oldal                      | ←                   | ←                   |
| Egy képernyőnyi görgetés lefelé  | PageDown vagy Space | PageDown vagy Space |
| Egy képernyőnyi görgetés felfelé | PageUp              | PageUp              |
| Görgetés lefelé                  | ↓                   | ↓                   |
| Görgetés felfelé                 | ↑                   | ↑                   |
| Ugrás az első oldalra            | Home                | Home                |
| Ugrás az utolsó oldalra          | End                 | End                 |

Ezekhez nem tartozik módosítóbillentyű, így mindkét platformon ugyanúgy működnek. A ↓ és az ↑ az egyetlen sor, amely a fókusztól függ: a dokumentumban csak kicsit görgetnek, a miniatűr oldalsávon viszont egy teljes oldalt lépnek. Amíg egy mezőbe írsz, nem csinálnak semmit.

### Kijelölés és szerkesztés

| Művelet                      | macOS                 | Windows / Linux       |
| ---------------------------- | --------------------- | --------------------- |
| Kijelölt széljegyzet törlése | Delete vagy Backspace | Delete vagy Backspace |

Csak akkor aktív, ha széljegyzet, bélyegző, kitakarás vagy kép van kijelölve.

### Escape és párbeszédpanelek

| Művelet                             | macOS  | Windows / Linux |
| ----------------------------------- | ------ | --------------- |
| A legfelső állapot megszakítása     | Escape | Escape          |
| Nyitott párbeszédpanel megerősítése | Enter  | Enter           |

Az Escape egyszerre csak egy réteget bont le — először a nyitott párbeszédpanelt, aztán az AI felugró paneljét, egy szövegtervezetet, a keresési panelt, végül magát a kijelölést. Ha a kurzor a keresőmezőben áll, az Enter és a ⇧Enter a következő, illetve az előző találatra lép, az Escape bezárja a panelt; ha a kurzor egy új szöveges jegyzetben áll, a ⌘⏎ véglegesíti azt.

## Markdown

A GenOffice Markdown a Markdownot szerkeszti, mellette élő előnézettel.

### Fájl és nézet

| Művelet                      | macOS      | Windows / Linux    |
| ---------------------------- | ---------- | ------------------ |
| Mentés                       | ⌘S         | Ctrl+S             |
| Mentés másként               | ⇧⌘S        | Ctrl+Shift+S       |
| Nyomtatás                    | ⌘P         | Ctrl+P             |
| Keresés                      | ⌘F         | Ctrl+F             |
| Forrás / előnézet kapcsolása | ⌘E         | Ctrl+E             |
| Nagyítás                     | ⌘= vagy ⌘+ | Ctrl+= vagy Ctrl++ |
| Kicsinyítés                  | ⌘-         | Ctrl+-             |
| Tényleges méret (100%)       | ⌘0         | Ctrl+0             |

A Kicsinyítés a ⌘_ és a Ctrl+_ kombinációval is működik. A trackpad csippentése, vagy a ⌘ illetve a Ctrl nyomva tartása melletti görgetés a dokumentum fölött szintén átméretezi a nézetet.

### Blokkok

| Művelet                                                           | macOS | Windows / Linux      |
| ----------------------------------------------------------------- | ----- | -------------------- |
| Kijelölt blokkok duplikálása                                      | ⌘D    | Ctrl+D               |
| Kijelölt blokkok törlése                                          | ⇧⌘⌫   | Ctrl+Shift+Backspace |
| Kijelölt blokk mozgatása felfelé                                  | ⇧⌘↑   | Ctrl+Shift+ArrowUp   |
| Kijelölt blokk mozgatása lefelé                                   | ⇧⌘↓   | Ctrl+Shift+ArrowDown |
| Hivatkozás beszúrása a kijelölésre                                | ⌘K    | Ctrl+K               |
| Hivatkozás eltávolítása (nyomd meg újra, amíg a hivatkozás aktív) | ⌘K    | Ctrl+K               |

### Matematikai blokkon belül

Ezek csak akkor érvényesek, amíg a lebegő LaTeX-szerkesztő nyitva van.

| Művelet                               | macOS | Windows / Linux |
| ------------------------------------- | ----- | --------------- |
| Képlet alkalmazása                    | ⌘⏎    | Ctrl+⏎          |
| Szerkesztő bezárása alkalmazás nélkül | Esc   | Esc             |

Üres képlet alkalmazása esetén a blokk törlődik. A szerkesztő a dokumentum bármilyen módosításakor is bezárul, és a szerkesztőn kívülre kattintva is eltűnik.

## HTML

A GenOffice HTML egy kódszerkesztőt mutat élő előnézet mellett. Az oldalelemekre ható billentyűk csak akkor működnek, ha maga az előnézet kapta a billentyűfókuszt.

### Fájl és forrásnézet

| Művelet                                          | macOS      | Windows / Linux    |
| ------------------------------------------------ | ---------- | ------------------ |
| Mentés                                           | ⌘S         | Ctrl+S             |
| Mentés másként                                   | ⇧⌘S        | Ctrl+Shift+S       |
| Nézet váltása (szerkesztés / osztott / előnézet) | ⌘\         | Ctrl+\             |
| Keresés                                          | ⌘F         | Ctrl+F             |
| Nagyítás                                         | ⌘= vagy ⌘+ | Ctrl+= vagy Ctrl++ |
| Kicsinyítés                                      | ⌘-         | Ctrl+-             |
| Tényleges méret (100%)                           | ⌘0         | Ctrl+0             |

A Kicsinyítés a ⌘_ és a Ctrl+_ kombinációval is működik.

### Formázás és előzmények

| Művelet                       | macOS | Windows / Linux |
| ----------------------------- | ----- | --------------- |
| Félkövér                      | ⌘B    | Ctrl+B          |
| Dőlt                          | ⌘I    | Ctrl+I          |
| Visszavonás                   | ⌘Z    | Ctrl+Z          |
| Újra                          | ⇧⌘Z   | Ctrl+Shift+Z    |
| Újra (másik billentyűparancs) | ⌘Y    | Ctrl+Y          |

A kódszerkesztőben vagy egy űrlapmezőben a mező saját szerkesztőbillentyűi lépnek érvénybe helyette.

### Az előnézeti keretben, kijelölt elemmel

| Művelet                                | macOS      | Windows / Linux    |
| -------------------------------------- | ---------- | ------------------ |
| AI kérdezése az elemről                | ⌘K         | Ctrl+K             |
| Elem törlése                           | ⌫ / Delete | Backspace / Delete |
| Előző testvérelem                      | ↑          | ArrowUp            |
| Következő testvérelem                  | ↓          | ArrowDown          |
| Szülő kijelölése                       | ←          | ArrowLeft          |
| Első gyermekelem kijelölése            | →          | ArrowRight         |
| Elem mozgatása felfelé a dokumentumban | ⌥↑         | Alt+ArrowUp        |
| Elem mozgatása lefelé a dokumentumban  | ⌥↓         | Alt+ArrowDown      |
| Kijelölés megszüntetése                | Esc        | Esc                |

Mindegyikhez már ki kell legyen jelölve egy elem; ha semmi sincs kijelölve, a billentyű nem csinál semmit.

### Az előnézeti keretben, szöveg szerkesztése a helyén

| Művelet                                            | macOS | Windows / Linux |
| -------------------------------------------------- | ----- | --------------- |
| Szerkesztés véglegesítése                          | Esc   | Esc             |
| Szerkesztés véglegesítése (másik billentyűparancs) | ⌘⏎    | Ctrl+⏎          |
| Kijelölt rész félkövérré tétele                    | ⌘B    | Ctrl+B          |
| Kijelölt rész dőltté tétele                        | ⌘I    | Ctrl+I          |

A véglegesítés megtartja, amit beírtál, és az elem visszatér a kijelölt állapotba.
