# Docs: szövegszerkesztés

A Docs a Wordhöz hasonló szövegszerkesztő: valódi .docx fájlokat olvas és ír, valódi WYSIWYG oldaltördeléssel.

## A menüszalag

Lapok: **Kezdőlap / Beszúrás / Rajzolás / Elrendezés / Tervezés / Hivatkozások / Véleményezés / Nézet**, továbbá környezetfüggő lapok a kijelölt objektumhoz — táblázatnál a Táblázattervezés és az Elrendezés, képnél a Képformátum, alakzatnál az Alakzatformátum, élőfej vagy élőláb szerkesztésekor pedig az Élőfej és élőláb. A menüszalagnak a Docsban saját sűrűségi szintje is van: a **kompakt** (⌥⌘K / Ctrl+Alt+K) meghagyja az ikonokat, a feliratokat elhagyja; a választás megmarad.

- **Kezdőlap**: vágólap; betűtípus (CJK-méretek és kiemelőjelek is) az **Összes formázás törlése** paranccsal és a **Formázási jelek megjelenítése/elrejtése** kapcsolóval; bekezdés (igazítás/behúzás/sorköz/listák), valamint az **Új felsorolásjel definiálása / Új számozási formátum definiálása / Új többszintű lista definiálása**, amelyek a saját listastílusaidat a dokumentumba mentik; stílusok (Címsor 1–6 / Normál / Idézet, módosíthatók) a **Stílusok panellel** a teljes listához. Egy stílus a kelet-ázsiai és a latin betűtípust külön is megadhatja (például egy CJK betűkészlet mellé Times New Roman); a páros a .docx mentése és újranyitása után is megmarad.
- **Beszúrás**: oldal- és szakasztörések; táblázatok (sor × oszlop rács, vagy **Táblázat beszúrása…** a pontos mérethez); képek, alakzatok, szövegdobozok; **Címlap** és **Üres oldal** előre elkészített galériából; **Diagram**; **Iniciálé**; **WordArt**; mezők (dátum, idő, oldal, oldalak száma, fájlnév); hivatkozások, **Könyvjelző** és kereszthivatkozások; megjegyzések; élőfej/élőláb és oldalszámok; szimbólumok és egyenletek.
  - A **Diagram** valódi diagramobjektumot szúr be saját adatokkal — oszlop-, vonal- vagy kördiagramot —, nem képet. A Wordből ismert _Adatok szerkesztése_ megnyitja a mögötte lévő számokat.
- **Rajzolás**: tintázás az oldalon, a **Rajzeszközök** csoportban — a **Kijelölés** visz vissza a szövegszerkesztéshez, utána jön a **Toll**, a **Kiemelő** és a **Radír** (egy kattintás vagy végighúzás a teljes vonást törli). Mellettük a **Tollstílus** / **Kiemelőstílus** egyetlen vezérlő, amely színmintákat és egy sor vonalvastagságot tartalmaz; a felirata az éppen aktív eszközt követi. A tintát a dokumentum a szöveg fölött lebegő széljegyzetként tárolja, így mentés és újranyitás után is megmarad, a következő csoport **Összes törlése** parancsa pedig mindet eltávolítja.
- **Elrendezés**: margók, tájolás és papírméret, hasábok, bekezdésbehúzások és sorközök.
- **Tervezés**: témák, színkészletek, vízjel, oldalszegélyek.
- **Hivatkozások**: tartalomjegyzék (frissíthető), lábjegyzetek/végjegyzetek, képaláírások, kereszthivatkozások.

  ![A Hivatkozások lap](img/docs-references.png)

- **Véleményezés**: a a **Szerkesztő** az egész dokumentumot átnézi helyesírás, nyelvtan és központozás szempontjából; **Fordítás**; helyesírás-ellenőrzés; megjegyzések (az **AI-megjegyzések megoldása** végigmegy a nyitottakon); változáskövetés Összes jelölés / Egyszerű jelölés nézettel, elfogadás/elutasítás, és a **Módosítások AI-összefoglalása**; szavak száma; **Összehasonlítás** másik fájllal; **Dokumentum védelme**.
- **Nézet**: öt mód a fájl megtekintésére — **Nyomtatási elrendezés**, **Webes elrendezés**, **Vázlat**, **Olvasó mód** és **Oldal-előnézet**; kicsinyítés/nagyítás/100%/oldalszélesség/egy oldal; **AI-panel**; **Sötét mód**; vonalzó, rácsvonalak és a navigációs panel; **Új lap**, **Felosztás** és **Váltás lapok között**; a kereshető **billentyűparancs-ablak**.
  - A **Sötét mód** besötétíti az oldalt és a körülötte lévő vásznat, de sosem a menüszalagot — ugyanúgy, ahogy a Wordben is különválik a sötét szerkesztőfelület és a sötét ablak. A választás megmarad, és minden esetben elsőbbséget élvez az alkalmazás témájával szemben.
  - A **Felosztás** egy második panelt nyit alul, amely önállóan görget, és az elsőt tükrözi; a szélén lévő ×-szel zárható be.

## A navigációs panel

A **Nézet ▸ Navigációs panel** egy oldalsávot nyit meg, amely a dokumentum címsorvázlatát, az egész dokumentumra kiterjedő keresőmezőt és oldalanként egy miniatűrt tartalmaz. Az, hogy nyitva van-e, az indítások között is megmarad, így a vázlat alapján böngészett dokumentum a következő megnyitáskor is így böngészhető.

A **vázlat** a címsorfa. Kattints benne jobb gombbal egy címsorra, hogy ne csak navigálj, hanem összecsukd és átrendezd a szerkezetet:

- Az **Összecsukás / Kibontás** egy címsoron a címsor teljes részfáját összecsukja vagy kibontja — a fejezet eltűnik a vázlatból, a szövege a dokumentumban marad.
- Az **Összes összecsukása / Összes kibontása** egy csapásra az egészet összecsukja vagy kibontja. Egy hosszú jelentésnél ez a különbség az áttekinthető vázlat és a szövegtömb között.
- A **Címsorszintek megjelenítése** a fát arra a mélységre szűkíti, amelyik érdekel, így a _Címsorszint megjelenítése: 1_ olyan tartalomjegyzéket hagy meg, amelyet tényleg át tudsz tekinteni.
- A **Szint emelése / Szint csökkentése** módosítja a címsor szintjét, és vele együtt az alatta lévő összes címsor szintjét is — így lesz egy fejezetből alfejezet.
- Az **Új címsor beszúrása elé / mögé** egyet szúr be a kurzornál, a panel elhagyása nélkül.
- A **Törlés** a címsort _és mindent, ami alatta van_ eltávolítja, ezért ezzel óvatosan: ez egy részfa törlése, nem egy soré.
- A **Címsor és tartalom kijelölése** a címsortól a részfája végéig jelöl ki, így egy egész szakaszt egyben szerkeszthetsz.

## Helyi menü

Kattints jobb gombbal bárhol a szövegtörzsben — a menü attól függ, amire kattintottál. Fő csoportok:

- **Vágólap**: kivágás / másolás / beillesztés / **beillesztés egyszerű szövegként**.
- **Betűtípus, bekezdés**: betűtípus és méret, félkövér/dőlt/aláhúzott, igazítás/behúzás/sorköz módosítása a menüszalag használata nélkül.
- **Szinonimák**: felsorolja a kijelölt szó szinonimáit; kattints az egyikre a cseréhez.
- **Fordítás** (AI): a kijelölés lefordítása a célnyelvre (angol, egyszerűsített kínai, japán, koreai, francia, német, spanyol, ...) az AI-panelen keresztül.
- **Új megjegyzés**: megjegyzés csatolása a kijelöléshez.
- **Helyesírás** (elírt szónál): javasolt javítások, az összes kihagyása, felvétel a szótárba, a nyelvi ellenőrzés nyelvének beállítása.
- **Hivatkozás**: megnyitás / szerkesztés / hivatkozás másolása / hivatkozás eltávolítása.
- **Kép**: kép megtekintése, kép mentése másként, **szövegtördelés** (szöveggel egy sorban / négyzet / felül és alul / szöveg mögött / szöveg előtt), elrendezési sorrend.
- **Mezők** (tartalomjegyzék, oldalszámok): mező frissítése / mezőkódok megjelenítése vagy elrejtése / mező szerkesztése.
- **Listaszámozás** (listán belül): számozás újraindítása / számozás folytatása / listaszint módosítása / számozási érték megadása.
- **Táblázat** (kurzor egy táblázatban): sorok/oszlopok beszúrása, cellák egyesítése / felosztása, táblázat felosztása, automatikus illesztés, cellaigazítás, sorok/oszlopok elosztása, táblázattulajdonságok, a Törlés almenü, kijelölés.

## Szerkesztés

- Keresés és csere (ctrl+F / ctrl+H): kis- és nagybetűk megkülönböztetése, csak teljes szavak, reguláris kifejezés.
- Formátummásoló; többszintű visszavonás és újra; beillesztési beállítások.
- Táblázatok: cellák egyesítése/felosztása, sor-/oszlopműveletek, szegélyek és árnyékolás, rendezés, képletek. A sorok csak ott szakadnak oldalak közé, ahol kell, és egy cellán belül a Ctrl+Enter pontosan ott töri az oldalt — a nyomtatás és a PDF is ezt követi.
- Képek: szövegtördelés, körbevágás, tömörítés; rajzvászon.

## CJK tipográfia

- Az írásjel-tömörítés és a kinsoku sortörés megegyezik a Wordével; teljes és fél szélességű karakterek közötti átalakítás.
- A betűtípus-jelöltek lefedik a Windows és a macOS gyakori CJK betűtípus-neveit.

## AI

- AI-gomb a menüszalagon és az oldalsáv: átfogalmazás, bővítés, fordítás, összefoglalás, táblázatbeszúrási előbeállítások, valamint szabad szöveges utasítások.
- Minden AI-kör előtt pillanatkép készül; a verziólistáról visszaállhatsz, és maga a visszaállítás is visszavonható.

## Mentés és exportálás

- A .docx mentésekor csak a módosított bekezdéseket írja újra — az érintetlen tartalom bájtra azonos marad.
- PDF-be exportál (az oldaltördelés megtartásával), valamint oldalankénti képekbe.

## Nyomtatás

ctrl+P a rendszer nyomtatási párbeszédpaneljén keresztül, WYSIWYG oldalakkal.
