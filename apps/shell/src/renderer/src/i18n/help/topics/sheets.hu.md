# Sheets: táblázatok

A Sheets egy Excelhez hasonló szerkesztő; a számítás egy különálló Rust-motorfolyamatban fut (ha az összeomlik, az soha nem viszi magával az alkalmazást). Valódi .xlsx fájlokat nyit meg és ment; a .csv és a .tsv táblázatként nyílik meg.

## A felület

- **Menüszalag**: nyolc lap, amelyeket lentebb egyenként ismertetünk.
- **Szerkesztőléc**: az aktív cella képletét mutatja és szerkeszti; a gyakori függvények támogatottak.
- **Munkalapfülek** (alul): munkalapok hozzáadása / átnevezése / törlése / áthelyezése.
- **Cellaszerkesztés**: kattints duplán, vagy egyszerűen kezdj el gépelni; az Enter jóváhagyja és lefelé lép, a Tab jobbra lép, az Esc megszakítja (mint az Excelben).
- **Billentyűparancsok**: az Excelhez igazodnak (ctrl+C/V/X, ctrl+Z/Y, ctrl+F, ...).

## A menüszalag lapjai

- **Kezdőlap**: betűtípus, kitöltés, szegélyek, számformátumok (pénznem/százalék/ezres tagolás, tizedesjegyek növelése/csökkentése), igazítás, cellák egyesítése, sorok/oszlopok beszúrása és méretezése, feltételes formázás, formázás táblázatként, cellastílusok, vágólap és formátummásoló, rendezés és szűrés.
- **Beszúrás**: képek, alakzatok, szövegdobozok, hivatkozások, megjegyzések, jelölőnégyzet, élőfej/élőláb, szimbólumok, egyenletek; **kilenc diagramtípus** (oszlop, sáv, vonal, terület, kör, pont, radar, fánk, valamint oszlop+vonal kombináció) plusz a **Javasolt diagramok**, amely a kijelöléshez rangsorolja a típusokat, és előnézetet mutat róluk; **Kimutatásdiagram** abból a cellából, amelyben éppen vagy; **Értékgörbék** (vonal, oszlop, nyereség/veszteség); **Szeletelő** és **Idővonal** kimutatás szűréséhez.
  - A diagramok, kimutatásdiagramok és értékgörbék valódi objektumok a mentett munkafüzetben.
  - A **Szeletelő** és az **Idővonal** csak a munkamenetre vonatkozó vezérlők: az általuk végzett szűrés mentődik, és az Excel ugyanazt a szűrt kimutatást mutatja, de maga a szeletelő gombja nem része a fájlnak.
- **Lapelrendezés**: téma színei és betűtípusai, a rácsvonalak/fejlécek nyomtatásának kapcsolói, oldaltörés-előnézet.
- **Képletek**: AutoSzum és függvények beszúrása, nevek definiálása (kijelölésből is), hivatkozás-előzmények/-követők nyomon követése, a Figyelőablak, munkalap/munkafüzet újraszámolása.
- **Adatok**: rendezés és szűrés (köztük a speciális szűrő és a szűrő törlése), szöveg oszlopokba, munkafüzetek egyesítése, az összes frissítése.
- **Véleményezés**: megjegyzések böngészése (megjelenítés, előző/következő), fordítás, valamint egy **Védelem** csoport — **Lapvédelem**, **Munkafüzet védelme** és **Tartományok szerkesztésének engedélyezése**.
  - A védelem a .xlsx fájlba íródik, és az alkalmazás által beállított védelemhez nincs jelszó, így ugyanez a gomb **Védelem feloldása …** feliratúvá változik, és visszavonja azt. Egy másik programban beállított, _jelszóval_ is védett védelmet innen nem lehet eltávolítani.
  - A két védelmi gomb azonnal érvénybe lép — nincs megszakítható párbeszédpanel, csak egy állapotsori megjegyzés arról, hogy a védelem mentéskor íródik a fájlba.
  - A **Tartományok szerkesztésének engedélyezése** megjelöli azokat a cellákat, amelyek szerkeszthetők maradnak, miközben a munkalap többi része zárolva van.
- **Nézet**: rácsvonalak, szerkesztőléc, fejlécek és az aktív sor és oszlop kiemelésének kapcsolói; nagyítás; **Normál** és **Oldaltörés-előnézet**. A rácsvonalak és a fejlécek a munkalappal együtt mentődnek; az aktív sor és oszlop kiemelése a saját beállításod.
- **Diagramtervezés**: kijelölt diagram esetén jelenik meg — diagramtípus, stílusok és színek, az adattartomány szerkesztése.

Az Adatok lap, gombonként (a képen balról jobbra):

![Az Adatok lap](img/sheets-data.png)

- **Kimutatás**: kimutatás készítése az aktuális tartományból; a mezők áthúzásával összesíthetsz.
- **Frissítés**: az aktuális kimutatás adatainak frissítése.
- **Szövegből/CSV-ből**: .csv/.txt importálása új munkalapként, elválasztó karakter szerint felbontva.
- **Munkafüzetek egyesítése**: munkalapok átvétele más .xlsx fájlokból ebbe a munkafüzetbe.
- **Az összes frissítése**: minden kimutatás és külső adathalmaz frissítése.
- **Rendezés** (legördülő): növekvő / csökkenő / egyéni rendezés (több oszlopra vonatkozó szabályokkal).
- **Szűrő**: ▼ legördülő nyilakat tesz a fejlécsorra; pipáld be a megtartandó értékeket.
- Mellette a kis, egymás alatti gombok: **Törlés** (minden sor visszakerül), **Újraalkalmazás** (az aktuális szűrő újbóli futtatása), **Speciális** (szűrés feltételtartománnyal).
- **Szöveg oszlopokba** (legördülő): egy oszlop felbontása több oszlopra elválasztó karakter vagy rögzített szélesség szerint.
- **Villámkitöltés**: adj meg egy példát, és az oszlop többi része a példa alapján töltődik ki (ctrl+E).
- **Ismétlődések eltávolítása**: az ismétlődő sorok eltávolítása a kijelölt oszlopok alapján.
- **Adatérvényesítés** (legördülő): beviteli szabályok a kijelöléshez (legördülő listák, számtartományok...).
- **Konszolidálás**: több tartomány összesítése egy helyre, kategória szerint.
- **Feltételezéses elemzés** (legördülő): Célérték keresése — egy bemeneti cellát úgy határoz meg, hogy egy képletet tartalmazó cella elérje a célértéket.
- **Csoportosítás / Csoportbontás** (legördülő): sor-/oszlopcsoportok összecsukással és kibontással.
- **Részösszeg**: részösszegsorok beszúrása kategóriánként.

A Képletek lap, gombonként:

![A Képletek lap](img/sheets-formulas.png)

- **Függvény beszúrása** (fx): függvények keresése, argumentumokat segítő varázslóval.
- **AutoSzum** (legördülő): SZUM egy kattintással, plusz átlag/darabszám/max/min.
- **Legutóbb használt / Pénzügyi / Logikai / Szöveg / Dátum és idő / Keresés és hivatkozás / Matematikai és trigonometriai / Továbbiak**: függvények böngészése és beszúrása kategória szerint.
- **Névkezelő**: elnevezett tartományok megtekintése, létrehozása és törlése.
- **Név definiálása** (legördülő): a kijelölés elnevezése; a **Használat képletben** egy meglévő nevet szúr be; a **Létrehozás kijelölésből** a fejlécsor/-oszlop alapján nevez el tartományokat.
- **Előzmények nyomon követése / Követők nyomon követése**: kék nyilak, amelyek megmutatják, honnan származnak egy képlet adatai, és hová jutnak tovább; a **Nyilak eltávolítása** törli őket.
- **Képletek megjelenítése**: a cellák az eredmény helyett magát a képletet mutatják.
- **Hibaellenőrzés**: képlethibák megkeresése és magyarázata.
- **Figyelőablak**: tűzd ki a számodra fontos cellákat, és kövesd az élő értékeiket.
- **Számítási beállítások** (legördülő): automatikus vagy kézi újraszámolás; kézi módban a **Számítás most / Munkalap számítása** indítja el kézzel.

## Számok és formázás

- Számformátumok: általános, szám, pénznem, százalék, dátum/idő, tört, tudományos és továbbiak.
- Igazítás, sortöréssel több sorba, egyesített cellák, szegélyek és kitöltések.
- A sormagasság és az oszlopszélesség húzással állítható; kattints duplán egy határvonalra az automatikus illesztéshez.

## Adatok

**Rendezés és szűrés** (például csökkenő sorrend egy oszlop szerint):

1. Kattints **az adott oszlop bármelyik cellájára** (nem kell a teljes oszlopot kijelölni).
2. Kezdőlap ▸ **Rendezés és szűrés** ▸ **Csökkenő**; a teljes sorok együtt rendeződnek át (a terület egységként rendeződik).
3. Egyéni szabályokhoz (több oszlop, szín szerint): ugyanezen az úton, az **Egyéni rendezés** parancs.
4. Szűrés: jelöld ki a fejlécsort, és kattints a **Rendezés és szűrés ▸ Szűrő** parancsra — minden fejléc kap egy ▼ legördülő nyilat, ahol pipával jelölheted a megtartandó értékeket; a szűrő törlésével minden visszakerül.

- Rendezés és szűrés.
- Rögzített ablaktáblák.
- .csv / .tsv: közvetlenül táblázatként nyílik meg (a tabulátorral tagolt .tsv is táblázatként értelmeződik); mentéskor az eredeti formátum íródik vissza.

## Helyi menük

- **A rácsban**: a szerkesztő (Univer) saját menüje — kivágás/másolás/beillesztés, sorok/oszlopok beszúrása és törlése, elrejtés, cellák egyesítése, ablaktáblák rögzítése és más gyakori parancsok.
- **Az alsó állapotsoron**: kiválaszthatod, mely statisztikákat mutassa az állapotsor (átlag / darabszám / összeg, ...); a választás megmarad.
- **Egy alul lévő munkalapfülön**: munkalapok hozzáadása / átnevezése / törlése / színezése / elrejtése (az Univer lapmenüje).
- A felső lapsáv helyi menüjéről a [Lapok és ablakkezelés](help://tabs-and-windows) cikk szól.

## AI

- Az oldalsáv AI-panelje: jelölj ki egy tartományt, és adj utasítást köznyelven (formázd át, generálj adatokat, írj képleteket).
- Csatolj fájlokat a prompthoz a 📎 gombbal, vagy húzd őket a panelre; a kérdéssel együtt elmennek, a képek pedig miniatűrként jelennek meg a válaszban.
- A válasz hivatkozhat egy cellára — kattints a hivatkozásra, és a rács odaugrik.
- Az AI módosításai visszavonhatók a panelről.

## Mentés és exportálás

- .xlsx fájlként ment (a képletek és formátumok megmaradnak); Mentés másként; az exportálás PDF-be a nyomtatási oldaltördelést követi.
- Az automatikus mentés az általános szabályt követi (az első kézi mentés után kapcsol be).

## Stabilitás

- A számítást végző Rust-segédfolyamat külön folyamatban fut, elkülönítve a felülettől: ha szélsőséges adatok miatt leáll, üzenetet és munkamenet-helyreállítási kísérletet kapsz — nem az alkalmazás összeomlását.
