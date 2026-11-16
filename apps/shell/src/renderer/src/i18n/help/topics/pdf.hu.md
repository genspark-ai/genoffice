# PDF: olvasás, széljegyzetek és kitakarás

A PDF-szerkesztő menüszalagján öt lap van: **Kezdőlap / Széljegyzetek / Szerkesztés / Oldalak / Nézet**. Olvasásra és írásra is alkalmas: a szöveg szerkeszthető, a tartalom kitakarható, a dokumentum aláírható, az űrlapok pedig kitölthetők.

## Olvasás és navigáció

- Bal oldalsáv: **miniatűrök** (kattintással ugorhatsz oda, az éppen látható tartomány ki van emelve) vagy **vázlat** (könyvjelzők, ha vannak).
- Nagyítás: a nagyítás szabályozója jobbra lent; a ctrl+görgő lépésenként változtatja a nagyítást.
- Elforgatás: oldalanként vagy az összes oldalra egyszerre, az Oldalak lapról; az elforgatások mentéskor visszaíródnak.
- Keresés: a ctrl+F teljes szövegű keresést indít, az összes találat ki van emelve.
- A titkosított PDF-eket egy jelszókérő (külön kis ablakban megjelenő) párbeszédpanel nyitja meg; a jelszó csak erre a munkamenetre érvényes.

## Szöveg kijelölése és jelölése (Széljegyzetek)

Próbáld ki bármelyik bekezdésen:

1. **Húzd végig az egeret egy mondaton** — az egérgomb elengedésekor egy széljegyzetsáv jelenik meg fölötte:

![A széljegyzetsáv a szöveg kijelölése után](img/pdf-highlight.png)

2. Válaszd a **kiemelést** (a sárga színminta színpalettát nyit meg), az **aláhúzást** vagy az **áthúzást**; az **AI kérdezése** a kijelölést a kérdéseddel együtt az AI-panelnek küldi.
3. Egy széljegyzet visszavonásához jelöld ki újra ugyanazt a részt egérhúzással, és kattints a sávon az aktív gombra (a Wordhöz hasonló kapcsolóként működik), vagy jelöld ki a széljegyzetet, és nyomd meg a Delete billentyűt.

Részletek:

- Ha az egérrel végighúzod a szöveget, megjelenik egy felugró sáv: **kiemelés / aláhúzás / áthúzás / másolás / AI kérdezése**.
- A színek a palettáról jönnek; **ha ugyanazt a jelölést egy már megjelölt tartományra alkalmazod, az eltávolítja azt** (a Wordhöz hasonló kapcsoló).
- A fájlba már elmentett jelölések kijelölhetők és törölhetők (⋯ menü vagy Delete).
- **Fontos**: amíg egy rajzolóeszköz be van kapcsolva, a szövegréteg nem jelölhető ki — az eszköz minden elhelyezés után magától kikapcsol, így a következő művelethez már újra kijelölő módban vagy.

## Rajzolóeszközök (Széljegyzetek)

Hat eszköz: **toll, téglalap, ellipszis, nyíl, jegyzet**, továbbá a **kitakaró keret** a Széljegyzetek lapon.

- Minden eszköz kapcsolóként működik: kattints rá a bekapcsoláshoz; **miután egy alakzat a helyére került, az eszköz magától kikapcsol** (a folytatáshoz kattints újra az eszközre); a bekapcsolt eszközre kattintva is kikapcsolhatod.
- A toll a beállított vonalvastagságot követi; a téglalapot, ellipszist és nyilat húzással rajzolod meg; a színek a rajzpalettáról jönnek.
- A lerakott alakzatok kijelölhetők, törölhetők, áthúzhatók, és (téglalap/ellipszis) átméretezhetők.
- **Kitakarás, a teljes munkafolyamat** (egy szövegsor eltüntetése):

  1. Széljegyzetek lap ▸ kattints a **Terület kitakarása** gombra (az eszköz bekapcsol).
  2. **Húzz egy keretet a tartalom fölé** — azt egy vonalkázott jelölés fedi le, és az eszköztárban megjelennek a **Jelölések törlése / Kitakarások alkalmazása** gombok:

  ![Az oldal egy kitakarás megjelölése után](img/pdf-redact.png)

  3. Kattints a **Kitakarások alkalmazása** gombra, és erősítsd meg — az eredmény egy munkapéldány, amelyből a lefedett szövegek és képek fizikailag eltávolításra kerülnek (nem csupán el vannak takarva), és ez nem vonható vissza; az eredeti dokumentum érintetlen marad.

  Elrontottad? A Jelölések törlése gomb eltávolítja az aktuális jelöléseket, így újrarajzolhatod őket.

## Cetlik és megjegyzésszálak

- A **jegyzeteszköz** egy gombostűt helyez el, és a szöveg beírásához megnyit egy kártyát a margón (a szerző neve megadható); megerősítéskor szabványos PDF-szövegsíkú széljegyzetként (Text annotation) menti.
- Kattints egy gombostűre a szál megnyitásához: **válasz** (egyszintű, WPS/Acrobat-stílusú szálak), a saját megjegyzésed **szerkesztése**, egy megjegyzés vagy egy egész szál **törlése**.
- A folyamatban lévő szerkesztések addig megmaradnak, amíg a mentés az új szöveget vissza nem írja a fájlban ugyanabba a széljegyzetbe, így a válaszláncok érintetlenek maradnak.

## PDF-tartalom szerkesztése (Szerkesztés)

- **Szöveg szerkesztése**: kattints a szövegre, és blokkonként szerkesztheted (pdfium motor; a betűtípusokat a program a lehető legjobban megközelíti).
- **Szöveg beszúrása**: kereshető szöveget helyezhetsz el az általad választott betűtípussal, mérettel és színnel.
- **Kép / bélyegző beszúrása**.
- Űrlapok: az AcroForm-mezők közvetlenül kitölthetők; az értékek mentéskor íródnak a fájlba.

## Aláírások

- **Kézzel rajzolt aláírás**: rajzold meg; űrlap aláírásmezőjéhez is csatolható.
- **Képes aláírás**: helyezz el egy képet aláírásként.
- A mentett aláírások újrafelhasználhatók.
- **Tanúsítványos aláírás** (Széljegyzetek lap): írj alá a rendszer tanúsítványtárolójában lévő vagy egy .p12/.pfx fájlban található tanúsítvánnyal; igény szerint az aláírást megjelenítheted az oldalon, és hitelesítheted is a dokumentumot a későbbi módosításokkal szemben.
- Az aláírás **új fájlt** hoz létre (növekményes frissítés — a korábbi aláírások érvényesek maradnak), és az eredeti helyén az aláírt példányt nyitja meg. Előbb ments: az aláírás a fájl lemezen mentett állapotára vonatkozik.
- A dokumentum fölötti sáv felsorolja a meglévő aláírásokat; az **Aláírások megtekintése** megmutatja mindegyik érvényességét, lefedettségét és tanúsítványláncát. A visszavonást (CRL/OCSP) nem ellenőrzi.

## Oldalműveletek (Oldalak)

- **Elforgatás / törlés / átrendezés**: az átrendezéshez húzd a miniatűröket; a törlés megerősítést kér.
- **Oldalak importálása**: oldalakat vehetsz át egy másik PDF-ből. Az **Üres oldal beszúrása** egy üres oldalt illeszt be.
- Az **Oldalak cseréje** egy tartományt cserél le máshonnan származó oldalakra; az **Oldalak körbevágása** vág az oldalakból, és az összes oldalra is alkalmazható.
- Az **Oldalméret** minden oldalt egyetlen papírméretre méretez át; a **Fordított sorrend** a dokumentumot a végétől az elejéig megfordítja.
- **Oldalak kinyerése**: a kijelölt oldalak exportálása új PDF-be.
- **Felosztás**: kétféleképpen is lehetséges — tartományonként több fájlra bontás, vagy minden oldal felosztása kisebb oldalakból álló rácsra.
- **Egyesítés**: kétféleképpen is lehetséges — más PDF-ek hozzáfűzése, vagy több oldal összevonása egyetlen lapra. A program a méreteket **még azelőtt** összeadja, hogy bármit beolvasna, és ami összesen **1 GiB fölé megy, azt érthető hibaüzenettel visszautasítja** (így a memóriahasználat korlátok között marad).
- Az oldalszintű módosítások a következő mentéskor íródnak vissza; a Mentés másként az eredeti fájlt érintetlenül hagyja.

## Exportálás és nyomtatás

- **Exportálás Wordbe… / PowerPointba… / Excelbe…** a Fájl menüben, vagy ugyanez a három a menüszalagon a **PDF-átalakító** alatt — mindegyik helyben fut, feltöltés nélkül. A .pptx-ben oldalanként egy dia lesz, a .xlsx-ben pedig oldalanként egy munkalap. Mindegyik megkérdezi, hová mentsd.
- **Nyomtatás**: az aktuális oldalsorrenddel és elforgatásokkal, a rendszer nyomtatási párbeszédpaneljén keresztül; oldaltartományok is megadhatók.

## Mentés

- A normál mentés és az automatikus mentés a széljegyzeteket és a szerkesztéseket visszaírja a fájlba (atomi módon).
- **A kitakarás a saját, Alkalmazás nevű lépésein keresztül történik**: ez másolatot készít, és az eredetit érintetlenül hagyja, így az érzékeny tartalom nem marad a fájlban.
