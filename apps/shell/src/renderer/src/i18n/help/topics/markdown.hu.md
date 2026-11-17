# A Markdown-szerkesztő

A Markdown-szerkesztő .md / .markdown — illetve .txt / .json fájlokat nyit meg — ez utóbbiakat egyből forrásnézetben —, forrásszerkesztéssel és formázott előnézettel.

- **Megnyitás**: a Kezdőképernyőről vagy a Fájl ▸ Megnyitás paranccsal; a CLI-ből is megy.
- **Szerkesztés**: egyszerű szövegszerkesztés; a GFM-bővítmények (táblázatok, teendőlisták, áthúzott szöveg, automatikus hivatkozások) megjelennek az előnézetben.
- **Előnézet**: élő; a relatív hivatkozású erőforrások, például a képek a dokumentum mellől töltődnek be.
- **Mentés**: bájthű — a BOM, a CRLF és a záró soremelés megléte is megmarad; a módosítás nélküli mentés nem írja újra a fájlt.
- **Keresés és csere**: a ctrl+F a forrásszövegben keres; az Összes cseréje visszaírja a módosítást a fájlba.
- **AI**: az előre beállított gombokkal az asszisztens átírhatja, kibővítheti vagy lefordíthatja a dokumentumot. A kijelölés elrejthető az AI elől (lásd az AI-panel fejezetét).

## Az eszköztár

Egy sornyi gomb a szerkesztő fölött (a súgószövegekhez vidd fölé az egérmutatót):

![A Markdown eszköztár](img/md-toolbar.png)

- **Fájl és előzmények**: Mentés, Mentés másként, Visszavonás, Újra, Keresés; a jobb oldali **Automatikus mentés** kapcsoló időzítővel írja lemezre a módosításokat.
- **AI-gomb**: megnyitja az AI-panelt; mellette az átírás / bővítés / fordítás előbeállításai találhatók.
- **Bekezdésstílus** (legördülő lista): válthatsz a szövegtörzs és a címsorszintek között.
- **Karakterformázás**: **félkövér**, _dőlt_, ~~áthúzott~~, `inline code`, hivatkozás.
- **Listák**: felsorolás, számozás, teendőlista.
- **Beszúrás**: táblázat, kép, vízszintes elválasztó.
- **Tulajdonságok**: a fájl elején lévő YAML front matter blokk beszúrása, illetve ugrás oda.
- **Vázlat**: ugrás a címsorok hierarchiája alapján.
- **Helyesírás-ellenőrzés**: a helyesírás-ellenőrzés be- és kikapcsolása az adott dokumentumhoz.
- **Képtárhely-beállítások**: a beillesztett és behúzott képeket előbb a saját tárhelyedre tölti fel — S3-kompatibilis (R2/OSS/COS/MinIO), SM.MS vagy GitHub —, tartaléknak pedig a helyi assets-mappába készített másolat szolgál.

Három gyors példa:

- **Címsor**: állítsd a kurzort a sorba ▸ bekezdésstílus legördülő lista ▸ „Címsor 1”.
- **Táblázat**: kattints a **Táblázat beszúrása** gombra ▸ húzással add meg a sorok és oszlopok számát ▸ gépelj a cellákba; az előnézet azonnal megjeleníti.
- **Teendőlista**: jelölj ki néhány sort ▸ kattints a **Teendőlista** gombra ▸ minden sor `- [ ]` kezdetű lesz, ami az előnézetben jelölőnégyzetként jelenik meg.

## Exportálás

A Fájl menüben találhatók; mindegyik helyben fut, és mindegyik megkérdezi, hová kerüljön az eredmény:

- Az **Exportálás Wordbe…** és az **Exportálás PDF-ként…** valódi .docx, illetve .pdf fájlt hoz létre.
- **Exportálás képként…** tartalomelemenként egy-egy PNG-t ír az általad kiválasztott mappába.
- Az **Átalakítás és megnyitás a Docsban** .docx fájllá alakít, és az alkalmazás beépített Docs-lapján nyitja meg — nem adja át semmilyen felhős szolgáltatásnak, a konvertált másolat pedig egy gyorsítótármappában van, amelyet nagyjából egy hét után kitisztít a rendszer.

## Forrásnézet

A menüszalagon van egy **Forrás** kapcsoló (a neve az alkalmazás nyelvét követi). Kapcsold be, és a szerkesztő helyét a nyers Markdown veszi át: pontosan az a szöveg, amelyet a mentés kiír — semmi szépítés, a háttérben semmi átalakítás.

- **A szerkesztés bájthű.** A forrásnézetből mentve ugyanazok a bájtok keletkeznek, mint a szerkesztőből mentve — a BOM, a CRLF és a záró soremelés megléte is megmarad.
- **Ugyanaz a dokumentum.** Válts szabadon oda-vissza; a forrás a szerkesztő saját szövege, nem egy külön másolat, amelyet össze kellene vezetni.
- **A formázási eszköztár nem érhető el**, amíg a nézet nyitva van, mert a gombok többsége olyan szerkesztői elemeket szúr be, amelyeknek csak a formázott oldalon van értelmük. A nézet bezárásakor az eszköztár visszatér.
- **A .txt és a .json** itt egyből a forrásnézetben nyílik meg: nincs mit formázva megjeleníteni, így a forrás _maga_ a dokumentum.
