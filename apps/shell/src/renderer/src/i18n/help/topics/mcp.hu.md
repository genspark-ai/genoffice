# Kódoló ügynök csatlakoztatása

A GenOffice támogatja a Model Context Protocolt, így egy kódoló ügynök a dokumentumaidat ugyanazokkal a motorokkal olvashatja, írhatja és renderelheti, mint maga az alkalmazás. Az ügynök nem találgatja a fájlformátumot: a típusos op-sémákat ugyanazokból a definíciókból kapja, amelyek alapján a végrehajtó is ellenőriz.

## Regisztrálás az alkalmazásból

Ezt a **Beállítások ▸ Integrációk** panelen teheted meg. A panelnek két fele van, és használhatod mindkettőt, vagy csak az egyiket.

**A képesség.** Soronként egy kódoló ügynök, amelyet a program a gépen talált — Claude Code, Codex, Cursor, Gemini CLI, Copilot CLI, OpenCode, Windsurf —, mindegyiknél ott a **Telepítés**, a **Frissítés** és az **Eltávolítás** gomb, továbbá a **Telepítés másik mappába…**, a **Képesség letöltése (zip)** és az **Elérési út másolása**. Ha az asszisztensed nincs a listán, add meg a GenOffice-nak azt a mappát, amelyből az asszisztens a `SKILL.md` fájlt olvassa, vagy mentsd el a zip-fájlt, és az asszisztens telepítse. A képesség és az MCP egymás mellett is használható: az asszisztens az egyiket választja, és mindkettő pontosan ugyanazt tudja.

**MCP.** Két mód közül választhatsz: **Az asszisztens indítja (ajánlott)**, ilyenkor a megjelenő konfigurációt felveszed a kliensedbe, és az asszisztens maga indítja el a kiszolgálót; vagy a **Helyi HTTP-kiszolgáló**, amelyet az alkalmazás futtat helyetted. Mindkét esetben az asszisztens a GenOffice-szal fog beszélni, és neked sosem kell parancsot begépelned.

**Próbáld ki.** Nyiss új csevegést az asszisztensedben, és kérj valamit köznyelven, például:

- „Alakítsd át a ~/Downloads/report.md fájlt Word-dokumentummá”
- „Készíts egy 6 diás bemutatót a Q3-as eredményeinkről”
- „Konvertáld a budget.xlsx fájlt PDF-be, és nyisd meg a GenOffice-ban”

## Regisztrálás a parancssorból

Ugyanez terminálból — ez a haladó módszer, és ezt érdemes választani, ha az ügynök olyan helyen van, amelyet a panel nem talál meg:

```sh
genoffice mcp install all
```

Ez megkeresi a gépen lévő kódoló ügynököket, és az stdio-kiszolgáló bejegyzését mindegyik saját konfigurációjába beírja, a fájl többi részét érintetlenül hagyva.

```sh
genoffice mcp list             # where each agent stands
genoffice mcp install cursor    # just one
genoffice mcp uninstall cursor  # take it back out
```

Ha az ügynököt szokatlan helyre telepítetted, add meg a `--dir <path>` kapcsolót; a `--force` felülírja a már meglévő bejegyzést.

A kiszolgáló minden lehetősége egyetlen képernyőn látható — a regisztrálás, az eltávolítás és a listázás parancsa, a HTTP-n való kiszolgálás, valamint a két sémakapcsoló:

![A genoffice mcp --help valódi kimenete: az install, uninstall és list parancsok a --http, --host, --token, --compact-schemas, --dir és --force kapcsolókkal](img/mcp.png)

## Futtatás asszisztens nélkül

Ha a kliensed egy másik gépen fut, szolgáld ki inkább HTTP-n:

```sh
genoffice mcp --http 8765 --token <secret>
```

A `--host <addr>` módosítja, melyik címen figyeljen. Add meg a kliensnek is ugyanazt a tokent.

HTTP-n a fájlok is átjutnak: a `PUT /files/<name>` feltölt egyet, minden eszköz elfogad `http(s)` URL-t útvonal helyett, a kimenetek pedig letöltési URL-ként jönnek vissza — ha elég kicsik, beágyazott erőforrásként is. Az opokat, a specifikációkat és a Markdownt mindkét esetben közvetlenül, az üzenetben kell átadni.

## Mit kap az ügynök

Minden parancs egy eszköz. A legfontosabbak:

- **`docs`, `sheet`, `slides`** — fájl olvasása és szerkesztése az alkalmazás saját írási útvonalán, egyszerre egy **op-pal**. Egy új bemutató menete: `deck_start`, `deck_page`, `deck_build`.
- **`render`** — oldalanként egy PNG, az alkalmazás megjelenítője szerinti elrendezésben, így az ügynök meg tudja nézni a diát ahelyett, hogy találgatna.
- **`pdf`** — egy PDF szövegrétege, oldalanként, az alkalmazás elindítása nélkül.
- **`info`** — metaadatok és szerkezeti összefoglaló; ismeretlen fájlnál ez általában a legolcsóbb első hívás.
- **`search`, `image`, `media`** — az alkalmazásban beállított szolgáltatók, így az ügynöknek nincs szüksége saját kulcsokra.
- **`merge`** — `{{key}}` sablon kitöltése.

## Sémák és kisebb tokenkeret

Az `apply` és a `create` az `ops`, `cells` és `data` paramétereit a típusos, opokénti sémával hirdeti meg, amelyet a `genoffice guide <domain> --json` állít elő. Ez pontos, de nem kicsi. Egy szűk kontextusablakú kliens helyette egyszerű tömböket kérhet:

```sh
genoffice mcp --compact-schemas
GENOFFICE_MCP_COMPACT_SCHEMAS=1 genoffice mcp
```

## Miért a képesség

Az az ügynök, amely nem ismeri az op-szókincset, találgatni fog. A képesség tartalmazza a referenciát és a tervezési útmutatókat — ugyanazt az anyagot, amelyet a `genoffice guide` kiír —, így az asszisztens olyan opokat ír, amelyek specifikációját ténylegesen elolvasta. Telepítsd a fent említett panelről, vagy terminálból a `genoffice skill` paranccsal.

## Mihez fér hozzá az alkalmazásban

A kiszolgáló nem korlátozódik a lemezen lévő fájlokra. Amíg a GenOffice fut, az ügynök az ablakon keresztül is dolgozhat:

- **`open_in_genoffice`** megnyit egy fájlt egy új lapon, és előtérbe hozza.
- **`open_documents`** kilistázza az összes megnyitott dokumentumot — azonosító, típus, útvonal, valamint hogy van-e mentetlen módosítása —, majd beolvassa az egyik élő tartalmát, a helyén elmenti, vagy bezárja; bezárás előtt ment, hacsak nem kéred az elvetést.
- **A tartalomeszközök** ezt az azonosítót (vagy útvonalat) kapják `document` argumentumként, így a módosítás abba a lapba kerül, amelyet már megnyitottál, és az ablak átvált rá.

Két dolog elérhetetlen marad: nincs AI-panel, és az alkalmazáson belüli frissítési ablak sem érvényes.

## A Helyi HTTP-kiszolgáló szakasz

A **Helyi HTTP-kiszolgáló** alatt az alkalmazás maga is futtathatja a kiszolgálót ahelyett, hogy az asszisztensre bízná: van itt egy engedélyező kapcsoló, egy **Port** mező, egy **Fut / Nem fut** jelző, a **Generálás a háttérben** (dokumentumok közvetlen kiírása egy elérési útra a felület megnyitása nélkül), valamint a **Példa kliensbeállításra**, amelyet bemásolhatsz a kliensedbe. A **Speciális** megnyitása két kapcsolódási URL-t — a Streamable HTTP-t és a régebbi SSE-t —, egy **Állapotellenőrzés** URL-t, valamint a **Naplózás** kapcsolót jeleníti meg, amely a kiszolgáló és az eszközök tevékenységét egy helyi fájlba rögzíti; ezt onnan **Megnyithatod**, **Frissítheted** vagy **Törölheted**. Csak a localhoston figyel.
