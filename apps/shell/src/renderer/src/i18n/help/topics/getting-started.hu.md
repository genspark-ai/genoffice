# Gyors kezdés: a felület és az alapok

A GenOffice egy teljes egészében a gépeden futó irodai programcsomag: egyetlen ablak, egy lapsor, benne hat szerkesztő — Docs (szövegszerkesztés), Sheets (táblázatkezelés), Slides (bemutatók), PDF, Markdown és HTML. A fájlok valódi .docx / .xlsx / .pptx / .pdf fájlok, amelyeket a Word, az Excel és a PowerPoint is megnyit, és fordítva. Nincs szükség hálózatra.

## A felület áttekintése

![A kezdőképernyő](img/home-screen.png)

Az ablak három részből áll:

- **Lapsáv (fent)**: minden megnyitott fájl egy lap. A bal szélső lap, a Kezdőképernyő, mindig jelen van, és nem zárható be; a többi lap a dokumentumaid. Kattints duplán egy lapra, és a fájl nevét helyben átnevezheted.
- **Tartalomterület**: az aktív laphoz tartozó szerkesztő (vagy a Kezdőképernyő).
- **Menüsor**: macOS-en a rendszer menüsorában, Windows/Linux alatt az ablak tetején. A Fájl/Szerkesztés/Nézet menük az aktív szerkesztőhöz igazodnak.

## Dokumentum létrehozása

Az alábbiak bármelyike megfelel:

- Kattints egy gyorslétrehozó kártyára a **Kezdőképernyő** [Gyors kezdés](help://getting-started) szakaszában (AI Docs, AI Sheets, AI Slides, ...).
- **Fájl ▸ Új** menü: Docs (⌘N/ctrl+N), Sheets, Slides, Markdown, HTML vagy PDF.
- Húzz egy fájlt az ablakra, vagy kattints rá duplán a fájlkezelődben (ha a GenOffice az alapértelmezett alkalmazás hozzá).

Az új dokumentum névtelenül nyílik meg; a fájl a lemezen csak az első mentéskor jön létre.

## Fájlok megnyitása

- A **Fájl ▸ Megnyitás** (⌘O/ctrl+O) a rendszer fájlválasztóját nyitja meg: .docx / .doc / .xlsx / .xls / .csv / .tsv / .pptx / .ppt / .pdf / .md / .html.
- Kattints a Kezdőképernyő **Legutóbbiak** listájának bármelyik elemére.
- A `genoffice <file>` parancs terminálból kiadva szintén megnyitja a fájlokat.

## A mentési modell

- **Kézi mentés**: ⌘S/ctrl+S, vagy Fájl ▸ Mentés / Mentés másként. Az első mentésnél a program megkérdezi a helyet és a nevet.
- Az **automatikus mentés** csak azután kapcsol be, hogy legalább egyszer kézzel mentetted a fájlt — egy PDF-et, amelyet csak olvasol, sosem ír felül észrevétlenül. Az automatikus mentés a tartalom módosítása után röviddel lefut.
- Ha mentetlen módosításokat tartalmazó lapot zársz be, a program előbb rákérdez: Mentés / Elvetés / Mégse.
- Minden írás atomi módon kerül a lemezre (ideiglenes fájl + átnevezés), így áramkimaradás esetén sem marad félig megírt fájl.

## Gyakori billentyűparancsok

| Művelet                 | macOS | Windows / Linux |
| ----------------------- | ----- | --------------- |
| Új dokumentum           | ⌘N    | ctrl+N          |
| Megnyitás               | ⌘O    | ctrl+O          |
| Mentés                  | ⌘S    | ctrl+S          |
| Lap bezárása            | ⌘W    | ctrl+W          |
| A kézikönyv megnyitása  | F1    | F1              |
| Menüszalag összecsukása | ⌥⌘R   | Ctrl+F1         |

**A menüszalag összecsukása** minden szerkesztőben működik. A lapsor megmarad, az alatta lévő parancssáv pedig elrejtődik; a kijelölt lap egyben az összecsukás vezérlőeleme is, ezért amíg a menüszalag össze van csukva, egyetlen lap sem kijelölt, és bármelyik lapra kattintva a sáv visszatér. Egy lapra duplán kattintva ugyanez történik. Azt, hogy hogyan hagytad, szerkesztőnként megjegyzi a program. A Docs egy köztes sűrűségi szintet is kínál: a **⌥⌘K** / Ctrl+Alt+K a sávot csak ikonokra zsugorítja — ezt is ugyanígy megjegyzi.

Az egyes szerkesztőkön belüli billentyűparancsok (formátummásoló, keresés és csere, táblázatműveletek, ...) a saját fejezeteikben találhatók; a Docs ezenkívül kereshető billentyűparancs-ablakot is kínál (**⌘/**), amely a saját kombinációit teljes egészében tartalmazza.

## Az Option+Command kombinációk

Az Option+Command az a réteg, amelyet a Word a szerkezeti ugrásokra tart fenn, és a GenOffice is ugyanígy használja. Nagy részét a Docs veszi igénybe; a Sheets két saját kombinációt kap az Excellel való egyezés kedvéért; egy kombináció pedig mindenhol működik.

**Docs**

| Kombináció (macOS) | Hatás                   | Windows / Linux    |
| ------------------ | ----------------------- | ------------------ |
| ⌥⌘1 / ⌥⌘2 / ⌥⌘3    | Címsor 1 / 2 / 3        | Ctrl+Alt+1 / 2 / 3 |
| ⌥⌘0                | Normál (szövegtörzs)    | Ctrl+Alt+0         |
| ⌥⌘M                | Bekezdés párbeszédpanel | Ctrl+Alt+M         |
| ⌥⌘A                | Új megjegyzés           | Ctrl+Alt+A         |
| ⌥⌘F                | Lábjegyzet beszúrása    | Ctrl+Alt+F         |
| ⌥⌘E                | Végjegyzet beszúrása    | Ctrl+Alt+D         |
| ⌥⌘G                | Ugrás                   | Ctrl+G             |

Kettő közülük Windowson eltér, ugyanabból az okból, amiért a Wordben is eltérnek. **A ⌥⌘D a macOS-é** — ezzel jelenik meg és tűnik el a Dock —, így a végjegyzet a Macen ⌥⌘E, minden másutt Ctrl+Alt+D. Az **Ugrás** pedig elhagyja az Altot: Ctrl+G, miközben a Mac-kombinációban az is szerepel.

**Sheets**, amíg a fókusz a rácson van

| Kombináció (macOS) | Hatás           | Windows / Linux |
| ------------------ | --------------- | --------------- |
| ⌥⌘0                | Külső szegélyek | Ctrl+Shift+7    |
| ⌥⌘−                | Nincs szegély   | Ctrl+Shift+−    |

A Windows-kombináció nem a Mac-páros átírása. Az Excel for Mac a Sheetsnek **mindkettőt** megadja — a ⌘⇧7 és a ⌥⌘0 két billentyűkombináció ugyanarra a külső szegélyre —, így Windowson a parancs megtartja azt a Ctrl+Shift kombinációt, amelyet amúgy is használt, az Option réteg pedig egyszerűen nincs jelen.

Ne feledd, hogy **a ⌥⌘0 a Docsban a Normál stílust, a Sheetsben pedig a Külső szegélyeket jelenti**. Ez a kettő soha nem fut ugyanabban a szerkesztőben, így használat közben semmi nem ütközik, de a ⌥⌘0 foglalt, és nem használható globális kombinációként.

**Minden szerkesztőben**: a **⌥⌘R / Ctrl+F1** összecsukja a menüszalagot, ahogy fent leírtuk.

Így a ⌥⌘D szabadon marad macOS-en a GenOffice számára, ha egy későbbi parancs igényelné.

## Hová tovább

- Hol vannak a fájlok: [A kezdőképernyő](help://home-screen).
- Sok megnyitott fájl kezelése: [Lapok és ablakkezelés](help://tabs-and-windows).
- Ha az AI végezze el a munkát: [Az AI-asszisztenspanel](help://ai-panel).
- Nyelv, téma, alapértelmezett alkalmazások: [Beállítások, nyelv, téma és MCP-integrációk](help://settings-integrations).
