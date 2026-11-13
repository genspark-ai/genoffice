# Parancssor és ügynökök

Minden telepítés tartalmaz egy `genoffice` parancsot, amely ugyanazokat a motorokat használja, mint az ablak — ugyanazokat az elemzőket, ugyanazt az írót, ugyanazt a renderelőt. Az alkalmazás által mentett fájl és a parancs által írt fájl ugyanolyan, és ami az alkalmazás AI-paneljének ellenőrzésén átmegy, az a parancson is átmegy.

```sh
genoffice --help          # every command
genoffice guide slides    # the op reference, for writing your own
```

Ez a teljes felület egyetlen képernyőn — minden parancs egy-egy soros leírással, majd a globális kapcsolók és a kilépési kódok:

![A genoffice --help valódi kimenete: a verziófelirat, minden parancs egy soros leírással, valamint a globális kapcsolók és a kilépési kódok](img/cli.png)

## A parancs elérhetővé tétele

macOS-en és Windowson az alkalmazáscsomagban található. Ha névvel szeretnéd használni, futtasd egyszer a `genoffice install-cli` parancsot: ez a becsomagolt binárist belinkeli a `PATH`-on szereplő egyik könyvtárba — macOS-en a `/usr/local/bin` könyvtárba, ha az nem Homebrew-előtag (a Homebrew saját könyvtárait soha nem érinti), egyébként a felhasználói bin könyvtárba —, Windowson pedig a felhasználói `PATH`-ba.

## A fontosabb parancsok

| Parancs           | Mit tesz                                                                                                                                                                                                                                                                   |
| ----------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `open`            | Dokumentum megnyitása az alkalmazásban; elindítja az alkalmazást, ha az nem fut.                                                                                                                                                                                           |
| `selection`       | Amit a felhasználó éppen most kijelölt a fájlban, amíg az alkalmazás fut — ez a támpont az „ezt” / „itt” jellegű utasításokhoz. Blokktartományt, munkalaptartományt, diaelemeket vagy egy oldalt ad vissza, attól függően, hogy a fájl melyik szerkesztőben van megnyitva. |
| `convert`         | Átalakítás formátumok között az alkalmazás saját motorjaival.                                                                                                                                                                                                              |
| `create`          | Dokumentum létrehozása strukturált tartalomból.                                                                                                                                                                                                                            |
| `render`          | Oldalanként egy PNG, úgy, ahogy a renderelő elrendezi.                                                                                                                                                                                                                     |
| `pdf`             | PDF szövegrétegének olvasása oldalanként, az alkalmazás elindítása nélkül.                                                                                                                                                                                                 |
| `info`            | Dokumentum metaadatai és szerkezeti összefoglalója.                                                                                                                                                                                                                        |
| `search`          | Webes keresés vagy képkeresés az alkalmazásban beállított szolgáltatón keresztül.                                                                                                                                                                                          |
| `image` / `media` | Kép generálása, illetve kép-, videó- vagy hangfájl leírása, és kérdések feltevése róla.                                                                                                                                                                                    |
| `merge`           | `{{key}}` helyőrzők kitöltése egy `.docx`, `.pptx` vagy `.xlsx` sablonban.                                                                                                                                                                                                 |
| `capabilities`    | Megmutatja, mely felhőfunkciók vannak beállítva ezen a gépen.                                                                                                                                                                                                              |
| `guide`           | Az op-referencia és a tervezési útmutatók, ugyanazokból a definíciókból generálva, amelyek alapján a végrehajtó ellenőriz — így nem térhet el attól, amit az `apply` elfogad. A `--json` az egyes opok sémájával együtt adja vissza.                                       |
| `install-cli`     | A `genoffice` felvétele a `PATH`-ra.                                                                                                                                                                                                                                       |
| `skill`           | Felsorolja a gépen megtalált kódoló ügynököket, és telepíti vagy frissíti bennük a GenOffice képességet.                                                                                                                                                                   |
| `mcp`             | Minden parancs kiszolgálása Model Context Protocol-eszközként. Lásd: **Kódoló ügynök csatlakoztatása**.                                                                                                                                                                    |

## Szerkesztés: docs, sheets, slides

A `genoffice docs`, a `genoffice sheet` és a `genoffice slides` mindegyike először egy alparancsot vár — `read`, `apply`, `check`, slides esetén pedig `audit`, `render`, `replace` —, és ugyanazt az írási útvonalat használja, mint az alkalmazás. Közös a szókincsük: az **op** egyetlen szerkesztés, a **spec** pedig sorrendben alkalmazott op-ok listája.

```sh
genoffice sheet read report.xlsx
genoffice slides apply deck.pptx --ops ops.json --dry-run
genoffice docs apply brief.docx --ops ops.json
```

A `--dry-run` megmondja, mit tenne egy köteg, de semmit sem ír ki — ez a legolcsóbb módja annak, hogy ellenőrizd a specet, mielőtt élesben lefut. Az alkalmazás AI-panelje pontosan ezekkel az opokkal dolgozik, így bármit, amit tőle kérhetsz, szkriptből is elvégezhetsz.

## Model Context Protocol

A `genoffice mcp` minden parancsot MCP-eszközként szolgál ki, a `genoffice mcp install <agent|all>` pedig egy kódoló ügynök saját konfigurációjába regisztrálja. Erről bővebben lásd: **Kódoló ügynök csatlakoztatása**.

## Amit a parancs nem tesz

A parancs olvassa és írja a fájlt. Nem az alkalmazás: nincs ablak, és az alkalmazáson belüli frissítési ablak sem vonatkozik rá. Ami az ablakot igényli — az AI-panel, a megjelenített dia minőségellenőrzése —, ahhoz meg kell várni, amíg megnyitod a fájlt. A `genoffice render` a pixeleket enélkül is megadja, a `genoffice slides` pedig önállóan ellenőrzi a bemutató elrendezését.
