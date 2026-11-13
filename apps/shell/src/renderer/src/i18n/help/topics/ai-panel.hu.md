# Az AI-asszisztens panel

Minden szerkesztőben előhívható az AI-panel: jelölj ki valamit, adj egy utasítást, és kövesd, ahogy az eredmény folyamatosan megérkezik.

## Megnyitás és használat

![Az AI-panel a Docsban](img/ai-panel.png)

- Belépési pontok: az **AI gomb** az egyes szerkesztők menüszalagján, az **AI kérdezése** a helyi menükben, illetve az AI kérdezése a jelölősávon.
- Írd le a feladatot köznyelven (írd át ezt / ez az oszlop legyen százalékos / alakítsd át az oldal elrendezését...), majd nyomd meg az Entert.
- A válaszok **folyamatosan, szavanként** jelennek meg; ha az AI eszközöket használ (a dokumentum elolvasása, szerkesztése), végrehajtja őket, és addig folytatja, amíg el nem készül.
- **Leállítás**: az aktuális kört bármikor megszakíthatod.
- **Sorbaállítás**: amíg egy válasz érkezik, az Enter ehelyett sorba állítja az üzenetedet — a beviteli mező feletti sáv megszámolja őket; ha megnyitod, szerkesztheted, törölheted, átrendezheted, szüneteltetheted vagy kiürítheted a sort. A következő üzenet a kör végén magától elküldődik; az Esc továbbra is leállítja az aktuális kört.

## Mire képes

- **Docs**: átfogalmazás/bővítés/fordítás/összefoglalás, táblázatok és képek beszúrása, formázás módosítása; minden kör előtt pillanatkép készül.
- **Sheets**: képletek, adatkitöltés, tömeges átalakítások, formázás.
- **Slides**: teljes bemutató generálása, elrendezés igazítása, szövegek átfogalmazása.
- **PDF**: kérdések és összefoglalók a kijelölt szöveg vagy oldalak alapján.
- **Markdown / HTML**: átfogalmazás, bővítés, fordítás.
- A **Kezdőképernyő ▸ AI kérdezése** a csillagozott fájljaid alapján válaszol, és hivatkozik is rájuk — a tudásbázis-panelnek saját szakasza van a Kezdőképernyőn, a Docs AI-panelje pedig ugyanezt a képességet is tartalmazza.

## Visszaállítás és biztonság

- A Docs panel **verziólistát** vezet: körönként egy pillanatkép, bármelyikre visszaléphetsz, és maga a visszaállítás is visszavonható a Ctrl+Z-vel. A pillanatképek a dokumentum újranyitása után is megmaradnak.
- Az AI módosításai ugyanazon a szerkesztési folyamaton mennek át, mint a kézi szerkesztések (visszavonhatók, mentéshez kötöttek) — semmi nem kerüli meg a mentés jóváhagyását.

## Kijelölés elrejtése az AI elől

- A Slidesban, a HTML-ben és a Markdownban a kijelölést elrejtheted a modell elől: kattints rá jobb gombbal ▸ **A kijelölés elrejtése az AI elől**, nevezd el a helyőrzőt, és a modell a helyén a `{{name}}` jelölést fogja olvasni. (A Markdownban ez a menüpont a menüszalagon lévő áthúzott szemű kapcsolóval aktiválható; a HTML-nél a forráspanelen található.)
- A fájl megőrzi az eredeti szöveget — a mentés, az exportálás és a nyomtatás változatlan marad, és a jelölés az újranyitás után is megmarad. Ugyanez a menüpont fel is fedi a szöveget (a Markdownban erre külön **A kijelölés újbóli megjelenítése az AI számára** menüpont szolgál).
- A HTML-ben az elemek tartalma, az attribútumértékek és a szkriptek szövegliterálai rejthetők el; a megjelenített oldal soha nem változik.
- A Slidesban teljes képeket, videókat és hangfájlokat is elrejthetsz, a rejtett tartalmú dia pedig sosem kerül elküldésre képernyőképként: az automatikus elrendezés-ellenőrzés kihagyja azt a diát, és ezt jelzi is.
- Ez csak az alkalmazáson belüli AI-ra vonatkozik — a CLI, az MCP-eszközök és a felület nélküli (headless) exportálás közvetlenül a fájlt olvassa, ezért továbbra is látja a szöveget.

## Adatvédelem

- Az utasítások és a vonatkozó dokumentumtartalom az **általad beállított modellszolgáltatóhoz** kerül (Genspark által üzemeltetett vagy egyéni végpont, lásd a következő fejezetet); ha nincs beállítva, nem megy el semmi.
- A helyi fájlokat máshová nem tölti fel a program; a BYOK-kulcsok az alkalmazás beállítási fájljában tárolódnak ezen a gépen, és csak a kérések fejlécében kerülnek elküldésre.
