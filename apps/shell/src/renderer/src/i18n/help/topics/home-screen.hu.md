# A Kezdőképernyő: itt élnek a fájljaid

A Kezdőképernyő a GenOffice nyitóoldala: bal oldalt egy navigációs oldalsáv, jobb oldalt fájllisták és gyorsindító kártyák.

![A Kezdőképernyő](img/home-screen.png)

## Navigálás az oldalsávon

- **Legutóbbiak**: a nemrég megnyitott fájljaid. Minden soron látszik, mikor volt megnyitva — ma, tegnap vagy a dátum.
- **Csillagozott**: az általad csillagozott fájlok. Vidd az egérmutatót egy fájl sora fölé, és kattints a csillagra a hozzáadáshoz vagy eltávolításhoz. A ⋯ menü **Áthelyezés csoportba** parancsa elnevezett csoportokba sorolja a csillagozott fájlokat; a Csillagozott nézet ezek szerint szűr.
- **AI kérdezése**: a tudásbázis-panel — tedd fel a kérdéseidet, és a válaszokban a csillagozott fájljaidra hivatkozik a program (lásd lentebb).
- **Genspark-projektek**: a Genspark-fiókodba való bejelentkezés után megjeleníti a weben, a Genspark AI-jal létrehozott projektjeidet; kattints valamelyikre, és a böngészőben folytathatod a szerkesztést. Van keresés, rendezés idő szerint, frissítés és további elemek betöltése is.
- **Mappák**: mappák, amelyeket a **Mappa hozzáadása…** paranccsal adsz az oldalsávhoz, vagy a fájlkezelőből húzol be. Mindegyik gyökérmappává válik, amelyet megnyithatsz, átnevezhetsz, eltávolíthatsz, és amelyben almappákat hozhatsz létre; ha az egyik elérhetetlenné válik, „nem elérhető” jelzéssel jelenik meg, és kivehető a listáról. Az **Új mappa** további mappát hoz létre.

Itt nincs Kuka menüpont. A törölt fájlok a rendszer kukájába kerülnek, egy fájl visszaállítása pedig az operációs rendszer dolga.

## A fájllista

Minden sor mutat egy ikont, a fájlnevet, a módosítás idejét és egyéb adatokat. A sor **⋯ menüje** a következőket kínálja:

- **Megnyitás**, valamint **Megjelenítés mappában**, amellyel megkeresheted a fájlt a fájlkezelődben.
- **Elérési út másolása**.
- **Áthelyezés mappába…**: mappaválasztót nyit meg, és ténylegesen áthelyezi a fájlt; ha a célmappában már van ilyen nevű fájl, választhatsz: kihagyás / felülírás / átnevezés.
- **Átnevezés**: helyben, a kiterjesztés automatikusan megmarad.
- **Csillagozás / Csillag eltávolítása** — a csillagok újraindítás után is megmaradnak, és átnevezéskor is követik a fájlt.
- **Duplikálás**: másolatot hoz létre ugyanabban a mappában.
- **Törlés**: a fájlt a rendszer kukájába helyezi — nem végleges törlés.
- **Eltávolítás a listából** a legfelső szintű Legutóbbiak nézetben: a bejegyzést a fájl érintése nélkül távolítja el.

### Több fájl egyszerre

Jelöld be egy sor jelölőnégyzetét, vagy kattints ⌘/ctrl lenyomva tartásával, és összeáll a kijelölés; a fejléc jelölőnégyzete az összes éppen listázott elemet kijelöli, a lista feletti sáv pedig megmutatja, hány elem van kijelölve, és az egész kijelölésre kínálja az **Áthelyezés mappába…** és a **Fájlok törlése** parancsot. A többszörös kijelölést az oldalsáv valamelyik mappájára is ráhúzhatod.

## Keresés

A felül található keresőmező egyszerre kétféle találatot ad:

- **Fájlnevek**: gyors szűrés név szerint.
- **Fájltartalom**: a GenOffice a háttérben indexeli a fájljaidat (a docx/xlsx/pptx/pdf/md/html fájlokban lévő szöveget), így a szövegtörzsben keresve is megtalálod a fájlokat. A hatókör és a kapcsolók a keresés beállításaiban találhatók.

## Gyorsindító kártyák

A listák feletti kártyák egyetlen lépésben hoznak létre új dokumentumot. Egy kártyára kattintva létrejön az adott típusú fájl, és megnyílik a szerkesztője — kezdhetsz rögtön írni, vagy megkérheted az AI-t, hogy megírja az első változatot (minden szerkesztő menüszalagján ott az **AI-gomb**, a kijelölés helyi menüjében pedig az **AI kérdezése**).

Az új fájl az oldalsávon éppen kijelölt mappába kerül; ha nincs kijelölt mappa, az alapértelmezettbe.

Az egyes kártyák működése:

- **AI Docs** (.docx): üres szöveges dokumentum a Docs szerkesztőben. A fájl csak az **első mentéskor** íródik a lemezre; az új dokumentumok kinyitott AI-panellel nyílnak meg (ezt a Beállítások → „AI-panel megnyitása új dokumentumokban” kapcsolónál kapcsolhatod ki).
- **AI Sheets** (.xlsx): üres táblázat a Sheets szerkesztőben. Amíg nem mentesz, nem létezik fájl a lemezen — a név az első mentésig van fenntartva; az első AI-generálás után a fájl automatikusan átnevezhető a tartalma alapján.
- **AI Slides** (.pptx): üres bemutató a Slides szerkesztőben.
- **AI Markdown** (.md): üres Markdown-dokumentum a Markdown szerkesztőben.
- **AI HTML** (.html): üres weblap a HTML szerkesztőben.
- **AI PDF** (.pdf): eltér a többiektől — **azonnal** létrehoz egy valódi, üres, egyoldalas PDF-et a célmappában, és sima fájlként nyitja meg (a PDF-szerkesztő valódi fájlokkal dolgozik). Alkalmas széljegyzetekhez, kitakaráshoz vagy szöveg hozzáadásához; a fájl az első mentéskor automatikusan átnevezhető a tartalma alapján.
- **Helyi fájlok megnyitása**: a rendszer fájlválasztóját nyitja meg Word (.docx/.doc), Excel (.xlsx/.xlsm/.xls/.csv/.tsv), PowerPoint (.pptx/.ppt), PDF, Markdown (.md/.markdown) és weblap (.html/.htm) fájlokhoz. Több fájlt is kijelölhetsz; mindegyik saját lapot kap. Az egyszerű szöveges fájlok (.txt, .json) is itt nyílnak meg — ezek a Markdown szerkesztő forrásnézetébe kerülnek.

> Tipp: A menüsor Fájl ▸ Új parancsa ugyanezeket a dokumentumtípusokat hozza létre (a ⌘N/Ctrl+N alapértelmezetten szöveges dokumentumot nyit); ha egy fájlt az ablakba húzol, az megnyílik.

## Kérdezz a tudásbázisodtól

- Először csillagozz fájlokat (bármilyen típusút, amit az alkalmazás be tud olvasni); a panel fejlécében látszik a tudásbázis mérete, az üres tudásbázis pedig arra kér, hogy csillagozz valamit.
- A válaszok megnevezik azt a fájlt, amelyből származnak; kattints egy hivatkozásra, és a fájl a saját szerkesztőjében nyílik meg.
- Bejelentkezés szükséges a Gensparkba. A beszélgetés csak a memóriában létezik — az alkalmazás bezárásával véget ér.
- A keresés a helyi fájlindexen fut, akárcsak a Kezdőképernyő keresése; az eltűnt fájlok szó nélkül kikerülnek a tudásbázisból.
- A csoportok csak a Csillagozott listát rendszerezik — a tudásbázis mindig az összes csillagozott fájlt tartalmazza.

## Felhőprojektek (Genspark-projektek)

- Az első használathoz be kell jelentkezned a Genspark-fiókodba (eszközkódos folyamat: a GenOffice megjelenít egy kódot, a bejelentkezést pedig a böngészőben fejezed be).
- A projektlista szinkronban van a weben lévővel; a „Megnyitás böngészőben” oda visz, hogy ott folytasd.
- Ha nem jelentkezel be, az egyik helyi funkciót sem érinti.
