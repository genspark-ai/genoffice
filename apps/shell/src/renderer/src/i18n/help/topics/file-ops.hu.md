# Fájlműveletek: átnevezés, törlés, exportálás

Ez a fejezet a minden szerkesztőben közös fájlműveleteket ismerteti; az egyes szerkesztők saját exportálási lehetőségei a saját fejezetükben találhatók.

A fájl sorának **⋯ menüje** (vidd az egérmutatót egy fájl sora fölé) egy helyen gyűjti össze ezeket a műveleteket:

![Egy fájl sorának ⋯ menüje](img/file-ops.png)

## Átnevezés

Két belépési pont, ugyanazok az ellenőrzések:

- A Kezdőképernyőn a sor **⋯ ▸ Átnevezés** parancsa.
- **Kattints duplán egy fájl lapjára** a helyben történő átnevezéshez (lásd [Lapok és ablakkezelés](help://tabs-and-windows)).

Szabályok: a kiterjesztés automatikusan megmarad; az érvénytelen karaktereket, a ponttal végződő neveket és a fenntartott neveket (CON/NUL és társaik) a program üzenettel elutasítja; az ugyanabban a mappában lévő névvel való ütközést is megakadályozza. A lemezen lévő valódi fájl átneveződik, és a Legutóbbiak, valamint a csillagozott elemek is követik a változást.

## Törlés

- A Kezdőképernyőn a **⋯ ▸ Törlés** parancs: megkérdezi, mely fájlokat töröljön, majd áthelyezi őket a **rendszer kukájába**, ahonnan az operációs rendszer vissza tudja állítani őket. A GenOffice ehhez nem kínál saját visszavonást — a helyreállítás a kuka dolga, nem egy felugró értesítésé.

## Duplikálás

A **⋯ ▸ Duplikálás** „<név> másolat” nevű fájlt hoz létre ugyanabban a mappában; névütközés esetén automatikusan sorszámot fűz a névhez. A másolat nem nyílik meg, hanem a **Legutóbbiak** közé kerül, így csak egy kattintásra van tőled, nem kerül rögtön eléd.

## Mentés és Mentés másként

- A **⌘S / ctrl+S** menti az aktuális fájlt; névtelen fájlnál előbb a helyet és a nevet kéri.
- A **Mentés másként** új fájlt ír, az eredeti érintetlen marad; a további módosítások már az új fájlt érintik.
- Minden mentés atomi (ideiglenes fájl + átnevezés); a mentés közbeni kilépés sem rongálhatja meg a fájlt.
- Az automatikus mentés csak az első kézi mentés után kapcsol be (lásd [Gyorsindítás](help://getting-started)).

## Exportálás PDF-be

- **Docs**: Fájl ▸ Exportálás PDF-ként (vagy a menüszalag gombja), a kialakított oldaltördeléssel.
- **Slides**: az exportálás dianként rasztergrafikává alakít, nagy bemutatóknál előrehaladásjelzővel.
- **Sheets**: az exportálás a nyomtatási oldaltördelést követi.
- Az exportálás rejtett ablakban történik, az eredmény pedig oda kerül, ahová te választod.

## Exportálás Word-fájlba / képként

- **PDF ▸ Konvertálás Wordbe**: a PDF-ből .docx fájlt készít (a konvertálás helyben fut; összetett elrendezéseknél a tökéletes eredmény nem garantált).
- A **Docs** az oldalakat képként is exportálhatja (oldalanként egy PNG).

## Nyomtatás

Mindegyik szerkesztőben a Fájl ▸ Nyomtatás (⌘P/ctrl+P) megnyitja a rendszer nyomtatási párbeszédpaneljét; a PDF-ek az aktuális oldalsorrenddel és elforgatásokkal nyomtatódnak.

## Hol élnek a névtelen fájlok

Az első mentéskor kiválasztott hely lesz a fájl otthona; addig a dokumentum csak a memóriában létezik. Az automatikus mentés csak ez után az első mentés után lép működésbe.
