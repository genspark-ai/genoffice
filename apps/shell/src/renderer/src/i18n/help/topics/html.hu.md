# A HTML-szerkesztő

A HTML-szerkesztő két módban nyitja meg a .html / .htm fájlokat: **előnézet** (a megjelenített oldal) és **forrás**.

- **Előnézet**: valódi megjelenítés; a relatív stíluslapok és képek a fájl mellől töltődnek be.
- **Előnézet-vizsgáló**: kattintással jelölj ki egy elemet, dupla kattintással szerkesztheted a szövegét a helyén, az eszköztárról törölheted, és a kijelölt elemről kérdezhetsz az AI-tól.
- **Forrás mód**: a HTML szerkesztése; a ctrl+F a keresés; az Összes cseréje elmenti az átírt kódot.
- **Mentés**: bájthű (a BOM/CRLF/záró újsor megmarad); a módosítás nélküli mentés nem írja újra a fájlt.
- **Nagyítás**: a ctrl+görgő / csippentés nagyítja és kicsinyíti az előnézetet; az előnézetben a ctrl+Z visszavonja az utolsó szerkesztést.
- **Elrejtés az AI elől**: forrás módban kattints jobb gombbal egy kijelölésre — elemek tartalmára, attribútumértékekre, szkriptekben lévő sztringekre —, és a modell helyőrzőt lát helyette; a megjelenített oldal sosem változik (részletek az AI-panel fejezetében).

## Az eszköztár

Kattints az előnézet bármelyik elemére, és fölötte megjelenik egy lebegő eszköztár:

![A lebegő eszköztár egy kijelölt elem fölött](img/html-toolbar.png)

- **Fájl és előzmények**: Mentés, Mentés másként, Visszavonás, Újra, Keresés; az **Automatikus mentés** kapcsoló időzítővel menti a módosításokat.
- **Előnézet / Forrás** kapcsoló; a **Bemutatás** teljes képernyőn jeleníti meg az oldalt.
- **Formázás**: félkövér, dőlt, betűméret növelése/csökkentése; a kijelölt elem **stíluspanelje** (színek és egyebek).
- **Beszúrás**: címsor, bekezdés, táblázat, kép (hivatkozással) és más elemek.
- **Képműveletek** (kép kijelölésekor): kivágás, **háttér eltávolítása**, csere, képarány zárolása.
- **Elemműveletek** (elem kijelölésekor az előnézet-vizsgálóban): törlés, duplikálás, mozgatás fel/le.
- **AI gomb**: megnyitja az AI-panelt; a kijelölt elemről közvetlenül kérdezhetsz.

## Exportálás

A Fájl menüben találhatók; mindegyik helyben fut, és mindegyik megkérdezi, hová kerüljön az eredmény:

- Az **Exportálás Wordbe…** és az **Exportálás PDF-ként…** valódi .docx, illetve .pdf fájlt hoz létre.
- Az **Exportálás egyetlen HTML-fájlként…** egyetlen .html fájlt hoz létre, amelybe a képek be vannak ágyazva. Nem írja felül az aktuálisan megnyitott fájlt, és jelzi, hány képet nem tudott beágyazni.

## Alapváz beszúrása

Üres oldalhoz a **Beszúrás ▸ Alapváz beszúrása** egy minimális, szabványmódú dokumentumot szúr be:

```html
<!DOCTYPE html>
<html lang="en">
  <head>
    <meta charset="UTF-8" />
    <title></title>
  </head>
  <body></body>
</html>
```

Minden résznek oka van, ezért ez egy parancs, nem pedig kézzel begépelendő szöveg:

- a **doctype**: nélküle az előnézet quirks módban fut, ahol a dobozméretezés és a táblázatelrendezés más szabályok szerint működik, mint amire számítasz;
- a **`lang`**: nélküle a képernyőolvasó nem tudja, milyen nyelven olvassa fel az oldalt, a böngésző pedig rossz nyelvhez választ betűtípust és helyesírás-ellenőrzőt;
- a **charset**: nélküle egy nem latin írású oldal szövege olvashatatlan karaktersorként (mojibake) jelenhet meg.

A viewport meta szándékosan hiányzik: az oldal egy asztali panelben jelenik meg, ahol nincs mobil nézetablak, amelyre hatással lehetne.

A `lang` az alkalmazás felületi nyelvét követi, így a beszúrt alapváz olyan nyelvű lesz, amelyre az eszközeid már be vannak állítva. Utána szabadon szerkesztheted.

A menüpont csak szerkesztési módban jelenik meg, és csak amíg a dokumentum üres — ha már van tartalom, nincs _hová_ alapvázat beszúrni.
