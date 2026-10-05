# Práce se soubory: přejmenování, odstranění, export

Tato kapitola popisuje práci se soubory společnou všem editorům; vlastní možnosti exportu každého editoru najdete v jeho kapitole.

Nabídka **⋯** na řádku (objeví se po najetí na řádek souboru) sdružuje tyto akce:

![Nabídka ⋯ na řádku souboru](img/file-ops.png)

## Přejmenování

Dvě vstupní místa, jedna sada kontrol:

- **⋯ ▸ Přejmenovat** na řádku domovské stránky.
- **Dvojklik na kartu souboru** přejmenuje přímo na místě (viz [Karty a správa oken](help://tabs-and-windows)).

Pravidla: přípona se zachovává automaticky; nepovolené znaky, tečky na konci a rezervované názvy (CON, NUL a podobně) se odmítnou s hlášením; překrytí názvu ve stejné složce se rovněž zablokuje. Přejmenuje se skutečný soubor na disku a seznamy nedávných a oblíbených se aktualizují.

## Odstranění

- **⋯ ▸ Odstranit** na domovské stránce: přesune soubor do **systémového koše**, odkud jej lze obnovit z operačního systému.
- Po odstranění se na několik sekund zobrazí oznámení s možností vzít zpět — vrácení vrátí soubor přesně tam, kde byl.

## Duplikování

**⋯ ▸ Duplikovat** vytvoří ve stejné složce kopii souboru <jméno> a otevře ji v nové kartě; při kolizi názvů se automaticky připojí číslo.

## Uložit a Uložit jako

- **⌘S / ctrl+S** uloží aktuální soubor; soubor bez názvu se nejdřív zeptá na umístění a název.
- **Uložit jako…** zapíše nový soubor a originál ponechá nedotčený; další úpravy se týkají nového souboru.
- Každé uložení je atomické (dočasný soubor + přejmenování); ukončení uprostřed zápisu soubor nepoškodí.
- Automatické ukládání se spustí teprve po prvním ručním uložení (viz [Rychlý start](help://getting-started)).

## Export do PDF

- **Docs**: Soubor ▸ Exportovat jako PDF… (nebo tlačítko na pásku), s aktuálním stránkováním.
- **Slides**: export rastruje stránku po stránce, u velkých prezentací s ukazatelem průběhu.
- **Sheets**: export používá stránkování pro tisk.
- Exporty se vykreslují ve skrytém okně a ukládají tam, kam zvolíte.

## Export do Wordu a obrázků

- **Exportovat jako Word…** u PDF: převede PDF na soubor .docx (místní převod; složitá rozvržení se překládají, jak nejlépe lze).
- **Docs** umí exportovat stránky jako obrázky (PNG na stránku).

## Tisk

Soubor ▸ Tisk… v každém editoru (⌘P/ctrl+P) otevře systémové okno tisku; soubory PDF se tisknou s aktuálním pořadím stránek a otočením.

## Kam putují soubory bez názvu

Umístění, které zvolíte při prvním uložení, je jejich domovem; do té doby existuje dokument pouze v paměti. Automatické ukládání se ujme až po tomto prvním uložení.
