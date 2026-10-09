# PDF: čtení, anotování a začerňování

Editor PDF má pět karet na pásku: **Domů / Anotace / Úpravy / Stránky / Zobrazení**. Umí číst i zapisovat: text lze upravovat, obsah začerňovat a podepisovat a formuláře vyplňovat.

## Čtení a navigace

- Levý postranní panel: **miniatury** (kliknutím přeskočíte, viditelný rozsah je zvýrazněn) nebo **osnova** (záložky, pokud jsou k dispozici).
- Přiblížení: ovladač poměru vpravo dole; ctrl+kolečko mění přiblížení po stupních.
- Otáčení: jednotlivé stránky nebo všechny z nabídky Stránky; otočení se při uložení zapíší zpět do souboru.
- Vyhledávání: ctrl+F v celém textu se zvýrazněním všech nálezů.
- Šifrované soubory PDF: otevře je výzva k zadání hesla (v samostatném malém okně); heslo se použije pouze pro tuto relaci.

## Výběr textu a značení (Anotace)

Vyzkoušejte na libovolném odstavci:

1. **Přetáhněte myší přes větu** — po puštění se nad textem objeví lišta s anotací:

![Lišta s anotací po výběru textu](img/pdf-highlight.png)

2. Zvolte **zvýraznění** (žlutý vzorek otevře paletu barev), **podtržení** nebo **přeškrtnutí**; **Zeptat se AI** odešle výběr spolu s vaší otázkou do panelu AI.
3. Chcete-li anotaci zrušit, znovu přetáhnutím vyberte stejný úsek a klikněte na již aktivní tlačítko na liště (přepínač ve stylu Wordu), nebo výběr vyberte a stiskněte Odstranit.

Co očekávat:

- Přetažení přes text zobrazí místní lištu: **zvýraznění / podtržení / přeškrtnutí / kopírovat / Zeptat se AI**.
- Barvy pocházejí z palety; **použití stejného značení na již označený úsek jej odstraní** (přepínač ve stylu Wordu).
- Značení uložená již v souboru lze vybrat a odstranit (⋯ menu nebo Odstranit).
- **Pozor**: dokud je aktivní nástroj kreslení, textovou vrstvu nelze vybírat — nástroj se po každém umístění sám vypne, takže před další akcí jste opět v režimu výběru.

## Nástroje kreslení (Anotace)

Šest nástrojů: **tuš, obdélník, elipsa, šipka, poznámka**, navíc **začerňovací rameček** na kartě Anotace.

- Každý nástroj je přepínač: kliknutím jej zapnete; **sám se vypne, jakmile je umístěna figura** (chcete-li pokračovat, klikněte na nástroj znovu); kliknutí na již zapnutý nástroj jej rovněž vypne.
- Tuš sleduje šířku tahu; obdélník, elipsa a šipka se táhnou; barvy pocházejí z kreslicí palety.
- Umístěné figury lze vybrat, odstranit, přetáhnout a (obdélník a elipsu) změnit velikost.
- **Začerňování, celý postup** (například skrytí řádku textu):

  1. Karta Anotace ▸ klikněte na **Začernit oblast** (nástroj se zapne).
  2. **Přetáhněte rameček přes obsah** — bude překryt kresbou šrafování a na liště nástrojů se objeví tlačítka **vymazat značky / použít začernění**:

  ![Stránka po označení začernění](img/pdf-redact.png)

  3. Klikněte na **Použít začernění** a potvrďte — výsledkem je pracovní kopie, ve které je zakrytý text a obrázky fyzicky odstraněny (nikoli jen překryty), a operaci nelze vzít zpět; původní dokument zůstane nedotčen.

  Udělali jste chybu? Tlačítko pro vymazání značek smaže aktuální značky, takže můžete kreslit znovu.

## Připnuté poznámky a vlákna komentářů

- **Nástroj poznámky** umístí špendlík a otevře kartu v okraji pro text (jméno autora lze nastavit); po potvrzení se uloží jako standardní textová anotace PDF.
- Kliknutím na špendlík otevřete vlákno: **odpovědět** (plochá vlákna ve stylu WPS/Acrobatu), **upravit** vlastní komentář, **odstranit** jeden komentář nebo celé vlákno.
- Rozpracované úpravy přežijí, dokud uložení nezapíše nový text zpět do téže anotace v souboru, takže řetězce odpovědí zůstanou zachovány.

## Úprava obsahu PDF (Úpravy)

- **Upravit text**: kliknutím na text jej upravíte po blocích (engine pdfium; písma se přiřazují s nejlepším možným přiblížením).
- **Vložit text**: vloží dohledávatelný text s libovolným písmem, velikostí a barvou.
- **Vložit obrázek / razítko**.
- Formuláře: pole AcroForm se vyplňují přímo; hodnoty se zapisují při uložení.

## Podpisy

- **Ruční podpis**: nakreslete jej; lze jej připojit k poli podpisu ve formuláři.
- **Podpis jako obrázek**: umístěte podpis jako obrázek.
- Uložené podpisy lze opakovaně používat.

## Operace se stránkami (Stránky)

- **Otočit / odstranit / přesunout**: tažením miniatur měníte pořadí; odstranění vyžaduje potvrzení.
- **Importovat stránky**: přeneste stránky z jiného souboru PDF. **Vložit prázdnou stránku** přidá prázdnou stránku.
- **Nahradit stránky** nahradí celý rozsah stránkami odjinud; **Oříznout stránky** zkrátí okraje, s možností použít to na všechny stránky.
- **Velikost stránky** přepočítá všechny stránky na jediný rozměr papíru; **Obrátit pořadí** otočí dokument od konce k začátku.
- **Vyjmout stránky**: exportuje vybrané stránky do nového souboru PDF.
- **Rozdělit PDF**: má dvě podoby — rozdělení podle rozsahů na několik souborů nebo rozřezání každé stránky na mřížku menších stránek.
- **Sloučit PDF**: má dvě podoby — připojení dalších souborů PDF nebo spojení několika stránek na jeden list. Velikosti se sčítají **ještě před načtením čehokoli** a soubor o celkové velikosti nad **1 GiB je odmítnut** se srozumitelnou chybou (kvůli omezení paměti).
- Změny na úrovni stránek se zapíší zpět při příštím uložení; Uložit jako… ponechá originál nedotčený.

## Export a tisk

- **Exportovat jako Word… / PowerPoint… / Excel…** v nabídce Soubor, nebo ty stejné tři z **Převodník PDF** na pásku karet — vše lokálně, bez nahrávání. Ze souboru .pptx vyjde jeden snímek na stránku a ze souboru .xlsx jeden list na stránku. U každého se zeptá, kam uložit.
- **Tisk**: aktuální pořadí a otočení přes systémový dialog; podporuje rozsahy stránek.

## Ukládání

- Běžné ukládání, ruční i automatické, zapíše anotace a úpravy zpět do souboru (atomicky).
- **Začerňování prochází vlastním postupem použití** a vytvoří kopii, takže původní dokument zůstane nedotčen a citlivý obsah v něm nezůstane.
