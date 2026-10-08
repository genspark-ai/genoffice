# Obrazovka domovské stránky: kde žijí vaše soubory

Domů je startovní stránka GenOffice: navigační panel vlevo, seznamy souborů a karty pro rychlé vytvoření vpravo.

![Obrazovka domovské stránky](img/home-screen.png)

## Navigace v postranním panelu

- **Nedávné**: soubory, které jste nedávno otevřeli. Každý řádek nese svůj čas — dnes, včera nebo datum.
- **Oblíbené**: soubory, které jste si označili hvězdičkou. Přejděte na řádek souboru a klikněte na hvězdičku, abyste jej přidali nebo odebrali.
- **Uživatelská příručka**: otevře tuto příručku.
- **Genspark Projects**: po přihlášení k účtu Genspark zobrazí projekty, které jste vytvořili na webu pomocí Genspark AI; kliknutím na projekt pokračujete v úpravách v prohlížeči. K dispozici je vyhledávání, řazení podle času, obnovení a načtení dalších.
- **Složky**: adresáře, které do postranního panelu přidáte tlačítkem **Přidat složku…** nebo do něj přetáhnete ze správce souborů. Každý se stane kořenem, který můžete otevřít, vytvořit v něm podsložky, přejmenovat a odebrat; ten, který se stane nedostupným, se zobrazí jako nedostupný a lze ho ze seznamu sundat. **Nová složka** vytvoří další.

Zde není žádná položka Koš. Odstraněné soubory putují do systémového koše a jejich obnova je věcí operačního systému.

## Seznam souborů

Každý řádek zobrazuje ikonu, název souboru, čas změny a další údaje. Nabídka **⋯** na řádku obsahuje:

- **Otevřít**, a dále **Zobrazit ve složce**, abyste soubor našli ve správci souborů.
- **Kopírovat cestu**.
- **Přesunout do složky…**: otevře výběr složky a soubor skutečně přesune; pokud je v cíli už soubor se stejným názvem, můžete přeskočit, přepsat nebo přejmenovat.
- **Přejmenovat**: přejmenování na místě, přípona se zachová automaticky.
- **Přidat k oblíbeným / Odebrat z oblíbených** — hvězdičky přežijou restart a sledují soubor i po jeho přejmenování.
- **Duplikovat**: vytvoří kopii ve stejné složce.
- **Odstranit**: přesune soubor do systémového koše — nejde o trvalé smazání.
- **Odebrat ze seznamu**, v hlavním zobrazení **Nedávné**, aby položka zmizela, aniž by se souboru něco stalo.

### Více souborů najednou

Zaškrtněte pole u řádku nebo klepněte s ⌘/ctrl, abyste vytvořili výběr; pole v záhlaví vybere vše, co je právě vypsáno, a lišta nad seznamem hlásí (**Vybráno: {n}**), kolik je vybráno, a nabízí **Přesunout do složky…** a **Odstranit soubory** pro celou sadu. Výběr více souborů můžete také přetáhnout na složku v postranním panelu.

## Vyhledávání

Vyhledávací pole nahoře filtruje najednou dvě věci:

- **Názvy souborů**: rychlé filtrování podle názvu.
- **Obsah souborů**: GenOffice na pozadí indexuje vaše soubory (text uvnitř docx/xlsx/pptx/pdf/md/html), takže hledáním v textu najdete i soubory. Rozsah a přepínače se nastavují v nastavení vyhledávání.

## Karty rychlého startu

Karty nad seznamy vytvoří nový dokument jediným krokem. Kliknutím na kartu se vytvoří soubor daného typu a otevře jeho editor — můžete rovnou psát, nebo nechat AI připravit první návrh (každý editor má na pásku **tlačítko AI** a v kontextové nabídce výběru **Zeptat se AI**).

Nový soubor se uloží do složky právě vybrané v postranním panelu; pokud není vybrána žádná, dostane se do výchozí složky.

Co která karta dělá:

- **AI Docs** (.docx): prázdný textový dokument v editoru Docs. Soubor se na disk zapíše až při **prvním uložení**; nové dokumenty se otevírají s rozbaleným panelem AI (vypnout to lze v Nastavení → „Otevírat panel AI v nových dokumentech“).
- **AI Sheets** (.xlsx): prázdný list v editoru Sheets. Dokud neuložíte, na disku žádný soubor není — název je rezervován pro první uložení; po prvním vygenerování pomocí AI lze soubor také automaticky přejmenovat podle jeho obsahu.
- **AI Slides** (.pptx): prázdná prezentace v editoru Slides.
- **AI Markdown** (.md): prázdný dokument Markdown v editoru Markdown.
- **AI HTML** (.html): prázdná webová stránka v editoru HTML.
- **AI PDF** (.pdf): ostatní se liší — **okamžitě** vytvoří skutečný prázdný jednostránkový PDF v cílové složce a otevře ho jako běžný soubor (editor PDF pracuje se skutečnými soubory). Hodí se pro poznámky, začernění nebo přidávání textu; při prvním uložení lze soubor automaticky přejmenovat podle jeho obsahu.
- **Otevřít místní soubor**: systémový výběr souborů pro Word (.docx/.doc), Excel (.xlsx/.xlsm/.xls/.csv/.tsv), PowerPoint (.pptx/.ppt), PDF, Markdown (.md/.markdown) a webové stránky (.html/.htm). Lze vybrat více souborů; každý dostane vlastní kartu.

> Tip: Soubor ▸ Nový v pruhu nabídek vytvoří stejné typy dokumentů (⌘N/Ctrl+N ve výchozím stavu vytvoří textový dokument); přetažení souboru do okna ho otevře.

## Projekty v cloudu (Genspark Projects)

- Při prvním použití je nutné přihlásit se k účtu Genspark (postup s kódem zařízení: GenOffice zobrazí kód a vy dokončíte přihlášení v prohlížeči).
- Seznam projektů se synchronizuje s webem; Otevřít v prohlížeči vás přesune tam, kde lze pokračovat.
- Nepřihlášení nijak neomezuje žádnou místní funkci.
