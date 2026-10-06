# Rychlý start: rozhraní a základy

GenOffice je kancelářský balík, který běží celý na vašem počítači: jedno okno, jeden řádek karet a v něm šest editorů — Docs (textový editor), Sheets (tabulky), Slides (prezentace), PDF, Markdown a HTML. Soubory jsou skutečné soubory .docx / .xlsx / .pptx / .pdf, plně zaměnitelné s Wordem, Excelem a PowerPointem. Není potřeba žádné připojení k síti.

## Přehled rozhraní

![Obrazovka domovské stránky](img/home-screen.png)

Okno má tři části:

- **Pruh karet (nahoře)**: každý otevřený soubor je jedna karta. Karta Domů, úplně vlevo, je vždy přítomná a nelze ji zavřít; ostatní karty jsou vaše dokumenty. Dvojklikem na kartu přejmenujete soubor přímo na místě.
- **Obsahová plocha**: editor (nebo Domů) patřící aktivní kartě.
- **Pruh nabídek**: v systémovém pruhu nabídek na macOS, nahoře v okně na Windows/Linux. Nabídky Soubor/Úpravy/Zobrazení se přizpůsobují aktivnímu editoru.

## Vytvoření dokumentu

Kterýkoli z těchto způsobů:

- Klikněte na kartu rychlého vytvoření v části [Rychlý start](help://getting-started) na **domovské stránce** (AI Docs, AI Sheets, AI Slides, ...).
- Nabídka **Soubor ▸ Nový**: Docs (⌘N/ctrl+N), Sheets, Slides, Markdown, HTML nebo PDF.
- Přetáhněte soubor do okna, nebo na něj v Průzkumníku dvakrát klikněte (pokud je GenOffice výchozí aplikací).

Nový dokument se otevře bez názvu; soubor na disku vznikne až při prvním uložení.

## Otevírání souborů

- Nabídka **Soubor ▸ Otevřít…** (⌘O/ctrl+O) otevře systémový výběr: .docx / .doc / .xlsx / .xls / .csv / .tsv / .pptx / .ppt / .pdf / .md / .html.
- Klikněte na cokoliv v seznamu **Nedávné** na domovské stránce.
- Také `genoffice <soubor>` z terminálu soubory otevírá.

## Model ukládání

- **Ruční uložení**: ⌘S/ctrl+S, případně Soubor ▸ Uložit / Uložit jako… Při prvním uložení se zeptá na umístění a název.
- **Automatické ukládání** se zapne teprve poté, co soubor uložíte ručně alespoň jednou — PDF, které jste si jen přečetli, se nikdy tiše nepřepíše. Automatické ukládání se spustí krátce po změně obsahu.
- Zavřete-li kartu s neuloženými změnami, nejdřív se zeptá, zda uložit, zahodit, nebo zrušit.
- Každý zápis je atomický (dočasný soubor + přejmenování), takže výpadek napájení nenechá napůl zapsaný soubor.

## Běžné klávesové zkratky

| Akce                  | macOS | Windows / Linux |
| --------------------- | ----- | --------------- |
| Nový dokument         | ⌘N    | ctrl+N          |
| Otevřít               | ⌘O    | ctrl+O          |
| Uložit                | ⌘S    | ctrl+S          |
| Zavřít kartu          | ⌘W    | ctrl+W          |
| Otevřít tuto příručku | F1    | F1              |
| Sbalit pás karet      | ⌥⌘R   | Ctrl+F1         |

**Sbalení pásu karet** funguje v každém editoru. Řádek karet zůstane na místě a pás příkazů pod ním se skryje; vybraná karta zároveň slouží jako ovládací prvek sbalení, takže dokud je pás karet sbalený, není vybraná žádná karta a stiskem libovolné karty se pás znovu ukáže. Dvojklik na kartu udělá totéž. Způsob, jakým jste to nechali, se pamatuje pro každý editor zvlášť.

Zkratky uvnitř jednotlivých editorů (formátovací štětec, najít a nahradit, práce s tabulkami, ...) najdete v příslušných kapitolách; Docs navíc nabízí prohledávatelný dialog klávesových zkratek (**⌘/**) (viz jeho kapitola).

## Zkratky s kombinací Option+Command

Option+Command je vrstva, kterou Word vyhrazuje pro strukturované skoky, a GenOffice ji vyplňuje stejným způsobem. Docs zabere většinu z nich, Sheets si bere dva vlastní pro shodu s Excelem a jedna zkratka funguje všude.

**Docs**

| Zkratka (macOS) | Co dělá                   | Windows / Linux    |
| --------------- | ------------------------- | ------------------ |
| ⌥⌘1 / ⌥⌘2 / ⌥⌘3 | Nadpis 1 / 2 / 3          | Ctrl+Alt+1 / 2 / 3 |
| ⌥⌘0             | Normální                  | Ctrl+Alt+0         |
| ⌥⌘M             | Odstavec                  | Ctrl+Alt+M         |
| ⌥⌘A             | Nový komentář             | Ctrl+Alt+A         |
| ⌥⌘F             | Vložit poznámku pod čarou | Ctrl+Alt+F         |
| ⌥⌘E             | Vložit vysvětlivku        | Ctrl+Alt+D         |
| ⌥⌘G             | Přejít na                 | Ctrl+G             |

Dva z nich se na Windows mění, ze stejného důvodu, pro který je Word rozděluje. **macOS si vzal ⌥⌘D** — zobrazuje a skrývá Dock —, takže vysvětlivka je na Macu ⌥⌘E a jinde Ctrl+Alt+D. A **Přejít na** Alt zahazuje: Ctrl+G, zatímco zkratka na Macu ho nese s sebou.

**Sheets**, když má mřížka fokus

| Zkratka (macOS) | Co dělá           | Windows / Linux |
| --------------- | ----------------- | --------------- |
| ⌥⌘0             | Vnější ohraničení | Ctrl+Shift+7    |
| ⌥⌘−             | Bez ohraničení    | Ctrl+Shift+−    |

Windows není přepisem dvojice z Macu. Excel pro Mac dává Sheets **oba** — ⌘⇧7 a ⌥⌘0 jsou dvě klávesy pro stejné vnější ohraničení — takže si příkaz na Windows ponechá pozici Ctrl+Shift, kterou už měl, a vrstva Option jednoduše chybí.

Všimněte si, že **⌥⌘0 znamená Normální v Docs a Vnější ohraničení v Sheets**. Obě tyto funkce se nikdy neobjeví ve stejném editoru, takže se v praxi nic nepotká, ale ⌥⌘0 je obsazená a není k dispozici jako globální zkratka.

**Každý editor**: **⌥⌘R / Ctrl+F1** sbalí pás karet, jak je popsáno výše.

Tím pádem zůstává ⌥⌘D volný, aby ho GenOffice na macOS mohla použít, pokud ho některý budoucí příkaz potřebuje.

## Kam dál

- Kde jsou soubory: [Obrazovka domovské stránky](help://home-screen).
- Správa mnoha otevřených souborů: [Karty a správa oken](help://tabs-and-windows).
- Nechat práci udělat AI: [Panel AI asistenta](help://ai-panel).
- Jazyk, motiv, výchozí aplikace: [Nastavení, jazyk, motiv a integrace MCP](help://settings-integrations).
