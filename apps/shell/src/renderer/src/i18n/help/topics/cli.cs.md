# Příkazová řádka a agenti

Každá instalace obsahuje příkaz `genoffice`, který používá stejné enginy jako okno — stejné parsery, stejný zapisovač, stejný vykreslovač. Soubor, který aplikace uloží, a soubor, který zapíše příkaz, je stejný soubor, a kontrola, kterou projde panel AI v aplikaci, projde i příkaz.

```sh
genoffice --help          # every command
genoffice guide slides    # the op reference, for writing your own
```

## Získání příkazu

Systémy macOS a Windows jej dodávají uvnitř balíčku aplikace. Chcete-li jej volat jménem, spusťte jednou `genoffice install-cli`: vytvoří symbolický odkaz na přiložený spustitelný soubor v `/usr/local/bin`, nebo v uživatelské `PATH` v systému Windows.

## Příkazy, které stojí za to znát

| Příkaz | Co dělá |
| --- | --- |
| `open` | Otevře dokument v aplikaci; aplikaci spustí, pokud neběží. |
| `convert` | Převádí mezi formáty pomocí vlastních enginů aplikace. |
| `create` | Vytvoří dokument ze strukturovaného obsahu. |
| `render` | Jeden PNG na stránku, tak jak jej vykreslí vykreslovač. |
| `pdf` | Čte textovou vrstvu PDF stránku po stránce, bez procesu aplikace. |
| `info` | Metadata a souhrn struktury dokumentu. |
| `search` | Vyhledávání na webu nebo v obrázcích pomocí poskytovatele nastaveného v aplikaci. |
| `image` / `media` | Vygeneruje obrázek, nebo popíše soubor obrázku, videa či zvuku a umožní na něj dotazy. |
| `merge` | Vyplní zástupné značky `{{key}}` v šabloně `.docx`, `.pptx` nebo `.xlsx`. |
| `capabilities` | Označí, které cloudové funkce jsou na tomto počítači nastavené. |
| `guide` | Referenční příručka op a návrhové průvodce, vygenerované ze stejných definic, proti kterým executor validuje — takže se nemůže rozejít s tím, co přijímá `apply`. `--json` ji vrátí se schématem každé op. |
| `install-cli` | Umístí `genoffice` do `PATH`. |
| `skill` | Vypíše programové agenty nalezené na tomto počítači a nainstaluje do nich nebo v nich aktualizuje dovednost GenOffice. |
| `mcp` | Vystaví každý příkaz jako nástroj Model Context Protocol. Viz **Připojení programového agenta**. |

## Úpravy: Docs, Sheets, Slides

`genoffice docs`, `genoffice sheet` a `genoffice slides` čtou a upravují stejnou cestou zápisu, jakou používá aplikace, a sdílejí jednu slovní zásobu: **op** je jediná úprava a **spec** je seznam op aplikovaných v pořadí.

```sh
genoffice sheet read report.xlsx
genoffice slides apply deck.pptx --ops ops.json --dry-run
genoffice docs apply brief.docx --ops ops.json
```

`--dry-run` hlásí, co by dávka udělala, a nic nezapíše — to je nejlevnější způsob, jak si spec před použitím ověřit. Panel AI v aplikaci běží přesně na těchto op, takže všechno, co mu můžete zadat, můžete i zapsat skriptem.

## Model Context Protocol

`genoffice mcp` vystaví každý příkaz jako nástroj MCP a `genoffice mcp install <agent|all>` jej zaregistruje ve vlastní konfiguraci programového agenta. Této části se týká kapitola **Připojení programového agenta**.

## Co příkaz neumí

Čte a zapisuje soubor. Není to aplikace: není tu okno a dialog aktualizace v aplikaci zde neplatí. Všechno, co potřebuje okno — panel AI, kontrolu kvality vykreslené snímku — musí počkat, než soubor otevřete. `genoffice render` vám dá pixely bez něj a `genoffice slides` prověří rozvržení prezentace sám.
