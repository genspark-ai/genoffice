# Připojení programového agenta

GenOffice mluví protokolem Model Context Protocol, takže programový agent může číst, zapisovat a vykreslovat vaše dokumenty stejnými enginy, jaké používá aplikace. Agent netuší formát souboru: dostane typovaná schémata op ze stejných definic, proti kterým executor validuje.

## Jeho registrace

Běžný případ je jeden příkaz:

```sh
genoffice mcp install all
```

Najde programové agenty na tomto počítači — Claude Code, Codex, Cursor, Gemini CLI, Copilot CLI, OpenCode, Windsurf — a zapíše záznam stdio serveru do vlastní konfigurace každého z nich, zatímco zbytek souboru nechá být, jak ho našel.

```sh
genoffice mcp list             # where each agent stands
genoffice mcp install cursor    # just one
genoffice mcp uninstall cursor  # take it back out
```

Agent nainstalovaný na neobvyklém místě vyžaduje `--dir <path>`; `--force` přepíše záznam, který tam už je.

## Spuštění vlastníma rukama

Pro klienta na jiném počítači jej místo toho vystavte přes HTTP:

```sh
genoffice mcp --http 8765 --token <secret>
```

`--host <addr>` mění, kde naslouchá. Klientovi předáte stejný token.

Přes HTTP putují i soubory: `PUT /files/<name>` nahraje jeden, každý nástroj bere místo cesty adresu `http(s)` a výstupy se vracejí jako adresy ke stažení — a pokud jsou dost malé, jako vložené zdroje. Op, spec a Markdown se v obou případech předávají přímo v těle.

## Co agent dostane

Každý příkaz je nástroj. Zajímavé jsou tyto:

- **`docs`, `sheet`, `slides`** — čtou a upravují soubor cestou zápisu aplikace, jednu **op** po druhé. Nová prezentace jde přes `deck_start`, `deck_page`, `deck_build`.
- **`render`** — jeden PNG na stránku, vykreslený vykreslovačem aplikace, takže agent se může na snímek podívat, místo aby ho hádal.
- **`pdf`** — textová vrstva PDF, stránka po stránce, bez procesu aplikace.
- **`info`** — metadata a souhrn struktury, obvykle nejlevnější první volání na neznámém souboru.
- **`search`, `image`, `media`** — poskytovatelé nastavení v aplikaci, takže agent nepotřebuje vlastní klíče.
- **`merge`** — vyplní šablonu `{{key}}`.

## Schémata a menší rozpočet

`apply` a `create` oznamují své parametry `ops`, `cells` a `data` s typovaným schématem pro každou op, vygenerovaným z `genoffice guide <domain> --json`. Je to přesné a ne malé. Klient s úzkým kontextovým oknem si může místo toho požádat o prostá pole:

```sh
genoffice mcp --compact-schemas
GENOFFICE_MCP_COMPACT_SCHEMAS=1 genoffice mcp
```

## Dovednost

Agent, který nezná slovní zásobu op, bude hádat. `genoffice skill` nainstaluje dovednost GenOffice do agentů, které najde, a přináší s sebou referenční příručku a návrhové průvodce — přesně ty materiály, které vypíše `genoffice guide`.

## Co to není

Server MCP čte a zapisuje soubory. Není to okno: není tu panel AI a dialog aktualizace v aplikaci zde neplatí. Pokud nějaký krok potřebuje okno, soubor otevřete.
