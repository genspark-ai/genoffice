# Podłączenie z agentem programistycznym

GenOffice mówi językiem Model Context Protocol, więc agent programistyczny może czytać, zapisywać i renderować twoje dokumenty przez te same silniki, których używa aplikacja. Agent nie zgaduje formatu pliku: otrzymuje typowane schematy op z tych samych definicji, względem których waliduje executor.

## Rejestrowanie

Zwykły przypadek to jedno polecenie:

```sh
genoffice mcp install all
```

Znajduje agentów programistycznych na tym komputerze — Claude Code, Codex, Cursor, Gemini CLI, Copilot CLI, OpenCode, Windsurf — i zapisuje wpis serwera stdio we własnej konfiguracji każdego z nich, nie ruszając reszty pliku.

```sh
genoffice mcp list             # where each agent stands
genoffice mcp install cursor    # just one
genoffice mcp uninstall cursor  # take it back out
```

Agent zainstalowany w nietypowym miejscu wymaga `--dir <path>`; `--force` przepisuje wpis, który już tam jest.

## Uruchamianie samodzielne

Dla klienta na innym komputerze udostępnij go przez HTTP:

```sh
genoffice mcp --http 8765 --token <secret>
```

`--host <addr>` zmienia adres nasłuchu. Przekaż ten sam token klientowi.

Przez HTTP przesyłane są też pliki: `PUT /files/<name>` wgrywa jeden, każde narzędzie przyjmuje adres `http(s)` zamiast ścieżki, a wyniki wracają jako adresy do pobrania — a gdy są dość małe, jako zasoby osadzone. Op, specy i Markdown są przekazywane w treści w obu przypadkach.

## Co agent otrzymuje

Każde polecenie jest narzędziem. Najciekawsze:

- **`docs`, `sheet`, `slides`** — czytają i edytują plik przez ścieżkę zapisu aplikacji, jedna **op** na raz. Nowa prezentacja przechodzi przez `deck_start`, `deck_page`, `deck_build`.
- **`render`** — jeden PNG na stronę, ułożony przez renderujący aplikacji, więc agent może obejrzeć slajdę, zamiast ją zgadywać.
- **`pdf`** — warstwa tekstowa PDF-a, strona po stronie, bez procesu aplikacji.
- **`info`** — metadane i podsumowanie struktury, zwykle najtańsze pierwsze wywołanie przy nieznanym pliku.
- **`search`, `image`, `media`** — dostawcy skonfigurowani w aplikacji, więc agent nie potrzebuje własnych kluczy.
- **`merge`** — wypełnia szablon `{{key}}`.

## Schematy i mniejszy budżet

`apply` i `create` deklarują parametry `ops`, `cells` i `data` wraz z typowanym schematem dla każdej op, wygenerowanym z `genoffice guide <domain> --json`. To jest precyzyjne, ale nie małe. Klient o ciasnym oknie kontekstowym może zamiast tego poprosić o zwykłe tablice:

```sh
genoffice mcp --compact-schemas
GENOFFICE_MCP_COMPACT_SCHEMAS=1 genoffice mcp
```

## Skill

Agent, który nie zna słownictwa op, będzie zgadywał. `genoffice skill` instaluje skill GenOffice w znalezionych agentach, zabierając ze sobą referencję i przewodniki projektowe — dokładnie te materiały, które wypisuje `genoffice guide`.

## Czym nie jest

Serwer MCP czyta i zapisuje pliki. To nie jest okno: nie ma panelu AI, a okno aktualizacji w aplikacji nie ma tu zastosowania. Jeśli jakiś krok wymaga okna, otwórz plik.
