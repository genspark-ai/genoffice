# Podłączenie z agentem programistycznym

GenOffice mówi językiem Model Context Protocol, więc agent programistyczny może czytać, zapisywać i renderować twoje dokumenty przez te same silniki, których używa aplikacja. Agent nie zgaduje formatu pliku: otrzymuje typowane schematy op z tych samych definicji, względem których waliduje executor.

## Rejestrowanie w aplikacji

Robisz to w **Ustawienia ▸ Integracje**. Panel ma dwie połowy i możesz skorzystać z jednej albo z obu.

**Skill.** Po jednym wierszu na każdego znalezionego na tym komputerze agenta programistycznego — Claude Code, Codex, Cursor, Gemini CLI, Copilot CLI, OpenCode, Windsurf —, każdy z przyciskami **Zainstaluj**, **Aktualizuj** i **Odinstaluj**, a także **Zainstaluj w innym folderze…**, **Pobierz skill (zip)** i **Kopiuj ścieżkę**. Jeśli twojego asystenta nie ma na liście, wskaż GenOffice folder, z którego czyta `SKILL.md`, albo zapisz zip i pozwól asystentowi go zainstalować. Skill i MCP mogą działać obok siebie: asystent wybiera jedno z nich, a robią dokładnie te same rzeczy.

**MCP.** Dostępne są dwie drogi: **Uruchamiany przez asystenta (zalecane)** — tam dodajesz pokazaną konfigurację do swojego klienta, a asystent sam uruchamia serwer — oraz **Lokalny serwer HTTP**, który aplikacja prowadzi za ciebie. Tak czy inaczej asystent ostatecznie rozmawia z GenOffice, a ty nie wpisujesz żadnej komendy.

## Rejestrowanie z wiersza poleceń

To samo z terminala — to jest ścieżka zaawansowana, po którą sięgasz, gdy agent jest tam, gdzie panel nie potrafi go znaleźć:

```sh
genoffice mcp install all
```

Znajduje agentów programistycznych na tym komputerze i zapisuje wpis serwera stdio we własnej konfiguracji każdego z nich, nie ruszając reszty pliku.

```sh
genoffice mcp list             # where each agent stands
genoffice mcp install cursor    # just one
genoffice mcp uninstall cursor  # take it back out
```

Agent zainstalowany w nietypowym miejscu wymaga `--dir <path>`; `--force` przepisuje wpis, który już tam jest.

Wszystko, co przyjmuje serwer, mieści się na jednym ekranie — formularze rejestracji, usuwania i listowania, udostępnianie przez HTTP oraz dwie flagi schematów:

![Prawdziwe wyjście genoffice mcp --help: formularze install, uninstall i list wraz z opcjami --http, --host, --token, --compact-schemas, --dir i --force](img/mcp.png)

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

## Po co jest skill

Agent, który nie zna słownictwa op, będzie zgadywał. Skill zabiera ze sobą referencję i przewodniki projektowe — dokładnie te materiały, które wypisuje `genoffice guide` —, więc asystent pisze op, których specyfikację naprawdę przeczytał. Zainstaluj go z panelu powyżej albo poleceniem `genoffice skill` z terminala.

## Do czego sięga w aplikacji

Serwer nie ogranicza się do plików na dysku. Dopóki GenOffice działa, agent może pracować także przez okno:

- **`open_in_genoffice`** otwiera plik w karcie i ustawia na niej fokus.
- **`open_documents`** wymienia każdy otwarty dokument — id, typ, ścieżkę i to, czy ma niezapisane zmiany —, a następnie czyta bieżącą zawartość jednego z nich albo go zamyka, najpierw zapisując, chyba że każesz mu odrzucić zmiany.
- **Narzędzia od treści** przyjmują to id (albo ścieżkę) jako argument `document`, więc zmiana trafia do karty, którą już masz otwartą, a okno przełącza się na jej pokazanie.

Dwie rzeczy pozostają poza zasięgiem: nie ma panelu AI, a okno aktualizacji w aplikacji nie ma tu zastosowania.

## Panel lokalnego serwera HTTP

W sekcji **Lokalny serwer HTTP** aplikacja prowadzi serwer sama, zamiast zostawiać to asystentowi: przełącznik włączania, pole **Port**, wskaźnik **Działa / Zatrzymany** i **Generowanie w tle** (zapisywanie dokumentów prosto pod ścieżką bez otwierania interfejsu) oraz **Przykładowa konfiguracja klienta** do skopiowania. Otwarcie sekcji **Zaawansowane** dodaje oba adresy połączenia — Streamable HTTP i starszy adres SSE —, adres **Sprawdzenia stanu** oraz przełącznik **Rejestrowanie**, który zapisuje pracę serwera i narzędzi do lokalnego pliku, który można stamtąd **Otworzyć**, **Odświeżyć** lub **Wyczyścić**. Nasłuchuje tylko na localhost.
