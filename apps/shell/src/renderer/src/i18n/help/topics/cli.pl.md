# Wiersz poleceń i agenci

Każda instalacja zawiera polecenie `genoffice`, które napędza te same silniki co okno — te same parsery, ten sam moduł zapisu, ten sam renderujący. Plik zapisany przez aplikację i plik zapisany przez polecenie to ten sam plik, a kontrola, którą przechodzi panel AI aplikacji, przechodzi też polecenie.

```sh
genoffice --help          # every command
genoffice guide slides    # the op reference, for writing your own
```

To cała powierzchnia na jednym ekranie — każde polecenie z wierszem, który mówi, co robi, a potem opcje globalne i kody wyjścia:

![Rzeczywiste wyjście polecenia genoffice --help: banner wersji, każde polecenie z opisem w jednym wierszu oraz opcje globalne i kody wyjścia](img/cli.png)

## Zdobycie polecenia

Systemy macOS i Windows dostarczają je w pakiecie aplikacji. Aby używać go po nazwie, uruchom raz `genoffice install-cli`: tworzy dowiązanie symboliczne do dołączonego pliku binarnego w `/usr/local/bin` albo w `PATH` użytkownika w systemie Windows.

## Polecenia warte znajomości

| Polecenie         | Co robi                                                                                                                                                                                                                                  |
| ----------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `open`            | Otwiera dokument w aplikacji; uruchamia aplikację, jeśli nie działa.                                                                                                                                                                     |
| `selection`       | To, co użytkownik właśnie zaznaczył w tym pliku, gdy aplikacja działa — wskaźnik „to” / „tutaj”. Zwraca zakres bloków, zakres komórek w arkuszu, elementy slajdu albo jedną stronę, w zależności od edytora, w którym znajduje się plik. |
| `convert`         | Konwertuje między formatami przy użyciu silników aplikacji.                                                                                                                                                                              |
| `create`          | Tworzy dokument z treści strukturalnej.                                                                                                                                                                                                  |
| `render`          | Jeden PNG na stronę, tak jak układa je renderujący.                                                                                                                                                                                      |
| `pdf`             | Odczytuje warstwę tekstową PDF-a strona po stronie, bez procesu aplikacji.                                                                                                                                                               |
| `info`            | Metadane i podsumowanie struktury dokumentu.                                                                                                                                                                                             |
| `search`          | Wyszukiwanie w sieci lub obrazów przez dostawcę skonfigurowanego w aplikacji.                                                                                                                                                            |
| `image` / `media` | Generuje obraz albo opisuje plik obrazu, wideo lub dźwięk i pozwala o niego pytać.                                                                                                                                                       |
| `merge`           | Wypełnia symbole zastępcze `{{key}}` w szablonie `.docx`, `.pptx` lub `.xlsx`.                                                                                                                                                           |
| `capabilities`    | Zgłasza, które funkcje chmurowe są skonfigurowane na tym komputerze.                                                                                                                                                                     |
| `guide`           | Referencja op i przewodniki projektowe, wygenerowane z tych samych definicji, względem których waliduje executor — więc nie może się rozjechać z tym, co przyjmuje `apply`. `--json` zwraca je ze schematem każdej op.                   |
| `install-cli`     | Umieszcza `genoffice` w `PATH`.                                                                                                                                                                                                          |
| `skill`           | Wypisuje agentów programistycznych znalezionych na tym komputerze i instaluje w nich lub aktualizuje skill GenOffice.                                                                                                                    |
| `mcp`             | Udostępnia każde polecenie jako narzędzie Model Context Protocol. Zobacz **Podłączenie z agentem programistycznym**.                                                                                                                     |

## Edycja: Docs, Sheets, Slides

`genoffice docs`, `genoffice sheet` i `genoffice slides` czytają i edytują przez tę samą ścieżkę zapisu co aplikacja i dzielą jedno słownictwo: **op** to pojedyncza zmiana, a **spec** to lista op zastosowanych po kolei.

```sh
genoffice sheet read report.xlsx
genoffice slides apply deck.pptx --ops ops.json --dry-run
genoffice docs apply brief.docx --ops ops.json
```

`--dry-run` zgłasza, co zrobiłaby partia, i nic nie zapisuje — to najtańszy sposób na sprawdzenie spec przed zastosowaniem. Panel AI w aplikacji działa dokładnie na tych op, więc wszystko, o co możesz go zapytać, możesz też zapisać w skrypcie.

## Model Context Protocol

`genoffice mcp` udostępnia każde polecenie jako narzędzie MCP, a `genoffice mcp install <agent|all>` rejestruje je we własnej konfiguracji agenta programistycznego. Tej strony dotyczy rozdział **Podłączenie z agentem programistycznym**.

## Czego polecenie nie robi

Czyta i zapisuje plik. To nie jest aplikacja: nie ma okna, a okno aktualizacji w aplikacji nie ma tu zastosowania. Wszystko, co wymaga okna — panel AI, kontrola jakości wyrenderowanej slajdy — musi poczekać, aż otworzysz plik. `genoffice render` daje ci piksele bez okna, a `genoffice slides` sam sprawdza układ prezentacji.
