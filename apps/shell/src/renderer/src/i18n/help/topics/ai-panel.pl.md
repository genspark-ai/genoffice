# Panel asystenta AI

Każdy edytor może przywołać panel AI: zaznacz coś, wydaj polecenie i obserwuj wynik przesyłany strumieniowo.

## Otwieranie i używanie

![Panel AI w Docs](img/ai-panel.png)

- Wejścia: **przycisk AI** na wstążce każdego edytora, **Zapytaj AI** w menu kontekstowych albo Zapytaj AI na pasku oznaczeń.
- Opisz zadanie zwykłymi słowami (przepisz to / zamień tę kolumnę na procenty / ułóż tę stronę od nowa...) i naciśnij Enter.
- Odpowiedzi są renderowane **strumieniowo**; gdy AI potrzebuje narzędzi (odczytać dokument, go zmienić, uruchomić skrypt), wykonuje je i kontynuuje aż do końca.
- **Zatrzymaj**: w każdej chwili przerywa bieżącą turę.

## Co potrafi

- **Docs**: przepisywanie/rozbudowa/tłumaczenie/streszczanie, wstawianie tabel i obrazów, zmiana formatowania; każda tura zaczyna się od migawki.
- **Sheets**: formuły, wypełnianie danych, przekształcenia wsadowe (opcjonalnie przez środowisko izolowane run_script), formatowanie.
- **Slides**: generowanie całej prezentacji, korekta układu, przepisywanie tekstów.
- **PDF**: pytania i odpowiedzi oraz streszczenia na podstawie zaznaczonego tekstu lub stron.
- **Markdown / HTML**: przepisywanie, rozbudowa, tłumaczenie.

## Cofanie i bezpieczeństwo

- Panel w Docs prowadzi **listę wersji**: jedna migawka na turę, możesz cofnąć się do dowolnej, a samo cofnięcie da się jeszcze cofnąć przez Ctrl+Z. Miganek przetrwa ponowne otwarcie dokumentu.
- Zmiany AI przechodzą tę samą ścieżkę edycji co zmiany ręczne (można je cofnąć, wymagają zapisania) — nic nie omija potwierdzenia zapisania.

## Prywatność

- Polecenia i odpowiednia treść dokumentu trafiają do **usługi modelu, którą skonfigurowałeś** (Genspark w chmurze lub własny punkt końcowy, następny rozdział); bez konfiguracji nic nie jest wysyłane.
- Pliki lokalne nie są nigdzie indziej przesyłane; klucze BYK pozostają wyłącznie w nagłówkach żądań — nigdy na dysku ani w dziennikach.
