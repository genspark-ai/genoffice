# Ustawienia, język, motyw i integracje MCP

## Otwieranie ustawień

Wiersz konta w lewym dolnym rogu strony głównej otwiera panel ustawień (gdy nie jesteś zalogowany, pokazuje Zaloguj się); opcje dotyczące AI znajdują się w jego sekcji Model AI

![Okno Ustawień](img/settings-integrations.png) — konfiguracja modeli jest opisana w rozdziale Modele AI i ustawienia.

## Język

- Ustawienia oferują **21 języków interfejsu**: angielski, uproszczony chiński, japoński, koreański, francuski, niemiecki, hiszpański, tajski, indonejski, rosyjski, arabski, portugalski, włoski, polski, czeski, niderlandzki, malajski, hebrajski, hindi, tradycyjny chiński, wietnamski.
- Zmiana działa natychmiast i zostaje zapamiętana; natywny pasek menu jest odbudowywany w nowym języku.

## Motyw

Jasny / Ciemny / Zgodnie z systemem. Opcja Zgodnie z systemem podąża za wyglądem systemu operacyjnego, a edytory zmieniają wygląd synchronicznie, bez migotania.

## Powiązanie z aplikacją domyślną

Ustawienia mogą zarejestrować GenOffice jako obsługę plików .docx / .xlsx / .pptx / .pdf i podobnych (rejestracja aplikacji domyślnej na poziomie systemu; potwierdzenie, gdy aplikacja o to poprosi).

## Informacje o oprogramowaniu innych firm i aktualizacje

- Pomoc ▸ Informacje o oprogramowaniu innych firm: pełny wykaz licencji open source dołączony do aplikacji.
- Pomoc ▸ Sprawdź aktualizacje…: uruchamia ręczne sprawdzenie; nowsza wersja wyświetla monit o instalacji.

## Logowanie do Genspark

- Punkt logowania (w ustawieniach lub na liście projektów w chmurze) korzysta z przepływu z **kodem urządzenia**: GenOffice pokazuje kod i otwiera stronę logowania w przeglądarce; po zakończeniu proces przebiega dalej automatycznie.
- Logowanie służy wyłącznie do: listy projektów w chmurze i modeli hostowanych przez Genspark. Bez niego wszystkie funkcje lokalne i modele niestandardowe działają dalej.
- Wylogowanie to jedno kliknięcie w ustawieniach.

## Integracja MCP (dla zaawansowanych użytkowników i klientów AI)

GenOffice zawiera lokalny **serwer MCP**, dzięki czemu zewnętrzne klienty AI (Claude Desktop, Cursor, ...) mogą bezpośrednio odczytywać i zapisywać Twoje dokumenty:

- Uruchomienie: `genoffice mcp` w wierszu poleceń (port i token uwierzytelniający są konfigurowalne; domyślnie tylko loopback).
- Możliwości: tworzenie/otwieranie/edycja plików docx, xlsx i pptx, odczyt treści, konwersja formatów, eksport do PDF i więcej — ten sam zestaw narzędzi, którego używają aplikacje desktopowe.
- Bezpieczeństwo: uwierzytelnianie tokenem jest opcjonalne, ale zalecane; nasłuch pozostaje domyślnie na komputerze lokalnym; zobacz `genoffice mcp --help`.

## Ściągawka z poleceń

| Polecenie          | Co robi                      |
| ------------------ | ---------------------------- |
| `genoffice <plik>` | otwiera plik                 |
| `genoffice mcp`    | uruchamia lokalny serwer MCP |
| `genoffice --help` | wszystkie polecenia i opcje  |
