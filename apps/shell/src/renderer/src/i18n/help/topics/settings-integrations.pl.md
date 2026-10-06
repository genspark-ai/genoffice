# Ustawienia, język, motyw i integracje MCP

## Otwieranie ustawień

Wiersz konta w lewym dolnym rogu strony głównej otwiera panel ustawień (gdy nie jesteś zalogowany, pokazuje Zaloguj się). Ma sześć sekcji: Konto, Model AI, Media i wyszukiwanie AI, Ogólne, Integracje i O aplikacji.

![Ustawienia ▸ Ogólne, gdzie znajdują się język, motyw, automatyczny zapis i przełącznik statystyk użycia](img/settings-general.png)

Konfiguracja modeli ma własny artykuł; w sekcji **Media i wyszukiwanie AI** włączasz dla każdego dostawcy generowanie obrazów, analizę obrazów, analizę wideo, wyszukiwanie w sieci i wyszukiwanie plików lokalnych.

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

Do **Integracji** łączy GenOffice z agentem programistycznym i ma własny artykuł: Podłączenie z agentem programistycznym. Wersja skrócona — wybierz drogę (wiersz poleceń albo MCP), przejdź tę część, a potem rozpocznij nową rozmowę i zapytaj.

![Ustawienia ▸ Integracje: trzy kroki, potem wiersze umiejętności i opcje MCP](img/settings-integrations.png)

W sekcji **Lokalny serwer HTTP** aplikacja może też sama uruchomić ten serwer — przełącznik włączania i port —, a **Zaawansowane** dodaje adres sprawdzenia stanu i plik dziennika, zamiast zostawiać serwer asystentowi. Nasłuchuje wyłącznie na localhost.

## Ściągawka z poleceń

| Polecenie          | Co robi                      |
| ------------------ | ---------------------------- |
| `genoffice <plik>` | otwiera plik                 |
| `genoffice mcp`    | uruchamia lokalny serwer MCP |
| `genoffice --help` | wszystkie polecenia i opcje  |
