# Operacje na plikach: zmiana nazwy, usuwanie, eksport

Ten rozdział opisuje operacje na plikach wspólne dla wszystkich edytorów; opcje eksportu każdego edytora znajdziesz w jego rozdziale.

**⋯ menu** w wierszu (pojawia się po najechaniu na wiersz pliku) zbiera te działania:

![⋯ menu wiersza pliku](img/file-ops.png)

## Zmiana nazwy

Dwa wejścia, jeden zestaw sprawdzeń:

- **⋯ ▸ Zmień nazwę** w wierszu na stronie głównej.
- **Kliknięcie dwukrotnie karty pliku** zmienia nazwę w miejscu (patrz [Karty i zarządzanie oknami](help://tabs-and-windows)).

Zasady: rozszerzenie jest zachowywane automatycznie; niedozwolone znaki, kropki na końcu i nazwy zastrzeżone (CON, NUL i tym podobne) są odrzucane z komunikatem; konflikt nazwy w tym samym folderze również jest blokowany. Zmieniana jest nazwa prawdziwego pliku na dysku, a listy ostatnich i ulubionych aktualizują się wraz z nią.

## Usuwanie

- **⋯ ▸ Usuń** na stronie głównej: pyta, które pliki, a następnie przenosi je do **kosza systemowego**, skąd system operacyjny może je przywrócić. GenOffice nie trzyma własnego cofnięcia — przywracanie to zadanie kosza, nie komunikatu.

## Duplikowanie

**⋯ ▸ Utwórz kopię** tworzy kopię pliku <nazwa> w tym samym folderze; przy kolizji nazw automatycznie dopisywany jest licznik. Kopia trafia do **Ostatnie**, zamiast otworzyć się sama — jeden klik dalej, a nie przed nosem.

## Zapisz i Zapisz jako

- **⌘S / ctrl+S** zapisuje bieżący plik; plik bez tytułu najpierw pyta o lokalizację i nazwę.
- **Zapisz jako…** zapisuje nowy plik i pozostawia oryginał nietknięty; późniejsze edycje dotyczą nowego pliku.
- Każdy zapis jest atomowy (plik tymczasowy + zmiana nazwy); zamknięcie w trakcie zapisu nie może uszkodzić pliku.
- Autozapis włącza się dopiero po pierwszym zapisaniu ręcznym (patrz [Szybki start](help://getting-started)).

## Eksport do PDF

- **Docs**: Plik ▸ Eksportuj jako PDF… (lub przycisk na wstążce), z podziałem na strony zgodnym z układem.
- **Slides**: eksport rasteryzuje stronę po stronie, z paskiem postępu przy dużych prezentacjach.
- **Sheets**: eksport stosuje podział na strony używany przy drukowaniu.
- Eksporty są renderowane w ukrytym oknie i trafiają tam, gdzie wskażesz.

## Eksport do Worda i obrazów

- **Eksportuj jako Word…** w PDF: zamienia PDF w plik .docx (konwersja lokalna; złożone układy odtwarzane są w miarę możliwości).
- **Docs** może eksportować strony jako obrazy (PNG na stronę).

## Drukowanie

Plik ▸ Drukuj… w każdym edytorze (⌘P/ctrl+P) otwiera systemowe okno drukowania; pliki PDF drukują się z bieżącą kolejnością stron i obrotami.

## Gdzie trafiają pliki bez tytułu

Lokalizacja wybrana przy pierwszym zapisaniu jest ich domem; do tego czasu dokument istnieje wyłącznie w pamięci. Autozapis przejmuje nad nim kontrolę dopiero po tym pierwszym zapisaniu.
