# Czcionki: fonty systemowe i rodziny do pobrania

## Listy czcionek

Elementy do wyboru czcionek znajdują się w karcie Narzędzia główne każdego edytora (poniżej: grupa czcionek w Slides):

![Grupa czcionek na wstążce Slides](img/fonts.png)

Selektor czcionek każdego edytora scala: czcionki zainstalowane lokalnie + typowe kandydaty platformy (najpopularniejsze rodziny Windows i macOS, zachodnie i CJK, wraz z nazwami zlokalizowanymi, pod jakimi czcionki CJK zgłaszają się w swoich systemach). Gdy system potrafi wyliczyć czcionki lokalne, lista jest grupowana według tego, co faktycznie istnieje; w przeciwnym razie przechodzi na pełną listę kandydatów.

## Czcionki do pobrania (Slides)

- Selektor czcionek w Slides otwiera **katalog czcionek**: wyselekcjonowany zestaw OFL obejmujący wiele systemów pisma; każda rodzina występuje w wersji zwykłej i pogrubionej.
- Wybór powoduje pobranie i instalację z CDN (z przypiętym sumą kontrolną) do magazynu czcionek aplikacji — od tej chwili dostępne dla każdego dokumentu, bez instalacji na poziomie systemu.
- Instalacje mieszkają w katalogu prywatnym aplikacji i znikają razem z nią.

## Typografia CJK

- Czcionki CJK w Docs otrzymują kompresję interpunkcji i reguły łamania wierszy (kinsoku) zgodne z Wordem (patrz rozdział o Docs).
- Rodziny CJK w katalogu obejmują chiński uproszczony i tradycyjny, japoński i koreański, kroje szeryfowe i bezszeryfowe.

## Zarządzanie czcionkami lokalnymi

- Pozycja pozwalająca zainstalować lokalne pliki czcionek wczytuje pliki .ttf/.otf z dysku do magazynu czcionek aplikacji.
- Magazyn jest uporządkowany według rodzin; kilka wag tej samej rodziny współistnieje.
