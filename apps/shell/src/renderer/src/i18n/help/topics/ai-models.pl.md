# Modele AI i ustawienia

## Dostawcy i modele

Modele i klucze konfiguruje się w ustawieniach (wiersz konta w lewym dolnym rogu strony głównej):

![Okno Ustawień](img/settings-general.png)

- **Genspark w chmurze**: zaloguj się (przepływ z kodem urządzenia) i korzystaj — bez żadnej konfiguracji.
- **Własne punkty końcowe (BYOK)**: Ustawienia ▸ Model AI przyjmuje adres bazowy i klucz API dla każdego protokołu — zgodny z OpenAI, Anthropic, Gemini, DeepSeek, DashScope (qwen) i inne. Klucze pozostają wyłącznie w nagłówkach żądań — nigdy na dysku, w dziennikach ani w środowisku procesów potomnych.
- Dla każdej funkcji można wybrać inny model: czat/generowanie, generowanie obrazów, analiza obrazów.
- **Testuj połączenie**: sprawdza osiągalność punktu końcowego i widoczność modelu przed zapisaniem.
- Adresy bazowe mogą zawierać ścieżkę i ciąg zapytania (w stylu bramy); ścieżki punktów końcowych są dokładnie doklejane.

## Integracja z wierszem poleceń (klasy Codex)

- Ustawienia przyjmują ścieżkę lokalnego programu CLI (katalogi domowe o znakach spoza ASCII i prefiks ~ działają; ~ jest rozwijany automatycznie); Wykryj modele odpytuje CLI o dostępne modele.
- Walidacja sprawdza wyłącznie istnienie — bez ograniczeń dla zestawu znaków.

## Kiedy zmiany zaczynają obowiązywać

- Zmiany modelu i punktu końcowego działają natychmiast; trwająca rozmowa zachowuje poprzednią konfigurację do swojej następnej tury.
