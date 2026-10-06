# Modely AI a nastavení

## Poskytovatelé a modely

Modely a klíče se nastavují v Nastavení (řádek účtu vlevo dole na domovské stránce):

![Okno Nastavení](img/settings-general.png)

- **Genspark v cloudu**: přihlaste se (postup s kódem zařízení) a používejte — bez žádného nastavení.
- **Vlastní koncové body (BYOK)**: Nastavení ▸ Model AI přijímá základní URL a klíč API pro každý protokol — kompatibilní s OpenAI, Anthropic, Gemini, DeepSeek, DashScope (qwen) a další. Klíče zůstávají pouze v hlavičkách požadavků — nikdy na disku, v protokolech ani v prostředí podřízených procesů.
- Pro každou schopnost lze zvolit jiný model: chat/generování, generování obrázků, analýza obrázků.
- **Otestovat připojení**: ověří dosažitelnost koncového bodu a viditelnost modelu před uložením.
- Základní URL může obsahovat cestu a dotaz (ve stylu brány); cesty koncových bodů se správně připojují.

## Integrace s příkazovou řádkou (třída Codex)

- Nastavení přijímá cestu k místnímu programu CLI (domovské adresáře s neasci znaky a předpona ~ fungují; ~ se rozbalí automaticky); Zjistit modely se dotáže CLI na dostupné modely.
- Ověření kontroluje pouze existenci — žádné omezení znakové sady.

## Kdy se změny projeví

- Změny modelu a koncového bodu působí okamžitě; probíhající konverzace si drží původní nastavení až do svého dalšího kroku.
