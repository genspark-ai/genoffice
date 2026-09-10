<p align="center">
  <a href="https://genoffice.ai/">
    <picture>
      <source srcset="../assets/readme/hero-dark.webp" media="(prefers-color-scheme: dark)">
      <img src="../assets/readme/hero.webp" alt="GenOffice — свободный AI-офисный пакет с открытым исходным кодом: Docs, Sheets, Slides, PDF, Markdown и HTML со встроенной AI-панелью" width="100%">
    </picture>
  </a>
</p>

<h1 align="center">GenOffice</h1>

<p align="center"><b>Первый в мире полнофункциональный офисный пакет с открытым исходным кодом и AI.</b><br>
Файлы Word, Excel, PowerPoint и PDF, редактируемые вами и вашим AI, сохраняются обратно в реальных форматах.</p>

<p align="center">
  <a href="../../LICENSE"><img src="https://img.shields.io/github/license/genspark-ai/genoffice" alt="Лицензия: Apache-2.0"></a>
  <a href="https://github.com/genspark-ai/genoffice/releases/latest"><img src="https://img.shields.io/github/v/release/genspark-ai/genoffice" alt="Последний релиз"></a>
  <a href="https://github.com/genspark-ai/genoffice/releases"><img src="https://img.shields.io/github/downloads/genspark-ai/genoffice/total" alt="Загрузки"></a>
  <a href="https://github.com/genspark-ai/genoffice/stargazers"><img src="https://img.shields.io/github/stars/genspark-ai/genoffice?style=flat" alt="Звёзды на GitHub"></a>
</p>

<p align="center"><a href="../../README.md">English</a> · <a href="README.zh-CN.md">简体中文</a> · <a href="README.zh-TW.md">繁體中文</a> · <a href="README.ja.md">日本語</a> · <a href="README.ko.md">한국어</a> · <a href="README.de.md">Deutsch</a> · <a href="README.fr.md">Français</a> · <a href="README.es.md">Español</a> · <a href="README.pt-BR.md">Português (Brasil)</a> · <b>Русский</b> · <a href="README.it.md">Italiano</a> · <a href="README.nl.md">Nederlands</a> · <a href="README.pl.md">Polski</a> · <a href="README.cs.md">Čeština</a> · <a href="README.id.md">Bahasa Indonesia</a> · <a href="README.ms.md">Bahasa Melayu</a> · <a href="README.th.md">ไทย</a> · <a href="README.hi.md">हिन्दी</a> · <a href="README.ar.md">العربية</a> · <a href="README.he.md">עברית</a></p>

<p align="center">
  <a href="#download"><b>Скачать</b></a> ·
  <a href="https://genoffice.ai/"><b>Сайт</b></a> ·
  <a href="https://genoffice.ai/join"><b>Сообщество</b></a> ·
  <a href="../../PRIVACY.md"><b>Конфиденциальность</b></a>
</p>

GenOffice — бесплатная альтернатива Microsoft Office с открытым исходным
кодом для macOS, Windows и Linux. Она открывает и сохраняет нативные файлы
`.docx`, `.xlsx` и `.pptx`, редактирует PDF, Markdown и HTML и размещает
AI-агента рядом с каждым документом — не чат-бокс, прикрученный сбоку, а
редактор, который читает файл, вносит изменение и точно показывает вам, что
именно он затронул.

- **Настоящие форматы, побайтовое сохранение.** Переписывается только то, что
  вы редактируете. Всё остальное в файле сохраняется байт в байт, поэтому
  документы продолжают открываться в Word, Excel и PowerPoint.
- **AI, чью работу можно проверить.** Изменения оформляются как отслеживаемые
  правки и diff-ы с откатом одним щелчком. В таблицах появляются живые
  формулы, а не вставленные числа. Презентации и страницы создаются прямо на
  холсте и остаются полностью редактируемыми.
- **Локальность по своей природе.** Файлы открываются, редактируются,
  сохраняются и конвертируются на вашем компьютере. Преобразования PDF →
  Word / Excel / PowerPoint, Markdown → Word и HTML → Word выполняются прямо
  на устройстве. За пределы вашего компьютера уходят только запросы к AI — к
  выбранному вами провайдеру.
- **Свои ключи или никаких.** Войдите через Genspark и обойдитесь без ключей
  вообще, либо используйте собственный ключ для Claude, OpenAI, Gemini,
  DeepSeek, Kimi, GLM, Qwen, Doubao, MiniMax, Grok, Mistral, OpenRouter или
  любого совместимого с OpenAI эндпоинта, включая локальные серверы.

**Скачать:** [macOS](https://github.com/genspark-ai/genoffice/releases/latest) (Apple Silicon и Intel) ·
[Windows](https://github.com/genspark-ai/genoffice/releases/latest) (x64 и Arm) ·
[Linux](https://github.com/genspark-ai/genoffice/releases/latest) (deb, rpm, AppImage) —
подробности и системные требования в разделе [Скачать](#download).

## Демонстрация

Шесть приложений, одна AI-панель. Каждый скриншот — это реальное приложение
на macOS, а действия AI запускаются промптом, который вы можете прочитать на
панели.

### 1 · Docs — открывайте и редактируйте `.docx` с AI, чью работу можно проверить

<table>
<tr>
<td width="50%"><img src="../assets/readme/docs-report.webp" alt="GenOffice Docs отображает страницу годового отчёта в две колонки с полноширинным изображением на обложке, затемнённой таблицей KPI, верхним и нижним колонтитулами при масштабе 80% и свёрнутой AI-панелью"></td>
<td width="50%"><img src="../assets/readme/docs-ai.webp" alt="GenOffice Docs: обзор компании с изображением-баннером; AI сократил раздел Overview и добавил новый раздел со списком, а панель предлагает откат одним щелчком"></td>
</tr>
<tr>
<td><b>Открывает файл так же, как его вёрстку задаёт Word</b> — двухколоночные разделы, изображения на всю страницу, затемнённые таблицы, колонтитулы, разбиение на страницы по метрикам строк Word. Стили, комментарии, отслеживаемые изменения, формулы и рукописный ввод переживают round-trip без изменений.</td>
<td><b>Попросите внести изменение</b> — AI считывает нужные блоки, переписывает раздел Overview и добавляет новый раздел со списком. Каждый шаг AI сохраняется как снимок, который можно откатить; при включённом режиме <b>Track changes</b> правки приходят как ревизии в стиле Word.</td>
</tr>
</table>

### 2 · Sheets — `.xlsx` с живыми формулами и диаграммами, а не вставленными числами

<table>
<tr>
<td width="50%"><img src="../assets/readme/sheets-ai.webp" alt="GenOffice Sheets: AI добавил лист Summary с доходом по регионам и категориям на формулах SUMIF, плюс столбчатую диаграмму, и сообщает о 43 применённых изменениях с кнопкой отмены"></td>
<td width="50%"><img src="../assets/readme/sheets-qa.webp" alt="GenOffice Sheets: на вопрос о том, какой регион лидировал по доходу во втором квартале, AI отвечает «Европа» с разбивкой по категориям и указывает точные ячейки как ссылки-цитаты, рядом с листом Orders"></td>
</tr>
<tr>
<td><b>Постройте это</b> — из одного предложения агент добавляет лист Summary с настоящими формулами <code>SUMIF</code> по регионам и категориям, вставляет столбчатую диаграмму и применяет все 43 изменения одним отменяемым пакетом.</td>
<td><b>Спросите</b> — вопросы о книге возвращаются с объяснением и точными ячейками в виде кликабельных цитат. Под капотом: собственный Rust-движок `.xlsx`, сводные таблицы, срезы, условное форматирование и трассировка формул.</td>
</tr>
</table>

### 3 · Slides — от промпта до готовой презентации `.pptx`

<img src="../assets/readme/slides-generate.webp" alt="Таймлапс создания презентации GenOffice Slides для инвесторской колоды Aurora Home: ИИ планирует сюжет в панели, слайды появляются на холсте один за другим, а готовая презентация завершается финальным призывом к действию" width="100%">

<table>
<tr>
<td width="50%"><img src="../assets/readme/slides-cover.webp" alt="GenOffice Slides: титульный слайд презентации для инвесторов Aurora Home, созданной AI, на холсте, с исходным однострочным промптом и резюме AI о том, что было построено, на панели"></td>
<td width="50%"><img src="../assets/readme/slides-ai.webp" alt="GenOffice Slides: оформленный финальный слайд той же презентации из 11 слайдов, с полосой миниатюр слева и AI-панелью, подытоживающей сюжетную линию"></td>
</tr>
<tr>
<td><b>Одна строка на входе</b> — «Создай презентацию для инвесторов из 10 слайдов для Aurora Home…». GenOffice выстраивает сюжетную линию, собирает цифры и создаёт каждый слайд прямо на холсте как настоящий файл <code>.pptx</code>.</td>
<td><b>Готовая презентация на выходе</b> — одиннадцать оформленных слайдов с единой типографикой, иллюстрациями и завершающим призывом к действию; продолжайте редактировать с помощью мастер-слайдов, макетов, умных направляющих и неразрушающей обрезки, либо попросите панель изменить стиль, переписать текст или переставить слайды.</td>
</tr>
</table>

### 4 · PDF — редактируйте текст PDF на месте, конвертируйте PDF в Word на устройстве

<table>
<tr>
<td width="50%"><img src="../assets/readme/pdf-edit.webp" alt="GenOffice PDF: режим редактирования текста обводит каждый текстовый блок на странице для редактирования на месте, а AI-панель отвечает на вопрос об отчёте со ссылками на страницы"></td>
<td width="50%"><img src="../assets/readme/pdf-convert.webp" alt="GenOffice Docs показывает документ Word, конвертированный локально из квартального отчёта Helios в формате PDF, открытый во второй вкладке рядом с исходным PDF"></td>
</tr>
<tr>
<td><b>Редактируйте прямо на странице</b> — режим редактирования текста обводит каждый текстовый блок для повторного набора на месте; поток содержимого страницы переписывается через PDFium с сохранением исходных шрифтов, а не поверх маскирующей аннотацией. Спросите AI о длинном отчёте и получите ответ со ссылками на страницы.</td>
<td><b>Конвертируйте на устройстве</b> — <b>PDF Converter → PDF to Word</b> создаёт редактируемый файл <code>.docx</code>, который открывается в Docs рядом с исходником, с сохранёнными заголовками, строками статистики и абзацами. Excel и PowerPoint работают так же; отсканированные страницы проходят через системный OCR.</td>
</tr>
</table>

### 5 · HTML — AI-конструктор страниц и интерфейсов, начиная с дизайн-брифа

Скажите, для чего страница и для кого она. AI сначала предлагает
**дизайн-бриф** — крючок, палитру, типографику и стилевые направления — а
затем собирает единый самодостаточный файл `.html` на основе этих токенов.

<img src="../assets/readme/html-restyle-motion.webp" alt="Таймлапс изменения стиля GenOffice HTML для лендинга Lumen: один запрос Restyle в панели превращает тёмную страницу Midnight Studio в тёплую версию Solar Daybreak, при этом все разделы и весь текст остаются на месте" width="100%">

<table>
<tr>
<td width="50%"><img src="../assets/readme/html-ai.webp" alt="GenOffice HTML: сгенерированный лендинг для настольной лампы на солнечных батареях в тёмном направлении Midnight Studio, показанный в живом предпросмотре, с AI-панелью, резюмирующей только что созданную страницу"></td>
<td width="50%"><img src="../assets/readme/html-restyle.webp" alt="Тот же лендинг Lumen, переоформленный AI в тёплом направлении Solar Daybreak: бумажный фон, серифные заголовки и оранжевый акцент, с сохранением всех разделов и всего текста"></td>
</tr>
<tr>
<td><b>Сгенерировано из одного промпта</b> — яркий hero-блок, карточки функций, тарифы и форма списка ожидания для Lumen, созданные в направлении Midnight Studio. Кликните на любой элемент, чтобы изменить его стиль, дважды кликните, чтобы отредактировать текст, или переключитесь на исходный код в CodeMirror.</td>
<td><b>Тот же дизайн, новое направление</b> — один запрос <b>Restyle</b> заменяет токены брифа, и страница подстраивается под них: тёплая бумага, редакционный серифный шрифт, солнечно-оранжевый акцент — ничего не переписывается заново. Показывайте презентацию в полноэкранном режиме или экспортируйте в PDF либо в нативный редактируемый документ Word.</td>
</tr>
</table>
<table>
<tr>
<td width="50%"><img src="../assets/readme/html-dashboard.webp" alt="GenOffice HTML: сгенерированный интерфейс личного дашборда для дизайнера-фрилансера в тёплом «льняном» стиле, с боковой панелью, серифным приветствием и четырьмя карточками показателей"></td>
<td width="50%"><img src="../assets/readme/html-report.webp" alt="GenOffice HTML: сгенерированный отчёт о рынке электромобилей в стиле широкоформатной газеты, с серифной шапкой, заголовной цифрой 17,3 миллиона и строкой статистики"></td>
</tr>
<tr>
<td><b>Макеты интерфейсов</b> — стартовый шаблон «личный дашборд» превращает персону в рабочий макет: боковая панель, приветствие, спарклайн оплачиваемых часов, карточки счетов и загрузки — всё в виде настоящего HTML, который можно передать разработчику.</td>
<td><b>Истории на данных</b> — стартовый шаблон «отчёт по данным» строит редакционную газетную полосу: серифная шапка, одна заголовная цифра, строка статистики с разделительными линиями, встроенные SVG-диаграммы и примечание о методологии.</td>
</tr>
</table>

### 6 · Markdown — блочный редактор для обычного `.md`, с функцией Ask AI

<table>
<tr>
<td width="50%"><img src="../assets/readme/markdown-ai.webp" alt="GenOffice Markdown: у выделенного абзаца всплывает окно Ask AI с введённой инструкцией и чипами-подсказками, такими как «Отполировать», «Сделать короче», «Развернуть» и «Исправить грамматику», а также кнопками «Отправить сейчас» и «Добавить в очередь»"></td>
<td width="50%"><img src="../assets/readme/markdown-render.webp" alt="GenOffice Markdown отображает документ с заметками о релизе, содержащий таблицу, блок-схему Mermaid и список задач, со стартовыми промптами AI-панели слева"></td>
</tr>
<tr>
<td><b>Спросите AI о выделенном фрагменте</b> — выделите любой отрывок, и появится чип <b>Ask AI</b>: введите инструкцию или выберите подсказку, отправьте её сразу или поставьте несколько привязанных правок в очередь и выполните их за один проход. Такая же точка входа есть в каждом приложении.</td>
<td><b>Отрисовано, сохраняется как обычный Markdown</b> — заголовки, списки, таблицы, изображения, блоки кода и диаграммы Mermaid в блочном редакторе Tiptap, записываются обратно как обычный файл <code>.md</code>, с полностью локальным экспортом <b>Markdown → Word</b>.</td>
</tr>
</table>

## Почему GenOffice

- **Открытый исходный код**, Apache-2.0, разрабатывается открыто на GitHub.
- **Ваш собственный.** Нативные приложения для macOS, Windows и Linux; файлы
  остаются на вашем диске, а каждое редактирование, сохранение и конвертация
  происходят на вашем устройстве.
- **Настоящие файлы Office.** Нативные `.docx`, `.xlsx` и `.pptx` с побайтовым
  сохранением: нетронутые части файла копируются в точности как были.
- **AI, редактирующий сам документ.** Отслеживаемые изменения в Docs, живые
  формулы и диаграммы в Sheets, слайды рисуются прямо на холсте, каждый шаг AI
  создаёт снимок, к которому можно вернуться.
- **Своя модель, свой ключ.** Войдите через Genspark или используйте свой
  ключ для Claude, OpenAI, Gemini, DeepSeek и других, включая локальные
  серверы и любой совместимый с OpenAI эндпоинт.
- **PDF, сделанный правильно.** Редактируйте текст прямо на странице и
  конвертируйте PDF в Word, Excel или PowerPoint на устройстве, со встроенным
  в систему OCR для сканов.
- **А также Markdown и HTML**, с той же панелью AI и локальным экспортом в
  Word.
- **Бесплатно**, как для отдельных пользователей, так и для команд.

## AI-провайдеры

**Войдите через Genspark** — и настраивать ничего не нужно: вызовы моделей
идут через проксирующий сервис Genspark (семейства Claude, GPT и Gemini), а
агенты получают веб- и графический поиск, генерацию изображений и анализ
изображений, аудио и видео.

**Либо используйте собственный ключ.** В разделе Settings → AI перечислены
Claude, OpenAI, Gemini, DeepSeek, Kimi, GLM, Qwen, Doubao, MiniMax, Grok,
Mistral, OpenRouter и OpenCode Zen/Go, а также отдельный слот для любого
совместимого с OpenAI эндпоинта (базовый URL + ключ), включая локальные
серверы моделей. Поиск и работа с медиа настраиваются через отдельных
провайдеров по каждой возможности в разделе **AI Media & Search**: Serper
или Tavily для веб-поиска и OpenAI, Gemini, Doubao/Seedream, GLM, Grok, Qwen,
MiniMax или любой совместимый с OpenAI эндпоинт для генерации изображений и
анализа изображений/видео.

Весь пакет поддерживает светлую, тёмную и системную темы. Темы меняют только
то, что видно на экране: экспорт, печать и сохранённые файлы всегда сохраняют
собственные цвета документа.

<a id="download"></a>

## Скачать

| Платформа                            | Требования                                            | Скачать                                                                                       |
| ------------------------------------ | ----------------------------------------------------- | --------------------------------------------------------------------------------------------- |
| **macOS** — Apple Silicon (arm64)    | macOS 11+                                             | [Последний `.dmg` (arm64)](https://github.com/genspark-ai/genoffice/releases/latest)          |
| **macOS** — Intel (x64)              | macOS 11+                                             | [Последний `.dmg` (x64)](https://github.com/genspark-ai/genoffice/releases/latest)            |
| **Windows** (x64, большинство ПК)    | Windows 10+, Intel/AMD                                | [Последний установщик `-x64.exe`](https://github.com/genspark-ai/genoffice/releases/latest)   |
| **Windows** на Arm (ARM64)           | Windows 11 на Arm (Snapdragon X и аналогичные)        | [Последний установщик `-arm64.exe`](https://github.com/genspark-ai/genoffice/releases/latest) |
| **Linux** — Debian / Ubuntu          | x86_64, glibc 2.34+ (Ubuntu 22.04 или новее)          | [Последний `.deb`](https://github.com/genspark-ai/genoffice/releases/latest)                  |
| **Linux** — Fedora / RHEL / openSUSE | x86_64, glibc 2.34+ (Fedora 35+, RHEL 9+, Leap 15.6+) | [Последний `.rpm`](https://github.com/genspark-ai/genoffice/releases/latest)                  |
| **Linux** — другие дистрибутивы      | x86_64, glibc 2.34+, FUSE 2                           | [Последний `.AppImage`](https://github.com/genspark-ai/genoffice/releases/latest)             |

Все сборки собираются из ветки `main`; установщики для macOS и Windows
подписаны. Более старые версии доступны на странице
[Releases](https://github.com/genspark-ai/genoffice/releases).

<details>
<summary><b>Установка на Linux</b></summary>

Deb-пакет устанавливается через apt — он подтягивает зависимости и
добавляет GenOffice в меню приложений:

```bash
sudo apt install ./genoffice_<version>_amd64.deb
```

На Fedora / в семействе RHEL / openSUSE вместо этого установите rpm-пакет:

```bash
sudo dnf install ./genoffice-<version>.x86_64.rpm     # Fedora / RHEL family
sudo zypper install ./genoffice-<version>.x86_64.rpm  # openSUSE
```

AppImage запускается прямо на месте: установите runtime FUSE 2
(`sudo apt install libfuse2`; на Ubuntu 24.04 пакет называется
`libfuse2t64`), сделайте файл исполняемым и запустите его:

```bash
chmod +x GenOffice-<version>.AppImage
./GenOffice-<version>.AppImage
```

</details>

## Как это работает

Семь Electron-приложений — Docs, Sheets, Slides, PDF, Markdown, HTML и
общая оболочка с вкладками — используют один слой чистых TypeScript-пакетов
плюс Rust-компонент (sidecar) для `.xlsx`. Исходный файл всегда остаётся
источником истины: изменения применяются как узкие патчи, а всё, что
редактор не трогал, переживает цикл сохранения без изменений.

```
open docx ─► архивируем оригинал по хэшу (никогда не трогаем)
          ─► разбираем word/document.xml в дерево блоков, каждый блок привязан к исходному XML
          ─► редактор Tiptap (ручное и AI-редактирование, отслеживание изменённого)
save      ─► изменённые блоки → фрагменты OOXML (со ссылками только на существующие стили)
          ─► вклеиваем их в исходный document.xml; нетронутые блоки сохраняют свои байты
          ─► пересобираем zip-архив; все остальные записи копируются побайтово
```

Обзор пакетов по отдельности (движки docx/pptx, `pdf2docx`, `html2docx`,
ядро агента и провайдеры) находится в
[CONTRIBUTING.md](../../CONTRIBUTING.md#engine-packages).

## Разработка

```bash
npm install
npm run fixtures     # generate test .docx fixtures
npm test             # engine + app unit tests (docs/sheets/slides need no display)
npm run typecheck    # tsc --noEmit across every workspace
npm run dev          # all six editors + shell against Vite dev servers
npm run dev:docs     # a single app (same pattern works per workspace)
npm run dist:mac     # package macOS dmg (regenerates third-party notices)
npm run dist:win     # package Windows nsis installer
npm run dist:linux   # package Linux AppImage + deb + rpm
```

Приложению Sheets также нужен инструментарий Rust для его xlsx sidecar
(`cargo` в PATH); команда `npm run build -w @genoffice/sheets` собирает его
автоматически. Список проверок, которые должно проходить каждое изменение, и
порядок принятия pull request-ов — в [CONTRIBUTING.md](../../CONTRIBUTING.md).

## Сообщество

GenOffice активно развивается, и ваши отзывы влияют на его развитие.

- **Сообщите об ошибке или предложите функцию** в
  [GitHub Issues](https://github.com/genspark-ai/genoffice/issues).
- **Присоединяйтесь к групповому чату GenOffice** в
  [GenTeam](https://genoffice.ai/join), чтобы общаться с командой и другими
  пользователями.
- **Поставьте звезду репозиторию**, если GenOffice вам полезен — это лучший
  способ поддержать проект.

## Часто задаваемые вопросы

<details>
<summary><b>GenOffice бесплатен?</b></summary>

Да. GenOffice бесплатен и распространяется с открытым исходным кодом по
лицензии Apache-2.0 — ни пробного периода, ни платного тарифа для самих
приложений.

</details>

<details>
<summary><b>Может ли GenOffice открывать файлы Microsoft Word, Excel и PowerPoint?</b></summary>

Да. GenOffice открывает и сохраняет нативные файлы `.docx`, `.xlsx` и
`.pptx`. Сохранение выполняется побайтово: части файла, которые вы не
трогали, записываются обратно байт в байт, поэтому документы продолжают
работать в Microsoft Office.

</details>

<details>
<summary><b>Работает ли GenOffice офлайн?</b></summary>

Редактирование документов полностью локально — файлы никогда не покидают
ваш компьютер при открытии, редактировании, сохранении или конвертации.
Функции AI (агенты, поиск, инструменты для изображений) требуют подключения
к сети — либо через вход в Genspark, либо через ваш собственный ключ API
модели.

</details>

<details>
<summary><b>Может ли GenOffice редактировать файлы PDF?</b></summary>

Да — это настоящее редактирование текста и изображений PDF, которое
переписывает поток содержимого страницы с сохранением исходных шрифтов, а не
накладывает маскирующие аннотации.

</details>

<details>
<summary><b>Может ли GenOffice конвертировать PDF в Word, Excel или PowerPoint?</b></summary>

Да — полностью на устройстве: посимвольное извлечение через PDFium плюс
анализ вёрстки на основе геометрии, без облачных сервисов и без загрузки
файлов куда-либо. Отсканированные страницы тоже поддерживаются: на macOS и
Windows их читает системный OCR, поэтому они конвертируются в редактируемый
текст, а не в изображение страницы.

</details>

<details>
<summary><b>Могу ли я использовать свою AI-модель или API-ключ?</b></summary>

Да. Помимо входа через Genspark без ключей, GenOffice поддерживает
собственный ключ для Claude, OpenAI, Gemini, DeepSeek, Kimi, GLM, Qwen,
Doubao, MiniMax, Grok, Mistral, OpenRouter и OpenCode Zen/Go, а также любой
совместимый с OpenAI эндпоинт — включая локальные серверы моделей. Поиск,
генерация изображений и анализ изображений/видео используют собственные
ключи в разделе Settings → AI Media & Search.

</details>

<details>
<summary><b>Может ли GenOffice конвертировать HTML в Word?</b></summary>

Да — функция Export as Word в приложении HTML создаёт нативный редактируемый
файл `.docx` полностью на устройстве. Страница отрисовывается во встроенном
Chromium и сводится к настоящим структурам Word: заголовкам, абзацам,
спискам, таблицам, карточкам, строкам KPI, полям форм и фонам страниц; в виде
изображений встраивается только то, что не имеет аналога в Word (диаграммы,
иконки, декорированные блоки).

</details>

<details>
<summary><b>Собирает ли GenOffice какие-либо данные?</b></summary>

Официальные собранные сборки по умолчанию отправляют ограниченную
аналитику использования, и вы можете отключить её в любой момент в разделе
Settings → General. Аналитика никогда не передаёт содержимое документов,
имена и пути файлов, идентификатор учётной записи или адреса электронной
почты. Полный перечень событий и раскрытие данных — в
[GenOffice Privacy](../../PRIVACY.md).

</details>

## Безопасность

Описание модели безопасности процессов (изоляция renderer-процесса,
проверка IPC, ограничение внешних ссылок) и модели угроз для контента,
созданного AI, — в [SECURITY.md](../../SECURITY.md).

## Благодарности

GenOffice был бы невозможен без этих проектов с открытым исходным кодом:

- [Electron](https://www.electronjs.org/) — среда выполнения для каждого
  приложения.
- [Univer](https://github.com/dream-num/univer) (Apache-2.0) — основа
  интерфейса таблиц, на которой строится Sheets.
- [PDFium](https://pdfium.googlesource.com/pdfium/) (BSD-3-Clause,
  поставляется через [@embedpdf/pdfium](https://github.com/embedpdf/embed-pdf-viewer))
  — движок потока содержимого, стоящий за настоящим редактированием текста и
  изображений в PDF.
- [pdf.js](https://github.com/mozilla/pdf.js) (Apache-2.0) и
  [pdf-lib](https://github.com/Hopding/pdf-lib) (MIT) — рендеринг PDF и
  сборка документов.
- [Tiptap](https://tiptap.dev/) / [ProseMirror](https://prosemirror.net/) —
  блочные редакторы в Docs и Markdown.
- [CodeMirror](https://codemirror.net/) (MIT) — редактор исходного кода в
  HTML.
- [Konva](https://konvajs.org/) — рендеринг на канвасе для Slides и диаграмм
  Sheets.
- [HarfBuzz](https://github.com/harfbuzz/harfbuzz) (wasm) — метрики
  шейпинга текста для сложных систем письма.
- [calamine](https://github.com/tafia/calamine) и
  [IronCalc](https://github.com/ironcalc/IronCalc) — слои чтения и расчёта в
  Rust-компоненте xlsx.
- [libeot](https://github.com/umanwizard/libeot) (MPL-2.0) — декодер
  MicroType Express для встроенных шрифтов PowerPoint, портированный на
  TypeScript.
- [React](https://react.dev/) (MIT) — слой пользовательского интерфейса
  каждого приложения.
- [Mermaid](https://mermaid.js.org/) (MIT) и [KaTeX](https://katex.org/)
  (MIT) — диаграммы и математические формулы в Markdown и Docs.
- [opentype.js](https://opentype.js.org/) (MIT) — разбор шрифтов для
  метрик и поиска глифов.
- [JSZip](https://stuk.github.io/jszip/) (MIT) и
  [fast-xml-parser](https://github.com/NaturalIntelligence/fast-xml-parser)
  (MIT) — слои контейнера OOXML и XML.
- [Fluent UI System Icons](https://github.com/microsoft/fluentui-system-icons)
  (MIT) — набор значков в лентах команд.
- [electron-updater](https://www.electron.build/) (MIT) — обновления
  внутри приложения.
- Шрифты Liberation, Carlito, Caladea и Noto CJK (OFL/Apache-2.0) —
  встроенные шрифты документов.

Команда `npm run notices` пересобирает сводку сторонних лицензий,
поставляемых с приложением (`tools/gen-third-party-notices.mjs`); все
рантайм-зависимости распространяются по лицензиям
MIT/Apache-2.0/BSD-3-Clause/OFL.

## Лицензия

GenOffice распространяется по лицензии [Apache License 2.0](../../LICENSE),
с одним исключением: каталог `ee/` зарезервирован для будущих корпоративных
модулей и покрывается [GenOffice Enterprise License](../../ee/LICENSE).

Названия и логотипы GenOffice и Genspark являются товарными знаками Mainfunc,
Inc. Лицензия Apache-2.0 не даёт права использовать их (см. раздел 6);
форки должны использовать собственный брендинг.
