<p align="center">
  <a href="https://genoffice.ai/">
    <picture>
      <source srcset="../assets/readme/hero-dark.webp" media="(prefers-color-scheme: dark)">
      <img src="../assets/readme/hero.webp" alt="GenOffice — Microsoft Office に代わるオープンソースの AI オフィススイート：Docs、Sheets、Slides、PDF、Markdown、HTML と内蔵 AI パネル" width="100%">
    </picture>
  </a>
</p>

<h1 align="center">GenOffice</h1>

<p align="center"><b>世界初のフル機能を備えたオープンソース AI オフィススイート。</b><br>
Word、Excel、PowerPoint、PDF ファイルを、あなたとあなたの AI が編集し、実際のファイル形式のまま保存します。</p>

<p align="center">
  <a href="../../LICENSE"><img src="https://img.shields.io/github/license/genspark-ai/genoffice" alt="License: Apache-2.0"></a>
  <a href="https://github.com/genspark-ai/genoffice/releases/latest"><img src="https://img.shields.io/github/v/release/genspark-ai/genoffice" alt="Latest release"></a>
  <a href="https://github.com/genspark-ai/genoffice/releases"><img src="https://img.shields.io/github/downloads/genspark-ai/genoffice/total" alt="Downloads"></a>
  <a href="https://github.com/genspark-ai/genoffice/stargazers"><img src="https://img.shields.io/github/stars/genspark-ai/genoffice?style=flat" alt="GitHub stars"></a>
</p>

<p align="center"><a href="../../README.md">English</a> · <a href="README.es.md">Español</a> · <a href="README.pt-BR.md">Português (Brasil)</a> · <a href="README.de.md">Deutsch</a> · <a href="README.fr.md">Français</a> · <a href="README.zh-CN.md">简体中文</a> · <a href="README.zh-TW.md">繁體中文</a> · <a href="README.ko.md">한국어</a> · <b>日本語</b> · <a href="README.ar.md">العربية</a> · <a href="README.ru.md">Русский</a> · <a href="README.it.md">Italiano</a> · <a href="README.nl.md">Nederlands</a> · <a href="README.pl.md">Polski</a> · <a href="README.cs.md">Čeština</a> · <a href="README.id.md">Bahasa Indonesia</a> · <a href="README.ms.md">Bahasa Melayu</a> · <a href="README.th.md">ไทย</a> · <a href="README.hi.md">हिन्दी</a> · <a href="README.he.md">עברית</a></p>

<p align="center">
  <a href="#download"><b>ダウンロード</b></a> ·
  <a href="https://genoffice.ai/"><b>ウェブサイト</b></a> ·
  <a href="https://genoffice.ai/join"><b>コミュニティ</b></a> ·
  <a href="../../PRIVACY.md"><b>プライバシー</b></a>
</p>

GenOffice は、macOS・Windows・Linux で動作する、Microsoft Office に代わる無料のオープンソース製品です。ネイティブの `.docx`・`.xlsx`・`.pptx` ファイルを開いて保存できるほか、PDF・Markdown・HTML の編集にも対応し、あらゆるドキュメントの隣に AI エージェントを配置します。取って付けたようなチャットボックスではなく、ファイルを読み込み、変更を加え、何をどう変更したかを正確に示してくれるエディタです。

- **実際のファイル形式、バイト単位で保持。** 編集した部分だけが書き換えられます。それ以外の部分はファイル内でバイト単位そのまま保持されるため、ドキュメントは Word・Excel・PowerPoint でも問題なく動作し続けます。
- **レビューできる AI。** 編集内容は変更履歴と差分として反映され、ワンクリックで元に戻せます。スプレッドシートには貼り付けの数値ではなく、実際に機能する数式が入ります。デッキやページはキャンバス上に生成され、そのまま自由に編集できます。
- **設計からローカル動作。** ファイルの開く・編集・保存・変換はすべてお使いのマシン上で行われます。PDF → Word / Excel / PowerPoint、Markdown → Word、HTML → Word の変換もすべてオンデバイスで実行されます。マシンの外に出るのは、選択した AI プロバイダーへの呼び出しだけです。
- **キーを使うか使わないかはあなた次第。** Genspark でサインインすればキーの用意は不要です。あるいは、Claude、OpenAI、Gemini、DeepSeek、Kimi、GLM、Qwen、Doubao、MiniMax、Grok、Mistral、OpenRouter、API Route、Requesty、または任意の OpenAI 互換エンドポイント（ローカルサーバーを含む）向けに、自分の API キーを持ち込むこともできます。

**入手方法：** [macOS](https://github.com/genspark-ai/genoffice/releases/latest)（Apple Silicon および Intel）·
[Windows](https://github.com/genspark-ai/genoffice/releases/latest)（x64 および Arm）·
[Linux](https://github.com/genspark-ai/genoffice/releases/latest)（deb、rpm、AppImage）—
詳細と要件は[ダウンロード](#download)を参照してください。

## デモ

6 つのアプリ、1 つの AI パネル。すべてのスクリーンショットは macOS 上の実際のアプリで撮影したもので、AI はパネル内に表示されているプロンプトそのもので動かしています。

### 1 · Docs — `.docx` を開いて、レビューできる AI で編集

<table>
<tr>
<td width="50%"><img src="../assets/readme/docs-report.webp" alt="GenOffice Docs で、全幅の表紙画像、網掛けの KPI 表、ヘッダーとフッターを備えた 2 段組みの年次報告書ページを 80% ズームで表示している様子（AI パネルは折りたたみ）"></td>
<td width="50%"><img src="../assets/readme/docs-ai.webp" alt="GenOffice Docs：バナー画像付きの会社概要ページ。AI が Overview を整えて新しい箇条書きセクションを挿入し、パネルにはワンクリックで元に戻せるボタンが表示されている"></td>
</tr>
<tr>
<td><b>Word と同じレイアウトでファイルを開く</b> — 2 段組セクション、フチなし画像、網掛けの表、ヘッダーとフッター、Word の行メトリクスに基づくページ送りまで再現します。スタイル、コメント、変更履歴、数式、手書き（インク）もそのまま保持されます。</td>
<td><b>編集を依頼する</b> — AI が必要なブロックを読み取り、Overview を書き直して新しい箇条書きセクションを挿入します。AI の各ターンはロールバック可能なスナップショットとして記録され、<b>変更履歴の記録</b>を有効にすると、編集は Word 形式のリビジョンとして反映されます。</td>
</tr>
</table>

### 2 · Sheets — 貼り付けの数値ではなく、実際に機能する数式とグラフを備えた `.xlsx`

<table>
<tr>
<td width="50%"><img src="../assets/readme/sheets-ai.webp" alt="GenOffice Sheets：AI が SUMIF 数式を使って地域・カテゴリ別売上をまとめた Summary シートと縦棒グラフを追加し、適用した 43 件の変更と元に戻すボタンを表示している"></td>
<td width="50%"><img src="../assets/readme/sheets-qa.webp" alt="GenOffice Sheets：「Q2 の売上で最も高かった地域は？」という質問に対し、AI が Europe と回答し、カテゴリ別の内訳と、参照したセルへのリンクを Orders シートの隣に表示している"></td>
</tr>
<tr>
<td><b>作ってもらう</b> — 一文の指示だけで、エージェントが地域・カテゴリ別の本物の <code>SUMIF</code> 数式を使った Summary シートを追加し、縦棒グラフを挿入し、43 件の変更をひとまとめの取り消し可能な操作として適用します。</td>
<td><b>質問する</b> — ワークブックについての質問には、根拠となる考え方と、実際に使用したセルへのクリック可能な引用付きで回答が返ってきます。内部では、自社開発の Rust 製 <code>.xlsx</code> エンジン、ピボットテーブル、スライサー、条件付き書式、数式トレースが動いています。</td>
</tr>
</table>

### 3 · Slides — 1 つのプロンプトから `.pptx` デッキへ

<img src="../assets/readme/slides-generate.webp" alt="GenOffice Slides が Aurora Home の投資家向けデッキを生成するタイムラプス：AI がパネルでストーリーラインを設計し、スライドがキャンバス上に次々と現れ、完成したデッキは締めのアスクで終わる" width="100%">

<table>
<tr>
<td width="50%"><img src="../assets/readme/slides-cover.webp" alt="GenOffice Slides：AI が生成した Aurora Home の投資家向けデッキの表紙スライドをキャンバス上に表示し、パネルには元の 1 行プロンプトと AI が作成した内容の要約が表示されている"></td>
<td width="50%"><img src="../assets/readme/slides-ai.webp" alt="GenOffice Slides：同じ 11 枚構成のデッキのデザインされたクロージングスライド。左側にサムネイル一覧、AI パネルにストーリーラインの要約が表示されている"></td>
</tr>
<tr>
<td><b>入力は 1 行だけ</b> — 「Aurora Home 向けの 10 枚構成の投資家向けピッチデッキを作って…」。GenOffice がストーリーラインを組み立て、数値をリサーチし、すべてのスライドを本物の <code>.pptx</code> としてキャンバスに描き出します。</td>
<td><b>完成したデッキが出てくる</b> — 統一されたタイポグラフィと画像、締めの Call to Action を備えた、デザイン済みの 11 枚のスライド。マスター、レイアウト、スマートガイド、非破壊トリミングでそのまま編集を続けることも、パネルにデザインの変更・文章の書き直し・順序の入れ替えを頼むこともできます。</td>
</tr>
</table>

### 4 · PDF — PDF のテキストをその場で編集し、オンデバイスで Word に変換

<table>
<tr>
<td width="50%"><img src="../assets/readme/pdf-edit.webp" alt="GenOffice PDF：テキスト編集モードがページ内のすべてのテキストブロックを枠線で示し、その場での編集を可能にする一方、AI パネルはレポートについての質問にページ引用付きで回答している"></td>
<td width="50%"><img src="../assets/readme/pdf-convert.webp" alt="GenOffice Docs が、Helios の四半期レビュー PDF からローカルで変換された Word 文書を、元の PDF の隣の 2 つ目のタブで開いている様子"></td>
</tr>
<tr>
<td><b>ページの中で直接編集</b> — テキスト編集モードは各テキストブロックを枠線で示し、その場での打ち直しができます。上から隠す注釈ではなく、PDFium を通じてコンテンツストリーム自体を元のフォントのまま書き換えます。長いレポートについて AI に尋ねれば、ページ引用付きで回答が返ってきます。</td>
<td><b>オンデバイスで変換</b> — <b>PDF コンバーター → PDF から Word</b> を使うと、見出し・統計行・段落をそのまま保った編集可能な <code>.docx</code> が生成され、元の PDF の隣の Docs で開きます。Excel や PowerPoint への変換も同様に動作し、スキャンされたページはシステムの OCR を通して処理されます。</td>
</tr>
</table>

### 5 · HTML — デザインブリーフから始める AI ページ／UI ビルダー

ページの目的と対象読者を伝えるだけです。AI はまず**デザインブリーフ**（フック、カラーパレット、タイポグラフィ、スタイルの方向性）を提案し、そのトークンに沿って単一の自己完結型 `.html` ファイルを組み立てます。

<img src="../assets/readme/html-restyle-motion.webp" alt="GenOffice HTML が Lumen のランディングページを再デザインするタイムラプス：パネルでの1回の Restyle リクエストにより、ダークな Midnight Studio ページが暖色系の Solar Daybreak バージョンへと変わる一方、すべてのセクションと文言はそのまま保たれる" width="100%">

<table>
<tr>
<td width="50%"><img src="../assets/readme/html-ai.webp" alt="GenOffice HTML：ソーラーデスクランプ向けに生成された、ダークな Midnight Studio 方向性のランディングページをライブプレビューで表示し、AI パネルには作成したページの要約が表示されている"></td>
<td width="50%"><img src="../assets/readme/html-restyle.webp" alt="同じ Lumen のランディングページを AI が暖かみのある Solar Daybreak 方向性にリスタイルした様子：紙のような背景、セリフ体の見出し、オレンジのアクセントに変わりつつ、すべてのセクションと文章はそのまま保持されている"></td>
</tr>
<tr>
<td><b>1 つのプロンプトから生成</b> — Lumen 向けに、大胆なヒーロー、機能カード、価格表、ウェイトリストフォームを Midnight Studio の方向性で構築しています。任意の要素をクリックしてリスタイルしたり、ダブルクリックでテキストを編集したり、CodeMirror のソースビューに切り替えたりできます。</td>
<td><b>同じデザインで方向性だけ変える</b> — <b>リスタイル</b>を 1 回リクエストするだけでブリーフのトークンが入れ替わり、ページ全体が追従します：暖かみのある紙の質感、エディトリアルなセリフ体、太陽のようなオレンジのアクセント。文章は書き直されません。フルスクリーンで発表したり、PDF や編集可能なネイティブ Word ドキュメントとして書き出したりもできます。</td>
</tr>
</table>
<table>
<tr>
<td width="50%"><img src="../assets/readme/html-dashboard.webp" alt="GenOffice HTML：フリーランスデザイナー向けに生成された、暖かみのあるリネン調の個人ダッシュボード UI。左側のレール、セリフ体の挨拶文、4 つの指標カードを備えている"></td>
<td width="50%"><img src="../assets/readme/html-report.webp" alt="GenOffice HTML：ブロードシート紙のスタイルで生成された EV 市場のデータレポート。セリフ体のマストヘッド、17.3 million という見出し数字、統計行を備えている"></td>
</tr>
<tr>
<td><b>UI モックアップ</b> — 「personal dashboard」のスターターは、ペルソナをそのまま動くレイアウトに変換します：左側のレール、挨拶文、課金可能時間のスパークライン、請求書と稼働率のカード。すべて開発者に渡せる本物の HTML です。</td>
<td><b>データストーリー</b> — 「data report」のスターターは、エディトリアルなブロードシートを組み立てます：セリフ体のマストヘッド、1 つの見出し数字、区切り線付きの統計行、インライン SVG チャート、手法についての注記まで揃っています。</td>
</tr>
</table>

### 6 · Markdown — プレーンな `.md` を対象にした、Ask AI 付きブロックエディタ

<table>
<tr>
<td width="50%"><img src="../assets/readme/markdown-ai.webp" alt="GenOffice Markdown：選択した段落に Ask AI のポップオーバーが表示され、入力した指示と、Polish・Make more concise・Expand・Fix grammar といった提案チップ、さらに「今すぐ送信」と「キューに追加」のボタンが並んでいる"></td>
<td width="50%"><img src="../assets/readme/markdown-render.webp" alt="GenOffice Markdown が、表・Mermaid フローチャート・タスクリストを含むローンチノート文書を表示し、左側には AI パネルのスタータープロンプトが並んでいる"></td>
</tr>
<tr>
<td><b>選択範囲について Ask AI</b> — どの文章を選択しても <b>Ask AI</b> チップが表示されます。指示を入力するか提案を選び、その場で送信するか、複数の編集をキューに入れて一括で実行することもできます。同じ操作はどのアプリでも使えます。</td>
<td><b>表示はリッチに、保存はプレーンな Markdown</b> — 見出し、リスト、表、画像、コードブロック、Mermaid 図を Tiptap ベースのブロックエディタで扱いながら、保存時はプレーンな <code>.md</code> として書き戻されます。<b>Markdown → Word</b> の書き出しも完全にローカルで行われます。</td>
</tr>
</table>

## GenOffice を選ぶ理由

- **オープンソース**。Apache-2.0 ライセンスで、GitHub 上で公開開発。
- **自分のマシンで動く。** macOS・Windows・Linux 向けのネイティブアプリ。ファイル
  はディスク上に残り、編集・保存・変換もすべて自分のマシンで行われる。
- **本物の Office ファイル。** ネイティブの `.docx`・`.xlsx`・`.pptx`、バイト単位
  で保持：触れていない部分はそのままコピーされる。
- **ドキュメント自体を編集する AI。** Docs では変更履歴、Sheets では実際に機能す
  る数式とグラフ、スライドはキャンバス上に描画され、AI の操作ごとに巻き戻せるス
  ナップショットが残る。
- **自分のモデル、自分のキー。** Genspark でログイン、または Claude・OpenAI・
  Gemini・DeepSeek などのキーを持ち込める。ローカルサーバーや任意の OpenAI 互換
  エンドポイントにも対応。
- **PDF もしっかり対応。** ページ内でテキストをその場で編集し、オンデバイスで
  PDF を Word・Excel・PowerPoint に変換。スキャンにはシステム OCR を使用。
- **Markdown と HTML にも対応**。同じ AI パネルを使い、ローカルで Word に書き出せ
  る。
- **無料**。個人でもチームでも利用できる。

## AI バックエンド

**Genspark でサインイン**すれば、設定は一切不要です。モデル呼び出しは Genspark のプロキシ
（Claude、GPT、Gemini 系列）を経由し、エージェントは Web・画像検索、画像生成、画像／音声／
動画の解析機能も利用できます。

**あるいは自分のキーを持ち込む。** 設定 → AI には Claude、OpenAI、Gemini、DeepSeek、Kimi、
GLM、Qwen、Doubao、MiniMax、Grok、Mistral、OpenRouter、API Route、Requesty、OpenCode Zen/Go が並び、さらに任意
の OpenAI 互換エンドポイント（ベース URL + キー、ローカルモデルサーバーを含む）を登録できる
カスタム枠も用意されています。検索とメディア関連は **AI メディア＆検索** の下で機能ごとに個
別のプロバイダーを設定できます：Web 検索には Serper か Tavily、画像生成と画像／動画解析には
OpenAI、Gemini、Doubao/Seedream、GLM、Grok、Qwen、MiniMax、または任意の OpenAI 互換の画像
エンドポイントが使えます。

スイート全体でライト・ダーク・システムのテーマに対応しています。テーマが変わるのは画面表示
だけで、書き出し・印刷・保存されたファイルは常にドキュメント本来の色を保持します。

<a id="download"></a>

## ダウンロード

| プラットフォーム                           | 要件                                                                   | ダウンロード                                                                                   |
| ------------------------------------------ | ---------------------------------------------------------------------- | ---------------------------------------------------------------------------------------------- |
| **macOS** — Apple Silicon（arm64）         | macOS 11 以降                                                          | [最新の `.dmg`（arm64）](https://github.com/genspark-ai/genoffice/releases/latest)             |
| **macOS** — Intel（x64）                   | macOS 11 以降                                                          | [最新の `.dmg`（x64）](https://github.com/genspark-ai/genoffice/releases/latest)               |
| **Windows**（x64、大半の PC）              | Windows 10 以降、Intel/AMD                                             | [最新の `-x64.exe` インストーラー](https://github.com/genspark-ai/genoffice/releases/latest)   |
| **Windows** on Arm（ARM64）                | Windows 11 on Arm（Snapdragon X など）                                 | [最新の `-arm64.exe` インストーラー](https://github.com/genspark-ai/genoffice/releases/latest) |
| **Linux** — Debian / Ubuntu                | x86_64、glibc 2.34 以降（Ubuntu 22.04 以降）                           | [最新の `.deb`](https://github.com/genspark-ai/genoffice/releases/latest)                      |
| **Linux** — Fedora / RHEL / openSUSE       | x86_64、glibc 2.34 以降（Fedora 35 以降、RHEL 9 以降、Leap 15.6 以降） | [最新の `.rpm`](https://github.com/genspark-ai/genoffice/releases/latest)                      |
| **Linux** — その他のディストリビューション | x86_64、glibc 2.34 以降、FUSE 2                                        | [最新の `.AppImage`](https://github.com/genspark-ai/genoffice/releases/latest)                 |

すべてのビルドは `main` ブランチから生成され、macOS・Windows のインストーラーには署名が付いて
います。以前のバージョンは [Releases](https://github.com/genspark-ai/genoffice/releases) ページにあります。

<details>
<summary><b>Linux へのインストール</b></summary>

deb は apt でインストールできます — 依存関係も自動的に取得され、GenOffice がアプリケーション
メニューに追加されます。

```bash
sudo apt install ./genoffice_<version>_amd64.deb
```

Fedora / RHEL 系 / openSUSE では、代わりに rpm をインストールしてください。

```bash
sudo dnf install ./genoffice-<version>.x86_64.rpm     # Fedora / RHEL family
sudo zypper install ./genoffice-<version>.x86_64.rpm  # openSUSE
```

AppImage はそのまま実行できます：FUSE 2 ランタイムをインストールし
（`sudo apt install libfuse2`。Ubuntu 24.04 ではパッケージ名が `libfuse2t64` です）、ファイル
に実行権限を付けてから実行してください。

```bash
chmod +x GenOffice-<version>.AppImage
./GenOffice-<version>.AppImage
```

</details>

## 内部の仕組み

Docs、Sheets、Slides、PDF、Markdown、HTML、そしてタブ管理を行うシェルという 7 つの Electron
アプリが、純粋な TypeScript パッケージによる 1 つのエンジン層と、`.xlsx` 用の Rust サイド
カーを共有しています。常に元ファイルが正となり、編集は最小限のパッチとして適用されるため、
エディタが触れなかった部分は往復の間もそのまま保持されます。

```
docx を開く ─► 元ファイルをハッシュ付きでアーカイブ（一切変更しない）
          ─► word/document.xml をブロックツリーへ解析、各ブロックは元の XML に紐づく
          ─► Tiptap エディタ（手動編集 + AI 編集、dirty 追跡）
保存      ─► dirty なブロック → OOXML フラグメント（既存スタイルのみ参照）
          ─► 元の document.xml へ差し込み、触れていないブロックはバイトのまま
          ─► zip を再パック；他のすべてのエントリはバイト単位でそのままコピー
```

パッケージ単位の詳しい解説（docx/pptx エンジン、`pdf2docx`、`html2docx`、エージェントコアと
プロバイダー）は [CONTRIBUTING.md](../../CONTRIBUTING.md#engine-packages) にあります。

## 開発

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

sheets アプリでは、xlsx サイドカー用に Rust ツールチェーン（PATH 上に `cargo`）も別途必要
です。`npm run build -w @genoffice/sheets` を実行すれば自動的にビルドされます。すべての変更
が満たすべきチェック内容や、プルリクエストの取り込み方については
[CONTRIBUTING.md](../../CONTRIBUTING.md) を参照してください。

## コミュニティ

GenOffice は活発に開発が続いており、みなさんのフィードバックがその方向性を形づくっています。

- **バグの報告や機能のリクエスト**は
  [GitHub Issues](https://github.com/genspark-ai/genoffice/issues) で受け付けています。
- **GenOffice のグループチャット**に
  [GenTeam](https://genoffice.ai/join) から参加すると、開発チームや他のユーザーと直接話せます。
- GenOffice が役に立ったら**リポジトリにスターを付けてください** — プロジェクトを支援する
  一番簡単な方法です。

## よくある質問

<details>
<summary><b>GenOffice は無料ですか？</b></summary>

はい。GenOffice は Apache-2.0 ライセンスのもとで無料かつオープンソースで提供されています —
試用期間もアプリ自体の有料プランもありません。

</details>

<details>
<summary><b>GenOffice で Microsoft の Word、Excel、PowerPoint ファイルを開けますか？</b></summary>

はい。GenOffice はネイティブの `.docx`、`.xlsx`、`.pptx` ファイルを開いて保存できます。保存
はバイト単位で保持されるため、触れなかった部分はそのまま書き戻され、ドキュメントは Microsoft
Office でも問題なく動作し続けます。

</details>

<details>
<summary><b>GenOffice はオフラインで使えますか？</b></summary>

ドキュメントの編集は完全にローカルで行われます — ファイルは開く・編集・保存・変換のどの場面
でもマシンの外に出ることはありません。AI 機能（エージェント、検索、画像ツール）を使うには
ネットワーク接続が必要で、Genspark でのサインインか、自分のモデル API キーのいずれかが必要
です。

</details>

<details>
<summary><b>GenOffice で PDF ファイルを編集できますか？</b></summary>

はい — 上から隠す注釈ではなく、ページのコンテンツストリーム自体を元のフォントを保ったまま
書き換える、本物の PDF テキスト・画像編集に対応しています。

</details>

<details>
<summary><b>GenOffice で PDF を Word、Excel、PowerPoint に変換できますか？</b></summary>

はい — すべてオンデバイスで行われます：PDFium による文字レベルの抽出と、形状に基づくレイア
ウト解析を使い、クラウドサービスやアップロードは一切発生しません。スキャンされたページにも
対応しており、macOS と Windows ではシステムの OCR が読み取るため、ページ画像ではなく編集可能
なテキストに変換されます。

</details>

<details>
<summary><b>自分の AI モデルや API キーを使えますか？</b></summary>

はい。キー不要の Genspark サインインに加えて、GenOffice は Claude、OpenAI、Gemini、
DeepSeek、Kimi、GLM、Qwen、Doubao、MiniMax、Grok、Mistral、OpenRouter、API Route、Requesty、OpenCode Zen/Go 向け
に自分のキーを持ち込むことができ、任意の OpenAI 互換エンドポイント（ローカルモデルサーバー
を含む）にも対応しています。検索、画像生成、画像／動画解析は、設定 → AI メディア＆検索でそ
れぞれ別のキーを設定します。

</details>

<details>
<summary><b>GenOffice で HTML を Word に変換できますか？</b></summary>

はい — HTML アプリの「Word として書き出す」を使うと、完全にオンデバイスでネイティブかつ編
集可能な `.docx` が生成されます。ページは内蔵の Chromium でレンダリングされたうえで、見出
し、段落、リスト、表、カード、KPI 行、フォームフィールド、ページ背景といった本物の Word 構
造に落とし込まれます。Word に対応する表現がないビジュアル（チャート、アイコン、装飾ボックス
など）だけが画像として埋め込まれます。

</details>

<details>
<summary><b>GenOffice はデータを収集しますか？</b></summary>

公式のパッケージ版ビルドは、デフォルトで限定的な利用状況の分析データを送信しますが、設定 →
一般からいつでも送信を無効にできます。分析データにドキュメントの内容、ファイル名、ファイル
パス、アカウント情報、メールアドレスが含まれることはありません。収集されるイベントとデータ
の詳細は [GenOffice のプライバシー](../../PRIVACY.md)を参照してください。

</details>

## セキュリティ

プロセスのセキュリティ方針（レンダラーのサンドボックス化、IPC バリデーション、外部リンクの
ゲーティング）と、AI が生成したコンテンツに対する脅威モデルについては
[SECURITY.md](../../SECURITY.md) を参照してください。

## クレジット

GenOffice は、以下のオープンソースプロジェクトなしには実現できませんでした。

- [Electron](https://www.electronjs.org/) — すべてのアプリのデスクトップランタイム。
- [Univer](https://github.com/dream-num/univer)（Apache-2.0）— Sheets が拡張しているスプレッ
  ドシート UI のコア。
- [PDFium](https://pdfium.googlesource.com/pdfium/)（BSD-3-Clause、
  [@embedpdf/pdfium](https://github.com/embedpdf/embed-pdf-viewer) 経由でバンドル）— 本物の
  PDF テキスト・画像編集を支えるコンテンツストリームエンジン。
- [pdf.js](https://github.com/mozilla/pdf.js)（Apache-2.0）と
  [pdf-lib](https://github.com/Hopding/pdf-lib)（MIT）— PDF のレンダリングとドキュメント組み
  立て。
- [Tiptap](https://tiptap.dev/) / [ProseMirror](https://prosemirror.net/) — Docs と Markdown
  のブロックエディタ。
- [CodeMirror](https://codemirror.net/)（MIT）— HTML アプリのソースエディタ。
- [Konva](https://konvajs.org/) — Slides と Sheets のグラフのキャンバスレンダリング。
- [HarfBuzz](https://github.com/harfbuzz/harfbuzz)（wasm）— 複雑な文字体系のテキストシェイピ
  ング計測。
- [calamine](https://github.com/tafia/calamine) と
  [IronCalc](https://github.com/ironcalc/IronCalc) — Rust 製 xlsx サイドカーの読み取り層と計
  算層。
- [libeot](https://github.com/umanwizard/libeot)（MPL-2.0）— 埋め込み
  PowerPoint フォント用の MicroType Express デコーダーを TypeScript に
  移植したもの。
- [React](https://react.dev/)（MIT）— すべてのアプリの UI レイヤー。
- [Mermaid](https://mermaid.js.org/)（MIT）と [KaTeX](https://katex.org/)
  （MIT）— Markdown と Docs における図と数式。
- [opentype.js](https://opentype.js.org/)（MIT）— メトリクス取得とグリフ
  検索のためのフォント解析。
- [JSZip](https://stuk.github.io/jszip/)（MIT）と
  [fast-xml-parser](https://github.com/NaturalIntelligence/fast-xml-parser)
  （MIT）— OOXML コンテナと XML レイヤー。
- [Fluent UI System Icons](https://github.com/microsoft/fluentui-system-icons)
  （MIT）— 各リボンで使われているアイコンセット。
- [electron-updater](https://www.electron.build/)（MIT）— アプリ内
  アップデート。
- Liberation、Carlito、Caladea、Noto CJK フォント（OFL/Apache-2.0）— バンドルされているドキュ
  メント用フォント。

`npm run notices` を実行すると、バンドルされているサードパーティライセンスの概要
（`tools/gen-third-party-notices.mjs`）が再生成されます。すべてのランタイム依存関係は
MIT/Apache-2.0/BSD-3-Clause/OFL のいずれかです。

## ライセンス

GenOffice は [Apache License 2.0](../../LICENSE) のもとでライセンスされていますが、1 つ例外
があります：`ee/` ディレクトリは将来のエンタープライズモジュール用に予約されており、
[GenOffice Enterprise License](../../ee/LICENSE) が適用されます。

GenOffice および Genspark の名称とロゴは Mainfunc, Inc. の商標です。Apache-2.0 ライセンスは
これらの使用を許諾するものではありません（セクション 6 を参照）。フォークする場合は、独自の
ブランディングを使用してください。
