# コマンドラインとエージェント

どのインストールにも `genoffice` コマンドが含まれ、ウィンドウと同じエンジンを動かします。同じパーサー、同じ書き出し、同じレンダラーです。アプリが保存したファイルとコマンドが書き込むファイルはまったく同じものであり、アプリの AI パネルが通る検査にはコマンドも通ります。

```sh
genoffice --help          # every command
genoffice guide slides    # the op reference, for writing your own
```

## コマンドを用意する

macOS と Windows ではアプリバンドルの中にあります。名前で呼び出すには `genoffice install-cli` を一度実行してください。同梱のバイナリを `/usr/local/bin` に、Windows ではユーザーの `PATH` にシンボリックリンクします。

## 覚えておきたいコマンド

| コマンド | できること |
| --- | --- |
| `open` | アプリ内で文書を開きます。アプリが動いていなければ起動します。 |
| `convert` | アプリ本来のエンジンで形式を変換します。 |
| `create` | 構造化された内容から文書を作ります。 |
| `render` | レンダラーが配置したそのままを、1 ページ 1 枚の PNG として出力します。 |
| `pdf` | アプリプロセスを起動せずに、PDF のテキスト層をページ単位で読み取ります。 |
| `info` | 文書のメタデータと構造の概要。 |
| `search` | アプリで設定した提供元を通じて、ウェブ検索や画像検索を行います。 |
| `image` / `media` | 画像を生成します。画像・動画・音声ファイルを説明させ、質問することもできます。 |
| `merge` | `.docx`、`.pptx`、`.xlsx` のテンプレートの `{{key}}` プレースホルダーを埋めます。 |
| `capabilities` | このマシンでクラウド機能がどう設定されているかを報告します。 |
| `guide` | op のリファレンスとデザインガイドです。executor が検証に使うのと同じ定義から生成されるので、`apply` が受け付ける内容からずれることはありません。`--json` を付けると各 op のスキーマ付きで返します。 |
| `install-cli` | `genoffice` を `PATH` に追加します。 |
| `skill` | このマシンで見つかったコーディングエージェントを一覧し、GenOffice のスキルをインストールまたは更新します。 |
| `mcp` | すべてのコマンドを Model Context Protocol のツールとして公開します。**コーディングエージェントの接続** を参照してください。 |

## 編集：ドキュメント、スプレッドシート、スライド

`genoffice docs`、`genoffice sheet`、`genoffice slides` は、アプリと同じ書き込み経路で読み書きします。語彙は共通です。**op** は 1 回の編集、**spec** は op を順番に適用したリストです。

```sh
genoffice sheet read report.xlsx
genoffice slides apply deck.pptx --ops ops.json --dry-run
genoffice docs apply brief.docx --ops ops.json
```

`--dry-run` は、そのバッチが何をするかを報告するだけで、何も書き込みません。spec を適用する前に手早く確認するための、安い方法です。アプリ内の AI パネルが動くのもこの op なので、AI に依頼できることはすべてスクリプトにできます。

## Model Context Protocol

`genoffice mcp` はすべてのコマンドを MCP のツールとして公開し、`genoffice mcp install <agent|all>` はコーディングエージェント自身の設定に登録します。そちらは **コーディングエージェントの接続** を参照してください。

## コマンドがしないこと

ファイルを読んで書き込むことはできます。アプリそのものではありません。ウィンドウはなく、アプリ内の更新ダイアログも適用外です。ウィンドウが必要な処理 — AI パネル、レンダリング済みのスライドに対する QC パス — は、ファイルを開くまで待ってください。`genoffice render` ならウィンドウなしでピクセルが得られ、`genoffice slides` はデッキのレイアウトを自分で検査します。
