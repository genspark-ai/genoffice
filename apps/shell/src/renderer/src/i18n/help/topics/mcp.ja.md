# コーディングエージェントの接続

GenOffice は Model Context Protocol を話せます。コーディングエージェントは、アプリが使うのと同じエンジンであなたの文書を読み、書き、レンダリングできます。エージェントがファイル形式を当て推量しているわけではありません。executor が検証するのと同じ定義から、型付きの op スキーマを受け取ります。

## 登録する

ふつうはコマンド 1 つで済みます。

```sh
genoffice mcp install all
```

このマシンにあるコーディングエージェント（Claude Code、Codex、Cursor、Gemini CLI、Copilot CLI、OpenCode、Windsurf）を探し、それぞれの設定に stdio サーバーのエントリを書き、そのファイルの残りは見つかったとき的样子のままにしておきます。

```sh
genoffice mcp list             # where each agent stands
genoffice mcp install cursor    # just one
genoffice mcp uninstall cursor  # take it back out
```

通常以外の場所にインストールしたエージェントには `--dir <path>` を指定します。`--force` を付けると、すでに存在するエントリを上書きします。

## 自分で動かす

別のマシンにあるクライアントには、代わりに HTTP で配信します。

```sh
genoffice mcp --http 8765 --token <secret>
```

`--host <addr>` で待ち受ける先が変わります。クライアントには同じトークンを渡してください。

HTTP の上でもファイルは運べます。`PUT /files/<name>` で 1 つをアップロードでき、各ツールはパスの代わりに `http(s)` の URL を受け取ります。出力はダウンロード URL として戻り、小さい場合は埋め込みリソースとしても返ります。op、spec、Markdown はいずれの場合もインラインで渡します。

## エージェントが手に入るもの

すべてのコマンドがツールになります。中でも覚えておきたいのは次のものです：

- **`docs`、`sheet`、`slides`** — アプリ自身の書き込み経路でファイルを読み書きします。1 回の **op** 単位で。新規のデッキは `deck_start`、`deck_page`、`deck_build` の順です。
- **`render`** — 1 ページ 1 枚の PNG をアプリのレンダラーが配置して出力するので、エージェントはスライドを推測せずに見た目で確認できます。
- **`pdf`** — アプリプロセスを起動せずに、PDF のテキスト層をページ単位で読み取ります。
- **`info`** — メタデータと構造の概要。見慣れないファイルに対して、たいてい最も安い最初の一手です。
- **`search`、`image`、`media`** — アプリで設定した提供元です。エージェントが自分の鍵を用意する必要はありません。
- **`merge`** — `{{key}}` テンプレートを埋めます。

## スキーマと小さめの予算

`apply` と `create` は `ops`、`cells`、`data` パラメーターを、op ごとの型付きスキーマ付きで公開します。そのスキーマは `genoffice guide <domain> --json` から生成されます。正確ですが、大きなものです。コンテキストウィンドウの狭いクライアントは、代わりに単純な配列を要求できます：

```sh
genoffice mcp --compact-schemas
GENOFFICE_MCP_COMPACT_SCHEMAS=1 genoffice mcp
```

## スキル

op の語彙を知らないエージェントは当て推量します。`genoffice skill` は、検出したエージェントに GenOffice のスキルをインストールします。持ち込むのはリファレンスとデザインガイド — `genoffice guide` が出すのと同じ内容です。

## これが違うもの

MCP サーバーはファイルを読み書きします。ウィンドウではありません。AI パネルはなく、アプリ内の更新ダイアログも適用外です。ウィンドウが必要な手順は、ファイルを開いてください。
