# 連接程式開發代理

GenOffice 支援 Model Context Protocol，因此程式開發代理可以透過應用程式所用的同一套引擎讀寫並算繪你的文件。代理並不是靠猜檔案格式：它取得的型別化 op schema，與執行器驗證所依據的是同一份定義。

## 註冊

一般情況下一道命令就夠：

```sh
genoffice mcp install all
```

它會找出本機上的程式開發代理——Claude Code、Codex、Cursor、Gemini CLI、Copilot CLI、OpenCode、Windsurf——把 stdio 伺服器項目寫進各自的設定檔，其餘部分維持原樣。

```sh
genoffice mcp list             # where each agent stands
genoffice mcp install cursor    # just one
genoffice mcp uninstall cursor  # take it back out
```

如果某個代理裝在不常見的位置，加上 `--dir <path>`；`--force` 會改寫已經存在的項目。

## 自行執行

如果用戶端在另一台電腦上，就改用 HTTP 提供服務：

```sh
genoffice mcp --http 8765 --token <secret>
```

`--host <addr>` 決定它監聽的位址。把同一個權杖傳給用戶端。

走 HTTP 時檔案也能傳：`PUT /files/<name>` 上傳一個檔案，每個工具都接受 `http(s)` URL 來取代路徑，輸出則以下載 URL 回傳；檔案夠小時也可以當成內嵌資源。op、spec 與 Markdown 在兩種方式下都是內嵌傳遞。

## 代理能取得什麼

每個命令都是一個工具。其中值得一提的有：

- **`docs`、`sheet`、`slides`** — 透過應用程式自己的寫入路徑讀寫檔案，每次一個 **op**。新建一份簡報會依序走 `deck_start`、`deck_page`、`deck_build`。
- **`render`** — 每頁一張 PNG，由應用程式的算繪器排版，這樣代理可以直接看投影片，而不是靠猜。
- **`pdf`** — 逐頁讀取 PDF 的文字層，不需要應用程式行程。
- **`info`** — 中繼資料與結構摘要，通常是拿到陌生檔案後最省事的第一通呼叫。
- **`search`、`image`、`media`** — 應用程式裡已設定的供應商，代理不必再自備金鑰。
- **`merge`** — 填入 `{{key}}` 範本。

## Schema 與更小的預算

`apply` 與 `create` 會用逐 op 的型別化 schema 宣告它們的 `ops`、`cells` 與 `data` 參數，schema 由 `genoffice guide <domain> --json` 產生。這樣很精確，但體積不小。上下文視窗吃緊的用戶端可以改用單純的陣列：

```sh
genoffice mcp --compact-schemas
GENOFFICE_MCP_COMPACT_SCHEMAS=1 genoffice mcp
```

## 技能

不熟悉 op 詞彙的代理只能靠猜。`genoffice skill` 會把它找到的代理都裝上一份 GenOffice 技能，內含參考手冊與設計指南——和 `genoffice guide` 印出來的是同一份材料。

## 它不是什麼

MCP 伺服器負責讀寫檔案。它不是視窗：沒有 AI 面板，應用程式內的更新對話方塊也不適用。哪一步需要視窗，就把檔案開啟。
