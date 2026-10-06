# 連接程式開發代理

GenOffice 支援 Model Context Protocol，因此程式開發代理可以透過應用程式所用的同一套引擎讀寫並算繪你的文件。代理並不是靠猜檔案格式：它取得的型別化 op schema，與執行器驗證所依據的是同一份定義。

## 在應用程式裡註冊

**設定 ▸ 整合**就是在這裡做。這個面板分成兩半，你可以只用其中一半，或兩半一起用。

**技能。**本機上找到的每個程式開發代理各佔一列——Claude Code、Codex、Cursor、Gemini CLI、Copilot CLI、OpenCode、Windsurf——每一列都有**安裝**、**更新**與**解除安裝**，另外還有**安裝到其他資料夾…**、**下載 skill（zip）**和**複製路徑**。如果你的助理不在列表上，就讓 GenOffice 指向它讀取 `SKILL.md` 的那個資料夾，或者把 zip 存下來，交給助理自己去裝。技能與 MCP 可以並存：助理挑一個用，而它們做起來的事情完全一樣。

**MCP。**有兩條路線：**由助理啟動（建議）**，你把顯示出來的設定加進用戶端，助理就會自己把伺服器拉起來；以及**本機 HTTP 服務**，由應用程式替你執行。不論走哪一條，助理最後都是透過 GenOffice 做事，而你不必輸入任何命令。

## 從命令列註冊

同樣的事，在終端機裡也做得到——這是進階路線，當代理裝在面板找不到的地方時，就該走這條：

```sh
genoffice mcp install all
```

它會找出本機上的程式開發代理，把 stdio 伺服器項目寫進各自的設定檔，其餘部分維持原樣。

```sh
genoffice mcp list             # where each agent stands
genoffice mcp install cursor    # just one
genoffice mcp uninstall cursor  # take it back out
```

裝在不常見位置的代理要加上 `--dir <path>`；`--force` 會改寫已經存在的項目。

伺服器能接受的東西都列在一個畫面上——註冊、移除與列出這三種形式，改用 HTTP 提供服務，以及兩個 schema 開關：

![genoffice mcp --help 的真實輸出：install、uninstall 與 list 三種形式，以及 --http、--host、--token、--compact-schemas、--dir 和 --force 選項](img/mcp.png)

## 沒有助理也能執行

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

## 為什麼需要技能

不熟悉 op 詞彙的代理只能靠猜。技能裡附著參考手冊與設計指南——和 `genoffice guide` 印出來的是同一份材料——所以助理寫下的 op，它的規格是真的讀過的。可以從上面的面板安裝，也可以用終端機裡的 `genoffice skill`。

## 它在應用程式裡能碰到什麼

伺服器不只限於磁碟上的檔案。GenOffice 執行期間，代理還可以透過視窗來做事：

- **`open_in_genoffice`** 在一個分頁裡開啟檔案，並把它帶到前景。
- **`open_documents`** 列出你已經開啟的每份文件——id、類型、路徑，以及有沒有尚未儲存的變更——然後讀取某份文件的目前內容或把它關掉，預設會先儲存，除非你讓它丟棄。
- **內容類工具**把那個 id（或路徑）當作 `document` 參數，因此修改會落在你已經開著的那個分頁裡，視窗也會切過去顯示它。

有兩件事仍然碰不到：沒有 AI 面板，應用程式內的更新對話方塊也不適用。

## 本機 HTTP 服務面板

**本機 HTTP 服務**下，應用程式可以自己執行伺服器，而不是把它丟給助理：一個啟用開關、一個**連接埠**欄位、一個**執行中／未執行**的狀態指示，以及**後台生成**（直接把文件寫到某個路徑，不開啟介面），再加上一份可以複製的**用戶端設定範例**。展開**進階**後，會多出兩個連線網址——Streamable HTTP 和舊的 SSE——一個**健康檢查**網址，以及一個**日誌**開關，把伺服器與工具的活動記錄到本機檔案，可以在那裡**開啟**、**重新整理**或**清除**。它只監聽 localhost。
