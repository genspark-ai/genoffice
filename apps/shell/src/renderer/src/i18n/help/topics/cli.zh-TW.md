# 命令列與代理

每次安裝都會附帶一個 `genoffice` 命令，它驅動的是視窗所用的同一套引擎——同樣的解析器、同樣的寫入器、同樣的算繪器。應用程式儲存的檔案與命令寫出的檔案是同一個檔案；應用程式 AI 面板能通過的檢查，命令也能通過。

```sh
genoffice --help          # every command
genoffice guide slides    # the op reference, for writing your own
```

整個介面就在這一屏裡——每個命令後面都有一行說明它是做什麼的，然後是全域選項和結束碼：

![genoffice --help 的真實輸出：版本橫幅、每個命令帶一行說明，以及全域選項和結束碼](img/cli.png)

## 取得命令

macOS 與 Windows 上它位於應用程式套件內。想直接以名稱呼叫，先執行一次 `genoffice install-cli`：它會把套件內的二進位檔建立符號連結到 `/usr/local/bin`，Windows 上則連結到你的使用者 `PATH`。

## 值得記住的命令

| 命令              | 用途                                                                                                                                       |
| ----------------- | ------------------------------------------------------------------------------------------------------------------------------------------ |
| `open`            | 在應用程式裡開啟文件；應用程式沒有執行時會先啟動。                                                                                         |
| `convert`         | 用應用程式自己的引擎在格式之間轉換。                                                                                                       |
| `create`          | 從結構化內容建立文件。                                                                                                                     |
| `render`          | 每頁一張 PNG，依算繪器排出的版面輸出。                                                                                                     |
| `pdf`             | 逐頁讀取 PDF 的文字層，不需要啟動應用程式行程。                                                                                            |
| `info`            | 文件的中繼資料與結構摘要。                                                                                                                 |
| `search`          | 透過應用程式裡設定的供應商進行網頁或圖片搜尋。                                                                                             |
| `image` / `media` | 產生圖片，或描述圖片、影片、音訊檔案並向它們提問。                                                                                         |
| `merge`           | 填入 `.docx`、`.pptx` 或 `.xlsx` 範本中的 `{{key}}` 預留位置。                                                                             |
| `capabilities`    | 回報本機設定了哪些雲端功能。                                                                                                               |
| `guide`           | 操作參考與設計指南，由執行器驗證所依據的同一份定義產生，因此不會與 `apply` 接受的內容產生落差。`--json` 會連同每個操作的 schema 一起回傳。 |
| `install-cli`     | 把 `genoffice` 放進 `PATH`。                                                                                                               |
| `skill`           | 列出本機找到的程式開發代理，並在其中安裝或更新 GenOffice 技能。                                                                            |
| `mcp`             | 把每個命令當成 Model Context Protocol 工具提供。見**連接程式開發代理**。                                                                   |

## 編輯：docs、sheets、slides

`genoffice docs`、`genoffice sheet` 與 `genoffice slides` 都要先接一個子命令——`read`、`apply`、`check`，slides 另有 `audit`、`render`、`replace`——再透過應用程式所用的同一條寫入路徑讀寫。它們共用一套詞彙：一個 **op** 就是一次編輯，一個 **spec** 就是依序套用的一組 op。

```sh
genoffice sheet read report.xlsx
genoffice slides apply deck.pptx --ops ops.json --dry-run
genoffice docs apply brief.docx --ops ops.json
```

`--dry-run` 只會回報一批操作會做什麼，不寫入任何內容，這是讓一份 spec 落地前先自查最省事的辦法。應用程式裡的 AI 面板跑的就是這些 op，所以凡是你能請它做的事，都能寫成指令碼。

## Model Context Protocol

`genoffice mcp` 把每個命令當成 MCP 工具提供；`genoffice mcp install <agent|all>` 則會把它註冊到程式開發代理自己的設定裡。這一側見**連接程式開發代理**。

## 命令不會做的事

它只負責讀寫檔案。它不是應用程式：沒有視窗，應用程式內的更新對話方塊也不適用。凡是需要視窗的事——AI 面板、對算繪後投影片的 QC 檢查——都得等你開啟檔案。`genoffice render` 可以在沒有視窗的情況下取得像素，`genoffice slides` 也能自行稽核一份簡報的版面。
