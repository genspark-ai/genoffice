<p align="center">
  <a href="https://genoffice.ai/">
    <picture>
      <source srcset="../assets/readme/hero-dark.webp" media="(prefers-color-scheme: dark)">
      <img src="../assets/readme/hero.webp" alt="GenOffice — 開源 AI 辦公套件：Docs、Sheets、Slides、PDF、Markdown 與 HTML，內建 AI 面板" width="100%">
    </picture>
  </a>
</p>

<h1 align="center">GenOffice</h1>

<p align="center"><b>全球首款功能完整的開源 AI 辦公套件。</b><br>
Word、Excel、PowerPoint 與 PDF 檔案，由你與你的 AI 共同編輯，並以原生格式存回。</p>

<p align="center">
  <a href="../../LICENSE"><img src="https://img.shields.io/github/license/genspark-ai/genoffice" alt="License: Apache-2.0"></a>
  <a href="https://github.com/genspark-ai/genoffice/releases/latest"><img src="https://img.shields.io/github/v/release/genspark-ai/genoffice" alt="Latest release"></a>
  <a href="https://github.com/genspark-ai/genoffice/releases"><img src="https://img.shields.io/github/downloads/genspark-ai/genoffice/total" alt="Downloads"></a>
  <a href="https://github.com/genspark-ai/genoffice/stargazers"><img src="https://img.shields.io/github/stars/genspark-ai/genoffice?style=flat" alt="GitHub stars"></a>
</p>

<p align="center"><a href="../../README.md">English</a> · <a href="README.es.md">Español</a> · <a href="README.pt-BR.md">Português (Brasil)</a> · <a href="README.de.md">Deutsch</a> · <a href="README.fr.md">Français</a> · <a href="README.zh-CN.md">简体中文</a> · <b>繁體中文</b> · <a href="README.ko.md">한국어</a> · <a href="README.ja.md">日本語</a> · <a href="README.ar.md">العربية</a> · <a href="README.ru.md">Русский</a> · <a href="README.it.md">Italiano</a> · <a href="README.nl.md">Nederlands</a> · <a href="README.pl.md">Polski</a> · <a href="README.cs.md">Čeština</a> · <a href="README.id.md">Bahasa Indonesia</a> · <a href="README.ms.md">Bahasa Melayu</a> · <a href="README.th.md">ไทย</a> · <a href="README.hi.md">हिन्दी</a> · <a href="README.he.md">עברית</a></p>

<p align="center">
  <a href="#download"><b>下載</b></a> ·
  <a href="https://genoffice.ai/"><b>官網</b></a> ·
  <a href="https://genoffice.ai/join"><b>社群</b></a> ·
  <a href="../../PRIVACY.md"><b>隱私權</b></a>
</p>

GenOffice 是一套免費、開源的 Microsoft Office 替代方案，支援 macOS、
Windows 與 Linux。它能開啟並儲存原生的 `.docx`、`.xlsx` 與 `.pptx`
檔案，也能編輯 PDF、Markdown 與 HTML，並在每份文件旁放上一個 AI
代理——不是硬塞在旁邊的聊天視窗，而是真正讀懂檔案、動手修改，
並清楚告訴你它改了哪裡的編輯器。

- **原生格式，位元級保留。** 只有你編輯的部分會被重寫，檔案裡其他
  內容都會原樣保留，文件在 Word、Excel 與 PowerPoint 中依然能正常運作。
- **AI 的每一步都能檢視。** 編輯以追蹤修訂與差異呈現，一鍵即可回復。
  試算表拿到的是即時公式，不是貼上的數字。投影片與頁面會直接生成在
  畫面上，且完全可再編輯。
- **本機優先的設計。** 檔案的開啟、編輯、儲存與轉換都在你的裝置上完成。
  PDF → Word / Excel / PowerPoint、Markdown → Word、HTML → Word 全部
  在本機執行。只有 AI 呼叫會離開這台裝置，且僅送往你選擇的服務商。
- **用你自己的金鑰，或完全不用。** 用 Genspark 登入即可省去金鑰設定，
  或自行帶入 Claude、OpenAI、Gemini、DeepSeek、Kimi、GLM、Qwen、
  Doubao、MiniMax、Grok、Mistral、OpenRouter，或任何相容 OpenAI 介面
  的服務端點，也支援本機伺服器。

**立即取得：** [macOS](https://github.com/genspark-ai/genoffice/releases/latest)（Apple Silicon 與 Intel）·
[Windows](https://github.com/genspark-ai/genoffice/releases/latest)（x64 與 Arm）·
[Linux](https://github.com/genspark-ai/genoffice/releases/latest)（deb、rpm、AppImage）——
詳細內容與系統需求請見[下載](#download)。

## Demo

六款應用程式，一個 AI 面板。每張截圖都是在 macOS 上實際操作的真實畫面，
AI 的動作皆由面板中可見的提示詞驅動。

### 1 · Docs — 開啟並編輯 `.docx`，每一步 AI 編輯都能檢視

<table>
<tr>
<td width="50%"><img src="../assets/readme/docs-report.webp" alt="GenOffice Docs 呈現一份雙欄年度報告頁面，含滿版封面圖片、灰底 KPI 表格、頁首與頁尾，縮放 80%，AI 面板收合"></td>
<td width="50%"><img src="../assets/readme/docs-ai.webp" alt="GenOffice Docs：一份含橫幅圖片的公司概覽；AI 精簡了概覽段落並插入新的條列區塊，面板提供一鍵回復"></td>
</tr>
<tr>
<td><b>依 Word 排版方式開啟檔案</b>——雙欄版面、滿版出血圖片、灰底表格、
頁首頁尾，分頁依 Word 的行距量測。樣式、註解、追蹤修訂、公式與手寫
筆跡都能原樣往返不受影響。</td>
<td><b>直接說出想要的修改</b>——AI 只讀取需要的區塊，重寫概覽段落並插入
新的條列區塊。每一次 AI 操作都是可回復的快照；開啟<b>追蹤修訂</b>後，
編輯會以 Word 樣式修訂呈現。</td>
</tr>
</table>

### 2 · Sheets — `.xlsx` 有即時公式與圖表，不是貼上的數字

<table>
<tr>
<td width="50%"><img src="../assets/readme/sheets-ai.webp" alt="GenOffice Sheets：AI 新增了一張摘要工作表，用 SUMIF 公式按地區與類別統計營收，並加上直條圖，回報套用了 43 項變更並提供復原按鈕"></td>
<td width="50%"><img src="../assets/readme/sheets-qa.webp" alt="GenOffice Sheets：詢問哪個地區在第二季營收領先，AI 回答歐洲並附上分類明細，同時將使用到的儲存格以連結形式引用，旁邊是訂單工作表"></td>
</tr>
<tr>
<td><b>直接建好</b>——只憑一句話，代理就能新增一張摘要工作表，用真正的
<code>SUMIF</code> 公式依地區與類別統計，插入直條圖，並把 43 項變更
整合成一次可復原的批次操作。</td>
<td><b>直接詢問</b>——對工作簿的提問會附上推理過程，並把實際使用的
儲存格做成可點擊的引用連結。背後是自研的 Rust `.xlsx` 引擎、樞紐分析表、
篩選器、條件格式與公式追蹤。</td>
</tr>
</table>

### 3 · Slides — 從一句提示詞到一份 `.pptx` 投影片

<img src="../assets/readme/slides-generate.webp" alt="GenOffice Slides 產生 Aurora Home 投資人簡報的縮時過程：AI 在面板中規劃故事線，投影片依序出現在畫布上，最終簡報以收尾的訴求結束" width="100%">

<table>
<tr>
<td width="50%"><img src="../assets/readme/slides-cover.webp" alt="GenOffice Slides：畫面上是 AI 生成的 Aurora Home 投資人簡報封面投影片，附原始的一行提示詞與 AI 在面板中對已完成內容的摘要"></td>
<td width="50%"><img src="../assets/readme/slides-ai.webp" alt="GenOffice Slides：同一份 11 頁投影片的收尾頁設計完成，左側是縮圖列，AI 面板總結了整份簡報的敘事脈絡"></td>
</tr>
<tr>
<td><b>只需一句話</b>——「為 Aurora Home 製作一份 10 頁的投資人簡報……」。
GenOffice 會規劃敘事脈絡、研究相關數據，並把每一頁直接生成在畫面上，
成為一份真正的 `.pptx`。</td>
<td><b>拿到的是完整成品</b>——十一頁設計完整的投影片，字體排版、圖像
風格一致，並以行動呼籲收尾；可繼續用母片、版面配置、智慧輔助線與
不破壞原圖的裁切功能編輯，或請面板重新設計風格、改寫文字、調整順序。</td>
</tr>
</table>

### 4 · PDF — 就地編輯 PDF 文字，在本機把 PDF 轉成 Word

<table>
<tr>
<td width="50%"><img src="../assets/readme/pdf-edit.webp" alt="GenOffice PDF：編輯文字模式為頁面上每個文字區塊畫出框線以供就地編輯，同時 AI 面板正在回答一則關於報告的問題並附上頁碼引用"></td>
<td width="50%"><img src="../assets/readme/pdf-convert.webp" alt="GenOffice Docs 顯示由 Helios 季度報告 PDF 在本機轉換而成的 Word 文件，於第二個分頁開啟，旁邊是原始 PDF"></td>
</tr>
<tr>
<td><b>直接在頁面內編輯</b>——編輯文字模式為每個文字區塊畫出框線以供
重新輸入；內容串流是透過 PDFium 用原始字體重寫，而不是覆蓋式的注解。
向 AI 詢問長篇報告的問題，也能得到附頁碼引用的回答。</td>
<td><b>在本機完成轉換</b>——<b>PDF 轉換器 → PDF 轉 Word</b> 會產生可
編輯的 `.docx`，在 Docs 中與來源檔並排開啟，標題、統計列與段落皆完整
保留。Excel 與 PowerPoint 的轉換方式相同；掃描頁面則交由系統 OCR 處理。</td>
</tr>
</table>

### 5 · HTML — 先出設計簡報，再由 AI 建立網頁與介面

說明這個網頁是做什麼用的、給誰看。AI 會先提出一份**設計簡報**——
主打亮點、色盤、字體與風格方向——再依這些設計標記建出一個完整
獨立的 `.html` 檔案。

<img src="../assets/readme/html-restyle-motion.webp" alt="GenOffice HTML 重新設計 Lumen 到達頁面的縮時過程：面板中一次 Restyle 請求就把深色的 Midnight Studio 頁面變成暖色的 Solar Daybreak 版本，而每個區塊與所有文案都維持不變" width="100%">

<table>
<tr>
<td width="50%"><img src="../assets/readme/html-ai.webp" alt="GenOffice HTML：以深色 Midnight Studio 風格方向生成的太陽能桌燈登陸頁，畫面顯示即時預覽，AI 面板總結了剛完成的頁面"></td>
<td width="50%"><img src="../assets/readme/html-restyle.webp" alt="同一份 Lumen 登陸頁由 AI 重新設計成溫暖的 Solar Daybreak 風格方向：紙質背景、襯線標題與橘色重點色，所有版塊與文案皆保留"></td>
</tr>
<tr>
<td><b>一句提示詞就能生成</b>——為 Lumen 打造的醒目主視覺、功能卡片、
價格方案與候補名單表單，皆以 Midnight Studio 風格方向建成。點選任何
元素即可重新設計樣式，雙擊即可編輯文字，也能切換到 CodeMirror
原始碼視圖。</td>
<td><b>同一份設計，換個方向</b>——一次<b>重新設計</b>請求就會換掉設計簡報
的標記，頁面隨之調整：溫暖的紙感、雜誌風襯線字體、日落橘重點色，
文案完全不重寫。可全螢幕展示，或匯出為 PDF 或原生可編輯的 Word 文件。</td>
</tr>
</table>
<table>
<tr>
<td width="50%"><img src="../assets/readme/html-dashboard.webp" alt="GenOffice HTML：為自由接案設計師生成的個人儀表板介面，採溫暖亞麻風格，含左側導覽列、襯線問候語與四張指標卡片"></td>
<td width="50%"><img src="../assets/readme/html-report.webp" alt="GenOffice HTML：以報紙版式生成的電動車市場數據報告，含襯線報頭、一項 1730 萬的頭條數字與一列統計數據"></td>
</tr>
<tr>
<td><b>介面原型</b>——「個人儀表板」範本能把一個人物設定變成可用的
版面：左側導覽列、問候語、計費工時走勢圖、發票與使用率卡片，全都是
可以直接交給工程師的真實 HTML。</td>
<td><b>數據敘事</b>——「數據報告」範本會建出一份雜誌式報紙版面：襯線
報頭、一個頭條數字、以分隔線隔開的統計列、內嵌 SVG 圖表，以及一段
方法論說明。</td>
</tr>
</table>

### 6 · Markdown — 以區塊編輯器編輯純 `.md`，並內建 Ask AI

<table>
<tr>
<td width="50%"><img src="../assets/readme/markdown-ai.webp" alt="GenOffice Markdown：選取一段文字後出現 Ask AI 彈出視窗，內含輸入的指令與建議選項，如「潤飾」「更精簡」「擴寫」「修正語法」，並附「立即送出」與「加入佇列」按鈕"></td>
<td width="50%"><img src="../assets/readme/markdown-render.webp" alt="GenOffice Markdown 呈現一份上線公告文件，內含表格、Mermaid 流程圖與待辦清單，左側是 AI 面板的起手提示詞"></td>
</tr>
<tr>
<td><b>針對選取內容詢問 AI</b>——選取任意段落即會出現<b>Ask AI</b>選項：
輸入指令或挑選建議，可立即送出，也能把多個定位好的編輯加入佇列一次
執行。這個入口在每個應用程式中都相同。</td>
<td><b>渲染顯示，仍以純 Markdown 儲存</b>——標題、清單、表格、圖片、
程式碼區塊與 Mermaid 圖表都在 Tiptap 區塊編輯器中呈現，存回時仍是
純 `.md`，並支援完全在本機執行的 <b>Markdown → Word</b> 匯出。</td>
</tr>
</table>

## 為什麼選 GenOffice

- **開源**，採用 Apache-2.0 授權，在 GitHub 上公開開發。
- **由你掌控。** macOS、Windows 與 Linux 原生應用程式；檔案留在本機硬碟，每一次
  編輯、儲存與轉換都在你的裝置上完成。
- **真正的 Office 檔案。** 原生 `.docx`、`.xlsx` 與 `.pptx`，位元級保留：未變更
  的部分會原樣複製。
- **AI 直接編輯文件本身。** Docs 中的追蹤修訂、Sheets 中即時運算的公式與圖表、
  直接在畫布上生成的投影片，每一次 AI 操作都會留下可回復的快照。
- **自帶模型，自帶金鑰。** 使用 Genspark 登入，或帶上 Claude、OpenAI、Gemini、
  DeepSeek 等服務的金鑰，同時支援本機伺服器與任何相容 OpenAI 的服務端點。
- **認真做好 PDF。** 直接在頁面內編輯文字，並在本機將 PDF 轉換為 Word、Excel 或
  PowerPoint，掃描檔案支援系統 OCR。
- **同樣支援 Markdown 與 HTML**，共用同一套 AI 面板，並可在本機匯出為 Word。
- **免費**，個人與團隊皆可使用。

## AI 後端

**用 Genspark 登入**就完全不用設定：模型呼叫會透過 Genspark 代理路由
（涵蓋 Claude、GPT 與 Gemini 家族），代理也能使用網頁與圖片搜尋、
圖片生成，以及圖片／音訊／影片分析。

**或自帶金鑰。**「設定 → AI」列出了 Claude、OpenAI、Gemini、DeepSeek、
Kimi、GLM、Qwen、Doubao、MiniMax、Grok、Mistral、OpenRouter 與
OpenCode Zen/Go，另外還有一個自訂欄位可填入任何相容 OpenAI 的服務端點
（base URL + 金鑰），也包括本機模型伺服器。搜尋與媒體功能則各自在
**AI 媒體與搜尋**下設定獨立的服務商：網頁搜尋可選 Serper 或 Tavily，
圖片生成與圖片／影片分析可選 OpenAI、Gemini、Doubao/Seedream、GLM、
Grok、Qwen、MiniMax，或任何相容 OpenAI 的圖片服務端點。

整套軟體支援亮色、深色與跟隨系統三種主題。主題只會改變畫面顯示：
匯出、列印與儲存的檔案永遠保留文件本身的原始色彩。

<a id="download"></a>

## 下載

| 平台                                 | 系統需求                                               | 下載                                                                                     |
| ------------------------------------ | ------------------------------------------------------ | ---------------------------------------------------------------------------------------- |
| **macOS** — Apple Silicon（arm64）   | macOS 11 以上                                          | [最新版 `.dmg`（arm64）](https://github.com/genspark-ai/genoffice/releases/latest)       |
| **macOS** — Intel（x64）             | macOS 11 以上                                          | [最新版 `.dmg`（x64）](https://github.com/genspark-ai/genoffice/releases/latest)         |
| **Windows**（x64，一般 PC）          | Windows 10 以上，Intel/AMD                             | [最新版 `-x64.exe` 安裝程式](https://github.com/genspark-ai/genoffice/releases/latest)   |
| **Windows** on Arm（ARM64）          | Windows 11 on Arm（Snapdragon X 及同類機型）           | [最新版 `-arm64.exe` 安裝程式](https://github.com/genspark-ai/genoffice/releases/latest) |
| **Linux** — Debian / Ubuntu          | x86_64，glibc 2.34+（Ubuntu 22.04 或以上）             | [最新版 `.deb`](https://github.com/genspark-ai/genoffice/releases/latest)                |
| **Linux** — Fedora / RHEL / openSUSE | x86_64，glibc 2.34+（Fedora 35+、RHEL 9+、Leap 15.6+） | [最新版 `.rpm`](https://github.com/genspark-ai/genoffice/releases/latest)                |
| **Linux** — 其他發行版               | x86_64，glibc 2.34+，FUSE 2                            | [最新版 `.AppImage`](https://github.com/genspark-ai/genoffice/releases/latest)           |

所有版本皆由 `main` 分支建置，macOS 與 Windows 安裝程式皆有簽署。
較舊版本可在[發行頁面](https://github.com/genspark-ai/genoffice/releases)找到。

<details>
<summary><b>在 Linux 上安裝</b></summary>

deb 套件可用 apt 安裝——會自動處理相依套件，並將 GenOffice 加入
應用程式選單：

```bash
sudo apt install ./genoffice_<version>_amd64.deb
```

在 Fedora / RHEL 家族 / openSUSE 上，改安裝 rpm 套件：

```bash
sudo dnf install ./genoffice-<version>.x86_64.rpm     # Fedora / RHEL 家族
sudo zypper install ./genoffice-<version>.x86_64.rpm  # openSUSE
```

AppImage 可直接就地執行：先安裝 FUSE 2 執行環境
（`sudo apt install libfuse2`；Ubuntu 24.04 上的套件名稱是
`libfuse2t64`），再將檔案設為可執行，然後執行它：

```bash
chmod +x GenOffice-<version>.AppImage
./GenOffice-<version>.AppImage
```

</details>

## 運作原理

七個 Electron 應用程式——Docs、Sheets、Slides、PDF、Markdown、HTML
與分頁式外殼——共用一層純 TypeScript 套件所構成的引擎，再加上一個
處理 `.xlsx` 的 Rust sidecar。原始檔案永遠是唯一的真相來源：編輯以
局部差異的方式套用，編輯器沒動過的內容在往返過程中完全不受影響。

```
open docx ─► 依雜湊封存原始檔案（永遠不動它）
          ─► 解析 word/document.xml 成區塊樹，每個區塊都對應到原始 XML
          ─► Tiptap 編輯器（手動編輯 + AI 編輯，並追蹤已變更內容）
save      ─► 已變更的區塊 → OOXML 片段（只參照既有樣式）
          ─► 拼接回原始 document.xml；未變更的區塊保留原有位元組
          ─► 重新打包 zip；其他每個項目皆逐位元組複製
```

各套件（docx/pptx 引擎、`pdf2docx`、`html2docx`、代理核心與各服務商）
的逐一介紹，收錄在 [CONTRIBUTING.md](../../CONTRIBUTING.md#engine-packages)。

## 開發

```bash
npm install
npm run fixtures     # 產生測試用的 .docx fixture
npm test             # 引擎 + 應用程式單元測試（docs/sheets/slides 不需要顯示器）
npm run typecheck    # 對整個 workspace 執行 tsc --noEmit
npm run dev          # 六個編輯器 + 外殼，皆跑在 Vite 開發伺服器上
npm run dev:docs     # 只跑單一應用程式（其他 workspace 也是同樣模式）
npm run dist:mac     # 打包 macOS dmg（會重新產生第三方授權聲明）
npm run dist:win     # 打包 Windows nsis 安裝程式
npm run dist:linux   # 打包 Linux AppImage + deb + rpm
```

sheets 應用程式還需要 Rust 工具鏈來編譯其 xlsx sidecar
（`cargo` 需在 PATH 中）；執行 `npm run build -w @genoffice/sheets`
會自動完成編譯。每次變更需要通過哪些檢查、PR 如何合併，詳見
[CONTRIBUTING.md](../../CONTRIBUTING.md)。

## 社群

GenOffice 正在積極開發中，你的意見回饋會影響它的方向。

- 在 [GitHub Issues](https://github.com/genspark-ai/genoffice/issues)
  **回報問題或提出功能需求**。
- 到 [GenTeam](https://genoffice.ai/join) **加入 GenOffice 群組聊天**，
  和團隊及其他使用者交流。
- 如果 GenOffice 對你有幫助，**幫這個專案按個星星**——這是支持這個
  專案最簡單的方式。

## 常見問題

<details>
<summary><b>GenOffice 是免費的嗎？</b></summary>

是的。GenOffice 以 Apache-2.0 授權免費開源——應用程式本身沒有試用期，
也沒有付費方案。

</details>

<details>
<summary><b>GenOffice 能開啟 Microsoft Word、Excel 與 PowerPoint 檔案嗎？</b></summary>

可以。GenOffice 能開啟並儲存原生的 `.docx`、`.xlsx` 與 `.pptx` 檔案。
儲存時採位元級保留：你沒動過的部分會原樣寫回，文件在 Microsoft
Office 中依然能正常使用。

</details>

<details>
<summary><b>GenOffice 能離線使用嗎？</b></summary>

文件編輯完全在本機執行——檔案在開啟、編輯、儲存或轉換的過程中
都不會離開你的裝置。AI 功能（代理、搜尋、圖片工具）需要網路連線，
可透過 Genspark 登入或使用你自己的模型 API 金鑰。

</details>

<details>
<summary><b>GenOffice 能編輯 PDF 檔案嗎？</b></summary>

可以——這是真正的 PDF 文字與圖片編輯，會重寫頁面內容串流並保留
原始字體，而不是用覆蓋式的注解來遮蓋。

</details>

<details>
<summary><b>GenOffice 能把 PDF 轉換成 Word、Excel 或 PowerPoint 嗎？</b></summary>

可以——完全在本機完成：透過 PDFium 進行字元層級擷取，再搭配以幾何
資訊為基礎的版面分析，不經雲端服務，也不需要上傳檔案。掃描頁面
同樣支援：在 macOS 與 Windows 上會由系統 OCR 讀取內容，因此轉出的
是可編輯文字，而不是頁面圖片。

</details>

<details>
<summary><b>我可以使用自己的 AI 模型或 API 金鑰嗎？</b></summary>

可以。除了免金鑰的 Genspark 登入之外，GenOffice 也支援自帶金鑰，
涵蓋 Claude、OpenAI、Gemini、DeepSeek、Kimi、GLM、Qwen、Doubao、
MiniMax、Grok、Mistral、OpenRouter 與 OpenCode Zen/Go，以及任何相容
OpenAI 的服務端點——包括本機模型伺服器。搜尋、圖片生成與圖片／影片
分析則在「設定 → AI 媒體與搜尋」下使用各自獨立的金鑰。

</details>

<details>
<summary><b>GenOffice 能把 HTML 轉換成 Word 嗎？</b></summary>

可以——HTML 應用程式中的「匯出為 Word」會完全在本機產生原生、
可編輯的 `.docx`。頁面會在內建的 Chromium 中渲染，再轉換成真正的
Word 結構：標題、段落、清單、表格、卡片、KPI 列、表單欄位與頁面
背景；只有沒有對應 Word 結構的視覺元素（圖表、圖示、裝飾方塊）
才會以圖片形式嵌入。

</details>

<details>
<summary><b>GenOffice 會收集任何資料嗎？</b></summary>

官方打包版本預設會傳送有限的使用分析資料，你隨時可以在「設定 → 一般」
中關閉這項回報。分析資料絕不會包含文件內容、檔案名稱、檔案路徑、
帳號身分或電子郵件地址。完整的事件與資料揭露內容請見
[GenOffice 隱私權政策](../../PRIVACY.md)。

</details>

## 安全性

行程層級的安全防護措施（渲染器沙箱化、IPC 驗證、外部連結攔截）
以及針對 AI 生成內容的威脅模型，詳見 [SECURITY.md](../../SECURITY.md)。

## 致謝

沒有以下這些開源專案，就不會有 GenOffice：

- [Electron](https://www.electronjs.org/) — 每個應用程式所使用的桌面
  執行環境。
- [Univer](https://github.com/dream-num/univer)（Apache-2.0）— Sheets
  延伸使用的試算表 UI 核心。
- [PDFium](https://pdfium.googlesource.com/pdfium/)（BSD-3-Clause，
  透過 [@embedpdf/pdfium](https://github.com/embedpdf/embed-pdf-viewer)
  引入）— 支撐真正的 PDF 文字與圖片編輯的內容串流引擎。
- [pdf.js](https://github.com/mozilla/pdf.js)（Apache-2.0）與
  [pdf-lib](https://github.com/Hopding/pdf-lib)（MIT）— PDF 渲染與
  文件組裝。
- [Tiptap](https://tiptap.dev/) / [ProseMirror](https://prosemirror.net/)
  — Docs 與 Markdown 所使用的區塊編輯器。
- [CodeMirror](https://codemirror.net/)（MIT）— HTML 應用程式中的
  原始碼編輯器。
- [Konva](https://konvajs.org/) — Slides 與 Sheets 圖表所使用的畫布
  渲染引擎。
- [HarfBuzz](https://github.com/harfbuzz/harfbuzz)（wasm）— 複雜文字
  系統的文字排版度量。
- [calamine](https://github.com/tafia/calamine) 與
  [IronCalc](https://github.com/ironcalc/IronCalc) — Rust xlsx sidecar
  的讀取與計算層。
- [libeot](https://github.com/umanwizard/libeot)（MPL-2.0）— 用於內嵌
  PowerPoint 字型的 MicroType Express 解碼器，已移植為 TypeScript。
- [React](https://react.dev/)（MIT）— 每個應用程式的 UI 層。
- [Mermaid](https://mermaid.js.org/)（MIT）與 [KaTeX](https://katex.org/)
  （MIT）— Markdown 與 Docs 中的圖表與數學公式。
- [opentype.js](https://opentype.js.org/)（MIT）— 用於度量資訊解析與字形
  查詢的字型解析程式庫。
- [JSZip](https://stuk.github.io/jszip/)（MIT）與
  [fast-xml-parser](https://github.com/NaturalIntelligence/fast-xml-parser)
  （MIT）— OOXML 容器與 XML 層。
- [Fluent UI System Icons](https://github.com/microsoft/fluentui-system-icons)
  （MIT）— 各功能區所使用的圖示集。
- [electron-updater](https://www.electron.build/)（MIT）— 應用程式內更新。
- Liberation、Carlito、Caladea 與 Noto CJK 字型（OFL/Apache-2.0）—
  隨附的文件字型。

執行 `npm run notices` 會重新產生內建的第三方授權摘要
（`tools/gen-third-party-notices.mjs`）；所有執行期相依套件皆採用
MIT/Apache-2.0/BSD-3-Clause/OFL 授權。

## 授權

GenOffice 採用 [Apache License 2.0](../../LICENSE) 授權，唯一例外是
`ee/` 目錄，該目錄保留給未來的企業模組，適用
[GenOffice 企業授權](../../ee/LICENSE)。

GenOffice 與 Genspark 的名稱與標誌皆為 Mainfunc, Inc. 的商標。
Apache-2.0 授權並未授予使用這些商標的權利（詳見第 6 節）；
基於本專案的分支請使用自己的品牌識別。
