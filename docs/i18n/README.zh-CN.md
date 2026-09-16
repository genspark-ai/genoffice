<p align="center">
  <a href="https://genoffice.ai/">
    <picture>
      <source srcset="../assets/readme/hero-dark.webp" media="(prefers-color-scheme: dark)">
      <img src="../assets/readme/hero.webp" alt="GenOffice —— 开源 AI 办公套件：Docs、Sheets、Slides、PDF、Markdown 和 HTML，内置 AI 面板" width="100%">
    </picture>
  </a>
</p>

<h1 align="center">GenOffice</h1>

<p align="center"><b>全球首款功能完整的开源 AI 办公套件。</b><br>
Word、Excel、PowerPoint 和 PDF 文件，由你与你的 AI 共同编辑，并以原生格式保存。</p>

<p align="center">
  <a href="../../LICENSE"><img src="https://img.shields.io/github/license/genspark-ai/genoffice" alt="许可证：Apache-2.0"></a>
  <a href="https://github.com/genspark-ai/genoffice/releases/latest"><img src="https://img.shields.io/github/v/release/genspark-ai/genoffice" alt="最新版本"></a>
  <a href="https://github.com/genspark-ai/genoffice/releases"><img src="https://img.shields.io/github/downloads/genspark-ai/genoffice/total" alt="下载量"></a>
  <a href="https://github.com/genspark-ai/genoffice/stargazers"><img src="https://img.shields.io/github/stars/genspark-ai/genoffice?style=flat" alt="GitHub Star 数"></a>
</p>

<p align="center"><a href="../../README.md">English</a> · <a href="README.es.md">Español</a> · <a href="README.pt-BR.md">Português (Brasil)</a> · <a href="README.de.md">Deutsch</a> · <a href="README.fr.md">Français</a> · <b>简体中文</b> · <a href="README.zh-TW.md">繁體中文</a> · <a href="README.ko.md">한국어</a> · <a href="README.ja.md">日本語</a> · <a href="README.ar.md">العربية</a> · <a href="README.ru.md">Русский</a> · <a href="README.it.md">Italiano</a> · <a href="README.nl.md">Nederlands</a> · <a href="README.pl.md">Polski</a> · <a href="README.cs.md">Čeština</a> · <a href="README.id.md">Bahasa Indonesia</a> · <a href="README.ms.md">Bahasa Melayu</a> · <a href="README.th.md">ไทย</a> · <a href="README.hi.md">हिन्दी</a> · <a href="README.he.md">עברית</a></p>

<p align="center">
  <a href="#download"><b>下载</b></a> ·
  <a href="#command-line-and-agent-skill"><b>CLI</b></a> ·
  <a href="#mcp-server"><b>MCP</b></a> ·
  <a href="https://genoffice.ai/"><b>官网</b></a> ·
  <a href="https://genoffice.ai/join"><b>社区</b></a> ·
  <a href="../../PRIVACY.md"><b>隐私</b></a>
</p>

GenOffice 是一款免费、开源的 Microsoft Office 替代品，支持 macOS、Windows
和 Linux。它可以直接打开并保存原生 `.docx`、`.xlsx` 和 `.pptx` 文件，编辑
PDF、Markdown 和 HTML，并在每份文档旁边配备一个 AI 智能体 —— 不是在侧边硬
塞一个聊天框，而是一个真正会读取文件、执行修改，并清楚展示改动之处的编辑器。

- **原生格式，字节级保留。** 只有你编辑过的部分才会被重写，文件中的其他内容
  全部逐字节保留，文档在 Word、Excel 和 PowerPoint 中照常可用。
- **可审阅的 AI。** 修改以修订痕迹和差异对比的形式呈现，一键即可回滚。表格
  得到的是实时公式，而不是粘贴进来的数字；幻灯片和网页直接生成到画布上，
  仍然可以随意编辑。
- **本地优先的设计。** 文件的打开、编辑、保存和转换都在你的电脑上完成。
  PDF → Word / Excel / PowerPoint、Markdown → Word、HTML → Word 全部在本机运行。
  只有 AI 调用会发送到你选择的服务商。
- **用自己的密钥，或者不用密钥。** 登录 Genspark 即可免配置使用；也可以自带
  Claude、OpenAI、Gemini、DeepSeek、Kimi、GLM、Qwen、Doubao、MiniMax、Grok、
  Mistral、OpenRouter、Requesty 的密钥，或任何 OpenAI 兼容端点，包括本地模型服务。
- **可脚本化，随时可供智能体调用。** 应用自带 `genoffice` 命令行和一份面向
  Claude Code、Codex、Cursor、Gemini CLI、GitHub Copilot、OpenCode 和 Windsurf
  的 agent skill，让编程智能体无需打开任何窗口，就能在你的电脑上创建、转换、
  读取和编辑真正的 Office 文件。

**获取：** [macOS](https://github.com/genspark-ai/genoffice/releases/latest)（Apple Silicon 和 Intel）·
[Windows](https://github.com/genspark-ai/genoffice/releases/latest)（x64 和 Arm）·
[Linux](https://github.com/genspark-ai/genoffice/releases/latest)（deb、rpm、AppImage）——
详情与系统要求见[下载](#download)。

## 演示

六个应用，一个 AI 面板，再加一条供编程智能体使用的命令行。每张截图都是 macOS
上的真实应用，AI 的操作均由面板中可见的提示词驱动。

### 1 · Docs —— 打开并编辑 `.docx`，AI 改动可审阅

<table>
<tr>
<td width="50%"><img src="../assets/readme/docs-report.webp" alt="GenOffice Docs 以 80% 缩放渲染一页双栏年报，包含通栏封面图、底纹 KPI 表格、页眉和页脚，AI 面板处于折叠状态"></td>
<td width="50%"><img src="../assets/readme/docs-ai.webp" alt="GenOffice Docs：一份带横幅图片的公司简介；AI 精简了「概述」部分并插入了一个新的项目符号章节，面板提供一键回滚"></td>
</tr>
<tr>
<td><b>按 Word 的排版方式打开文件</b> —— 双栏章节、出血图片、底纹表格、页眉页脚，并按 Word 的行度量分页。样式、批注、修订、公式和墨迹原样往返，不受影响。</td>
<td><b>说出你想要的修改</b> —— AI 读取所需的内容块，重写「概述」并插入一个新的项目符号章节。每一轮 AI 操作都是一个可回滚的快照；开启<b>修订</b>后，改动会以 Word 风格的修订形式呈现。</td>
</tr>
</table>

### 2 · Sheets —— `.xlsx` 中的实时公式与图表，而不是粘贴的数字

<table>
<tr>
<td width="50%"><img src="../assets/readme/sheets-ai.webp" alt="GenOffice Sheets：AI 新增了一张「汇总」工作表，用 SUMIF 公式按地区和类别统计收入，并插入了柱状图，报告已应用 43 项更改并提供撤销按钮"></td>
<td width="50%"><img src="../assets/readme/sheets-qa.webp" alt="GenOffice Sheets：被问到哪个地区在第二季度收入领先时，AI 回答「欧洲」并给出类别拆分，同时以链接形式引用所用的单元格，旁边是「订单」工作表"></td>
</tr>
<tr>
<td><b>让它来做</b> —— 只需一句话，智能体就会新增一张「汇总」工作表，按地区和类别写入真正的 <code>SUMIF</code> 公式，插入柱状图，并把这 43 项更改作为一个可撤销的批次一次应用。</td>
<td><b>向它提问</b> —— 关于工作簿的问题会连同推理过程一起返回，所用的具体单元格以可点击的引用形式给出。底层是自研的 Rust <code>.xlsx</code> 引擎，支持数据透视表、切片器、条件格式和公式追踪。</td>
</tr>
</table>

### 3 · Slides —— 从一句提示词到一份 `.pptx` 演示文稿

<img src="../assets/readme/slides-generate.webp" alt="GenOffice Slides 生成 Aurora Home 投资人演示文稿的过程延时：AI 在面板中规划叙事线，幻灯片依次出现在画布上，最终演示文稿以收尾的诉求结束" width="100%">

<table>
<tr>
<td width="50%"><img src="../assets/readme/slides-cover.webp" alt="GenOffice Slides：画布上是 AI 生成的 Aurora Home 投资人演示文稿封面页，面板中显示最初的一行提示词以及 AI 对所生成内容的总结"></td>
<td width="50%"><img src="../assets/readme/slides-ai.webp" alt="GenOffice Slides：同一份 11 页演示文稿中经过设计的结尾页，左侧是缩略图栏，AI 面板正在总结故事线"></td>
</tr>
<tr>
<td><b>输入一句话</b> —— “为 Aurora Home 制作一份 10 页的投资人融资演示文稿……”。GenOffice 规划故事线、调研数据，并把每一页直接生成到画布上，成为真正的 <code>.pptx</code>。</td>
<td><b>产出一份完整的演示文稿</b> —— 十一页经过设计的幻灯片，字体、配图风格统一，并以行动号召收尾；可以继续用母版、版式、智能参考线和无损裁剪编辑，也可以让面板重新设计风格、改写文字或调整顺序。</td>
</tr>
</table>

### 4 · PDF —— 原位编辑 PDF 文字，本机将 PDF 转为 Word

<table>
<tr>
<td width="50%"><img src="../assets/readme/pdf-edit.webp" alt="GenOffice PDF：「编辑文字」模式为页面上每个文本块勾勒出边框以便原位编辑，同时 AI 面板正在回答关于这份报告的问题并附带页码引用"></td>
<td width="50%"><img src="../assets/readme/pdf-convert.webp" alt="GenOffice Docs 显示一份由 Helios 季度回顾 PDF 在本机转换而来的 Word 文档，在原 PDF 旁边的第二个标签页中打开"></td>
</tr>
<tr>
<td><b>在页面内直接编辑</b> —— 「编辑文字」模式为每个文本块勾勒边框，可原位重新输入；内容流通过 PDFium 用原始字体重写，而不是用注释遮盖。就一份长报告向 AI 提问，可以得到带页码引用的回答。</td>
<td><b>本机转换</b> —— <b>PDF 转换器 → PDF 转 Word</b> 生成可编辑的 <code>.docx</code>，在 Docs 中于源文件旁边打开，标题、数据行和段落都完整保留。转换为 Excel 和 PowerPoint 的方式相同；扫描页会经由系统 OCR 处理。</td>
</tr>
</table>

### 5 · HTML —— 先出设计简报的 AI 网页与 UI 构建器

说明这个页面的用途和受众。AI 会先提出一份**设计简报** —— 主视觉钩子、配色、
字体和风格方向 —— 然后依据这些设计令牌构建一个单文件、自包含的 `.html`。

<img src="../assets/readme/html-restyle-motion.webp" alt="GenOffice HTML 重新设计 Lumen 落地页的过程延时：面板中一次 Restyle 请求就把深色的 Midnight Studio 页面变成暖色的 Solar Daybreak 版本，而每个板块和全部文案都保持不变" width="100%">

<table>
<tr>
<td width="50%"><img src="../assets/readme/html-ai.webp" alt="GenOffice HTML：为一款太阳能台灯生成的落地页，采用深色的 Midnight Studio 风格方向，显示在实时预览中，AI 面板正在总结刚构建的页面"></td>
<td width="50%"><img src="../assets/readme/html-restyle.webp" alt="同一个 Lumen 落地页被 AI 重新设计为暖色的 Solar Daybreak 方向：纸质背景、衬线标题和橙色强调色，所有章节和文案全部保留"></td>
</tr>
<tr>
<td><b>一句提示词生成</b> —— 为 Lumen 生成的醒目主视觉、功能卡片、定价和候补名单表单，采用 Midnight Studio 风格方向。点击任意元素即可重新设计样式，双击编辑文字，或切换到 CodeMirror 源码视图。</td>
<td><b>同一设计，全新方向</b> —— 一次<b>重新设计风格</b>请求即可替换简报中的设计令牌，页面随之变化：暖色纸张、编辑风衬线体、日出橙强调色，文字一字未改。可全屏演示，也可导出为 PDF 或原生可编辑的 Word 文档。</td>
</tr>
</table>
<table>
<tr>
<td width="50%"><img src="../assets/readme/html-dashboard.webp" alt="GenOffice HTML：为一位自由设计师生成的个人工作台 UI，采用暖色亚麻风格，包含左侧导航栏、衬线体问候语和四张指标卡片"></td>
<td width="50%"><img src="../assets/readme/html-report.webp" alt="GenOffice HTML：以大报风格生成的电动车市场数据报告，包含衬线体报头、1730 万的头条数字和一行统计数据"></td>
</tr>
<tr>
<td><b>UI 原型</b> —— 「个人工作台」起手式把一个用户画像变成可用的布局：左侧导航栏、问候语、计费工时迷你图、发票和利用率卡片，全部是可以直接交给开发者的真实 HTML。</td>
<td><b>数据故事</b> —— 「数据报告」起手式构建一份编辑风格的大报版面：衬线体报头、一个头条数字、以分隔线区分的统计行、内联 SVG 图表和方法论说明。</td>
</tr>
</table>

### 6 · Markdown —— 基于纯 `.md` 的块编辑器，带 Ask AI

<table>
<tr>
<td width="50%"><img src="../assets/readme/markdown-ai.webp" alt="GenOffice Markdown：选中一个段落后弹出 Ask AI 浮层，其中有已输入的指令和「润色」「更简洁」「扩写」「修正语法」等建议芯片，以及「立即发送」和「加入队列」按钮"></td>
<td width="50%"><img src="../assets/readme/markdown-render.webp" alt="GenOffice Markdown 渲染一份发布说明文档，包含表格、Mermaid 流程图和任务列表，左侧是 AI 面板的起手提示词"></td>
</tr>
<tr>
<td><b>就选中内容询问 AI</b> —— 选中任意段落，就会出现 <b>Ask AI</b> 芯片：输入指令或选择建议，立即发送，或者把多处锚定的修改加入队列后一次执行。每个应用都有同样的入口。</td>
<td><b>所见即渲染，保存为纯 Markdown</b> —— 标题、列表、表格、图片、代码块和 Mermaid 图表都在 Tiptap 块编辑器中呈现，写回时仍是纯 <code>.md</code>，并支持完全本地的 <b>Markdown → Word</b> 导出。</td>
</tr>
</table>

### 7 · CLI —— 让你的编程智能体在本机驱动 GenOffice

GenOffice 自带 `genoffice` 命令行和一份 agent skill。安装这份 skill 之后，
Claude Code、Codex、Cursor、Gemini CLI、GitHub Copilot、OpenCode 或 Windsurf
就能通过与应用相同的引擎创建、转换、读取和编辑真正的 Office 文件，无需打开
任何窗口。

<img src="../assets/readme/cli-deck-in-app.webp" alt="GenOffice Slides 中显示一份由编程智能体通过 genoffice 命令行制作的八页太阳系演示文稿：画布上是封面页，左侧是八张缩略图，AI 面板处于打开状态" width="100%">

<table>
<tr>
<td width="50%"><img src="../assets/readme/cli-slides-grid.webp" alt="太阳系演示文稿渲染出的八页幻灯片并排展示：封面、探索时间线、四个关键数字、行星直径柱状图、岩质行星与巨行星对比、太阳占 99.8% 的主视觉数字、四大巨行星网格和要点总结"></td>
<td width="50%"><img src="../assets/readme/cli-integrations.webp" alt="GenOffice 设置的「集成」页面：genoffice skill 已安装到 Claude Code，Codex 和 Cursor 旁边显示着「安装」按钮"></td>
</tr>
<tr>
<td><b>给智能体一句提示词</b> —— “制作一份关于太阳系的八页演示文稿。”智能体读取 skill，依次写出样式表、大纲和每页一份的页面规格，用 <code>genoffice image</code> 生成两张图片，再由 <code>genoffice slides check</code> 拦下所有溢出或重叠的内容，最后由 <code>genoffice create</code> 组装出 <code>.pptx</code>，并用 <code>slides render</code> 为每一页交回一张 PNG 供查看。</td>
<td><b>在设置 → 集成中一次安装</b> —— GenOffice 会列出在这台电脑上找到的编程智能体，并把 skill 写入你选中的每一个。也可以把 skill 下载为 zip 压缩包，或运行 <code>npx skills add genspark-ai/genoffice</code>。命令和完整工作流见<a href="#command-line-and-agent-skill">命令行与 agent skill</a>。</td>
</tr>
</table>

### 8 · MCP —— 通过 Model Context Protocol 提供的同一套工具

每一条 `genoffice` 命令同时也是一个 MCP 工具。Claude Code、Claude Desktop、
Cursor 及任何其他 MCP 客户端都能自行启动 `genoffice mcp`，无需安装 skill，
也无需打开任何窗口，即可获得 29 个工具以及以资源形式提供的操作参考文档。
应用内置的第二个 HTTP 服务器，还能让智能体在一个可见的编辑器标签页中构建
Word 文档，供你实时查看。

<img src="../assets/readme/mcp-deck-motion.webp" alt="Claude Code 通过 genoffice MCP 服务器构建一份八页可再生能源投资简报的过程延时：它会搜索图表和照片，并用 media 检查每一张候选图片，deck_start 写入样式表和大纲，deck_page 逐页添加经过检查的页面，deck_build 组装出 .pptx，slides_render 返回每一页的图片；最终这份演示文稿会在 GenOffice Slides 中打开" width="100%">

<table>
<tr>
<td width="50%"><img src="../assets/readme/mcp-deck-in-app.webp" alt="GenOffice Slides 中显示 Claude Code 通过 genoffice MCP 服务器构建的八页可再生能源 2026 演示文稿：画布上是带有风电场照片的封面页，左侧是八张缩略图"></td>
<td width="50%"><img src="../assets/readme/mcp-integrations.webp" alt="GenOffice 设置的「集成」页面，MCP 部分：Claude Code 对应的一行 claude mcp add 命令，Cursor、Claude Desktop 及其他 MCP 客户端对应的 JSON 代码块，以及下方的本地 HTTP 服务器选项"></td>
</tr>
<tr>
<td><b>一句提示词，三十八次工具调用，不用终端</b> —— “制作一份关于 2026 年可再生能源的八页投资简报，封面和其他合适的地方都配上真实照片。”智能体先用 <code>search</code> 搜集图表和照片，再用 <code>media</code> 确认每一张候选图片是不是真实照片，然后带着样式表和大纲调用 <code>deck_start</code>，接着逐页调用 <code>deck_page</code>；每一页都会先对照大纲和配色方案检查过关才会被保留，接着 <code>deck_build</code> 组装出 <code>.pptx</code>，<code>slides_audit</code> 检查是否溢出，<code>slides_render</code> 为每一页交回一张供模型查看的 PNG 图片内容，最后 <code>deck_replace</code> 修复了它不满意的三页。</td>
<td><b>在设置 → 集成中连接一次即可</b> —— 复制 Claude Code 的 <code>claude mcp add</code> 命令行，或把 JSON 代码块粘贴进 Cursor、Claude Desktop 或任何其他 MCP 客户端。选项 B 会开启本地 HTTP 服务器，用于可见的 Word 编辑器。两者都在<a href="#mcp-server">MCP 服务器</a>中有详细说明。</td>
</tr>
</table>

## 为什么选择 GenOffice

- **开源**，Apache-2.0 协议，在 GitHub 上公开开发。
- **由你掌控。** macOS、Windows 和 Linux 原生应用；文件留在本地磁盘，每一次编辑、
  保存和转换都在你的机器上完成。
- **真正的 Office 文件。** 原生 `.docx`、`.xlsx` 和 `.pptx`，字节级保留：未改动
  的部分原样复制。
- **直接编辑文档的 AI。** Docs 中的修订痕迹、Sheets 中实时生效的公式和图表、直接
  在画布上生成的幻灯片，每一次 AI 操作都会留下可回滚的快照。
- **自带模型，自带密钥。** 使用 Genspark 登录，或带上 Claude、OpenAI、Gemini、
  DeepSeek 等的密钥，同样支持本地服务器和任意 OpenAI 兼容端点。
- **认真做好 PDF。** 在页面内直接编辑文字，本机将 PDF 转换为 Word、Excel 或
  PowerPoint，扫描件支持系统 OCR。
- **同样支持 Markdown 和 HTML**，共用同一个 AI 面板，并可本机导出为 Word。
- **可脚本化。** `genoffice` 命令行、agent skill 和 MCP 服务器让每一个引擎都能
  为 Claude Code、Claude Desktop、Codex、Cursor 及其他智能体所用，且仍然在本机
  运行。
- **免费**，个人和团队皆可使用。

## AI 后端

**登录 Genspark**，无需任何配置：模型调用通过 Genspark 代理路由（Claude、GPT
和 Gemini 系列），智能体同时获得网页与图片搜索、图片生成，以及图片/音频/视频
解析能力。

**或者自带密钥。** 设置 → AI 中列出了 Claude、OpenAI、Gemini、DeepSeek、Kimi、
GLM、Qwen、Doubao、MiniMax、Grok、Mistral、OpenRouter、Requesty 和 OpenCode Zen/Go，另有
一个自定义槽位可接入任何 OpenAI 兼容端点（Base URL + 密钥），包括本地模型服务。
搜索和媒体能力在 **AI 媒体与搜索** 下按能力分别配置服务商：网页搜索可选 Serper
或 Tavily；图片生成和图片/视频解析可选 OpenAI、Gemini、Doubao/Seedream、GLM、
Grok、Qwen、MiniMax 或任何 OpenAI 兼容的图片端点。

整个套件提供浅色、深色和跟随系统三种主题。主题只改变屏幕上的显示：导出、打印
和保存的文件始终保留文档自身的颜色。

<a id="command-line-and-agent-skill"></a>

## 命令行与 agent skill

应用能对文件做的一切，`genoffice` 命令行都能在终端里完成：检查、转换、创建、
读取和编辑 Word、Excel、PowerPoint、PDF、Markdown 和 HTML，用的是同一套引擎，
无需界面。它随 GenOffice 一起安装，不需要自己的运行时，也绝不会把文档发送到
任何地方。配合内置的 **agent skill**，它能把编程智能体变成一名文档工作者，
产出真正的 Office 文件，而不是用 Markdown 勉强凑合。

**支持：** Claude Code、Codex、Cursor、Gemini CLI、GitHub Copilot、OpenCode
和 Windsurf 开箱即用，任何其他能读取 skill 的智能体，以及通过
[MCP 服务器](#mcp-server) 支持的 Claude Desktop 和所有 MCP 客户端。

### 安装 skill

| 方式                                   | 效果                                                                                                                            |
| -------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------- |
| 应用内的 **设置 → 集成**               | 列出在这台电脑上找到的智能体；一键即可把 skill 写入你选中的每一个。当 GenOffice 新版本附带更新的 skill 时，会出现**更新**按钮。 |
| 同一页面上的 **下载为 zip**            | claude.ai、Claude 桌面应用及其他助手可作为上传 skill 接受的目录结构。                                                           |
| `npx skills add genspark-ai/genoffice` | 从本仓库安装到任何兼容 skills 的智能体。                                                                                        |

然后开启一个新对话，让它帮你做一份文档。这份 skill 会教智能体何时该用
`genoffice`、编辑前如何先读取文件，以及如何检查自己的成果。

### 从终端快速上手

```bash
genoffice --version
genoffice info report.docx --json                  # headings and blocks; or sheets, slides, pages
genoffice convert report.md --to pdf               # md/html/docx/xlsx/pptx → pdf, pdf → docx/xlsx/pptx, …
genoffice create --type docx --from notes.md --out notes.docx
genoffice create --type xlsx --from table.json --out sales.xlsx   # "=SUM(B2:B9)" cells stay live formulas
genoffice docs read report.docx --range 0-9 --json # then `docs apply --ops edits.json` edits in place
genoffice render report.docx --out shots/          # one PNG per page, to look at what you made
genoffice open sales.xlsx                          # hand the result to the editor
```

每条命令都会打印一行摘要，加上 `--json` 则输出单个 JSON 对象。编辑是原子性的：
被拒绝的操作不会改动文件，并会返回带有指引的错误信息。`genoffice help` 列出
当前的命令集合；完整参考见 [packages/cli/README.md](../../packages/cli/README.md)。

### 智能体实际执行了什么

上方演示中的太阳系演示文稿，在 Claude Code 里只用了一句提示词。在它背后，
智能体按照 skill 的分阶段工作流推进，而 CLI 在每个阶段开始下一步之前都会先做
检查：

```bash
genoffice capabilities --json                        # which cloud tools GenOffice has configured
genoffice guide slides design                        # the deck workflow and layout library
genoffice image "the eight planets in a row …" --aspect 16:9 --out deck/assets/cover.jpg
genoffice slides check deck/outline.json --json      # 8 pages, no findings
genoffice slides check deck/pages/01.json --json     # builds one slide, audits overflow and overlap
…                                                    # one page file per slide, fixed until each check is clean
genoffice create --type pptx --spec deck/pages --outline deck/outline.json --out deck/solar-system.pptx --json
genoffice slides render deck/solar-system.pptx --out deck/shots --json
genoffice slides audit deck/solar-system.pptx --json    # 8 slides, no layout issues
genoffice slides replace deck/solar-system.pptx --slide 4 --spec deck/pages/05.json --json
genoffice open deck/solar-system.pptx
```

`genoffice` 内部不会发起任何模型调用：思考由智能体负责，构建和检查由 CLI
负责，最后的成果会以一份普通的 `.pptx` 在 GenOffice 或 PowerPoint 中打开。

<a id="mcp-server"></a>

### MCP 服务器

同样这些命令也以 [Model Context Protocol](https://modelcontextprotocol.io) 工具
的形式提供，供那些无法运行终端、或者你不想给它终端权限的助手使用。接入方式
有两种，**设置 → 集成 → MCP** 中都提供了可直接复制的代码片段：

| 方式                            | 内容                                                                                                                                                                                                                                                                                                                                                                                                                                                                      |
| ------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| **A · `genoffice mcp`**（推荐） | 一个由助手自行启动的 stdio 服务器，GenOffice 无需处于打开状态。每条命令对应一个工具（`info`、`convert`、`create_docx`、`create_xlsx`、`create_pptx`、`create_pdf`、`docs_read` / `docs_apply` / `docs_check`、`sheet_*`、`slides_*`、`render`、`guide`、`search`、`image`、`media`、`open`），再加上分阶段的幻灯片构建流程 `deck_start` → `deck_page` → `deck_build` → `deck_replace`。Ops、spec 和 Markdown 都以内联方式传入，因此没有文件系统的客户端同样可以正常工作。 |
| **B · 本地 HTTP 服务器**        | 运行在 GenOffice 应用内部，地址为 `http://127.0.0.1:3093/mcp`（Streamable HTTP，兼容旧版 SSE）。它的工具驱动一个可见的 Word 编辑器标签页：`create_session`、`insert_content`、`replace_blocks`、`apply_ops`、`read_document`、`save_session`，你可以实时看着文档成形。默认关闭；在同一个设置面板中开启。                                                                                                                                                                  |

```bash
# Claude Code
claude mcp add --transport stdio genoffice -- genoffice mcp
```

```jsonc
// Cursor、Claude Desktop 或任何其他 MCP 客户端
{ "mcpServers": { "genoffice": { "command": "genoffice", "args": ["mcp"] } } }
```

这里的 `genoffice` 就是应用内置的那个 CLI（在 macOS 上位于
`/Applications/GenOffice.app/Contents/Resources/cli/genoffice`；设置面板会
打印出你本机安装的确切路径）。这个服务器自带工作流说明，并把操作参考文档以
`genoffice://guide/*` 资源的形式暴露出来，因此不需要额外的 skill；skill 和
MCP 服务器可以共存，由助手自行选择使用哪一个。云端能力（`search`、`image`、
`media`）仍然经由 GenOffice 中配置的服务商完成；其余一切都在本机运行，
`GENOFFICE_ALLOWED_ROOTS` 会把每个工具都限制在你列出的文件夹之内。

上方演示中的可再生能源演示文稿，就是仅接入 `genoffice` MCP 服务器时，
Claude Code 里一句提示词从协议层面呈现出的样子：

```text
capabilities · guide(slides, spec) · guide(slides, design)
search(query) ×4                         → IEA, BNEF and IRENA figures for the slides
search(query, images) ×7 · media(url, ask) ×7
                                         → candidate photos, each one checked to be a real photograph
deck_start(dir, style, outline)          → outline checked: 8 pages to write
deck_page(dir, 0, page) … deck_page(dir, 7, page)
                                         → each page checked against the outline and the palette; one page sent again
deck_build(dir, out)                     → renewables-2026.pptx, no image failures
slides_audit(file) · slides_render(file, out)
                                         → no layout findings; 8 PNGs come back as image content
deck_replace(dir, n, page) ×3 · slides_render(file, out)
                                         → three pages fixed after looking at the renders
```

三十八次调用，大约十三分钟，助手全程没有碰过终端：图表、照片、指南、检查和渲染结果都是作为 MCP 工具结果传递的。只有 `search` 和 `media` 离开过本机，发往 GenOffice 中配置的服务商。

<a id="download"></a>

## 下载

| 平台                                  | 系统要求                                               | 下载                                                                                   |
| ------------------------------------- | ------------------------------------------------------ | -------------------------------------------------------------------------------------- |
| **macOS** —— Apple Silicon (arm64)    | macOS 11+                                              | [最新 `.dmg`（arm64）](https://github.com/genspark-ai/genoffice/releases/latest)       |
| **macOS** —— Intel (x64)              | macOS 11+                                              | [最新 `.dmg`（x64）](https://github.com/genspark-ai/genoffice/releases/latest)         |
| **Windows**（x64，大多数 PC）         | Windows 10+，Intel/AMD                                 | [最新 `-x64.exe` 安装程序](https://github.com/genspark-ai/genoffice/releases/latest)   |
| **Windows** on Arm（ARM64）           | Windows 11 on Arm（Snapdragon X 及同类芯片）           | [最新 `-arm64.exe` 安装程序](https://github.com/genspark-ai/genoffice/releases/latest) |
| **Linux** —— Debian / Ubuntu          | x86_64，glibc 2.34+（Ubuntu 22.04 或更新）             | [最新 `.deb`](https://github.com/genspark-ai/genoffice/releases/latest)                |
| **Linux** —— Fedora / RHEL / openSUSE | x86_64，glibc 2.34+（Fedora 35+、RHEL 9+、Leap 15.6+） | [最新 `.rpm`](https://github.com/genspark-ai/genoffice/releases/latest)                |
| **Linux** —— 其他发行版               | x86_64，glibc 2.34+，FUSE 2                            | [最新 `.AppImage`](https://github.com/genspark-ai/genoffice/releases/latest)           |

所有构建均来自 `main` 分支；macOS 和 Windows 安装程序均已签名。
历史版本见 [Releases](https://github.com/genspark-ai/genoffice/releases) 页面。

<details>
<summary><b>在 Linux 上安装</b></summary>

deb 包通过 apt 安装 —— 会自动拉取依赖并把 GenOffice 添加到应用程序菜单：

```bash
sudo apt install ./genoffice_<version>_amd64.deb
```

在 Fedora / RHEL 系 / openSUSE 上，请改用 rpm 包：

```bash
sudo dnf install ./genoffice-<version>.x86_64.rpm     # Fedora / RHEL family
sudo zypper install ./genoffice-<version>.x86_64.rpm  # openSUSE
```

AppImage 无需安装即可运行：先安装 FUSE 2 运行时
（`sudo apt install libfuse2`；在 Ubuntu 24.04 上软件包名为 `libfuse2t64`），
给文件加上可执行权限，然后运行：

```bash
chmod +x GenOffice-<version>.AppImage
./GenOffice-<version>.AppImage
```

</details>

## 工作原理

七个 Electron 应用 —— Docs、Sheets、Slides、PDF、Markdown、HTML 以及多标签
外壳 —— 共享同一层由纯 TypeScript 包组成的引擎，外加一个处理 `.xlsx` 的 Rust
sidecar。原始文件始终是唯一事实来源：修改以窄范围补丁的形式应用，编辑器未触
碰的一切在往返过程中原样保留。

```
open docx ─► archive original by hash (never touched)
          ─► parse word/document.xml into a block tree, each block anchored to its original XML
          ─► Tiptap editor (manual + AI editing, dirty tracking)
save      ─► dirty blocks → OOXML fragments (referencing existing styles only)
          ─► splice into the original document.xml; untouched blocks keep their bytes
          ─► repack the zip; every other entry is copied byte-for-byte
```

逐个包的导览（docx/pptx 引擎、`pdf2docx`、`html2docx`、智能体核心与服务商）
见 [CONTRIBUTING.md](../../CONTRIBUTING.md#engine-packages)。

## 开发

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

Sheets 应用的 xlsx sidecar 还需要 Rust 工具链（`cargo` 在 PATH 中）；
`npm run build -w @genoffice/sheets` 会自动编译它。每项改动必须通过的检查以及
拉取请求的合入流程见 [CONTRIBUTING.md](../../CONTRIBUTING.md)。

## 社区

GenOffice 正在积极开发中，你的反馈决定它的走向。

- **报告 bug 或提出功能需求**，请到
  [GitHub Issues](https://github.com/genspark-ai/genoffice/issues)。
- **加入 GenOffice 群聊**，在
  [GenTeam](https://genoffice.ai/join) 上与团队和其他用户交流。
- **给仓库点个 Star** —— 如果 GenOffice 对你有用，这是支持项目最好的方式。

## 常见问题

<details>
<summary><b>GenOffice 是免费的吗？</b></summary>

是的。GenOffice 免费且开源，采用 Apache-2.0 许可证 —— 没有试用期，应用本身也
没有付费档。

</details>

<details>
<summary><b>GenOffice 能打开 Microsoft Word、Excel 和 PowerPoint 文件吗？</b></summary>

可以。GenOffice 直接打开并保存原生 `.docx`、`.xlsx` 和 `.pptx` 文件。保存是
字节级保留的：你未改动的部分会逐字节写回，文档在 Microsoft Office 中照常可用。

</details>

<details>
<summary><b>GenOffice 可以离线使用吗？</b></summary>

文档编辑完全在本地进行 —— 打开、编辑、保存和转换文件都不需要把文件发送到任何
地方。AI 功能（智能体、搜索、图片工具）需要网络连接，并且需要登录 Genspark 或
提供你自己的模型 API 密钥。

</details>

<details>
<summary><b>GenOffice 能编辑 PDF 文件吗？</b></summary>

可以 —— 真正的 PDF 文字和图片编辑，会在保留原始字体的前提下重写页面内容流，
而不是用注释遮盖。

</details>

<details>
<summary><b>GenOffice 能把 PDF 转换为 Word、Excel 或 PowerPoint 吗？</b></summary>

可以 —— 完全在本机完成：PDFium 字符级提取加上基于几何的版面分析，不依赖云
服务，不上传。扫描页同样支持：在 macOS 和 Windows 上由系统 OCR 识别，转换结果
是可编辑的文字，而不是一张页面图片。

</details>

<details>
<summary><b>我可以使用自己的 AI 模型或 API 密钥吗？</b></summary>

可以。除了免密钥的 Genspark 登录，GenOffice 还支持自带 Claude、OpenAI、Gemini、
DeepSeek、Kimi、GLM、Qwen、Doubao、MiniMax、Grok、Mistral、OpenRouter、Requesty 和
OpenCode Zen/Go 的密钥，以及任何 OpenAI 兼容端点 —— 包括本地模型服务。搜索、
图片生成和图片/视频解析在 设置 → AI 媒体与搜索 下使用各自的密钥。

</details>

<details>
<summary><b>GenOffice 能把 HTML 转换为 Word 吗？</b></summary>

可以 —— HTML 应用中的「导出为 Word」会完全在本机生成原生、可编辑的 `.docx`。
页面先在内置 Chromium 中渲染，再归约为真正的 Word 结构：标题、段落、列表、
表格、卡片、KPI 行、表单字段和页面背景；只有在 Word 中没有对应结构的视觉元素
（图表、图标、装饰性方框）才会以图片形式嵌入。

</details>

<details>
<summary><b>我可以从 Claude Code、Codex、Cursor 或脚本中驱动 GenOffice 吗？</b></summary>

可以。GenOffice 会安装一个 `genoffice` 命令行，以无界面方式运行同一套引擎：
在终端或脚本中检查、转换、创建、读取和编辑文档，并可通过 `--json` 输出供程序
使用。内置的 agent skill 会教 Claude Code、Codex、Cursor、Gemini CLI、GitHub
Copilot、OpenCode 和 Windsurf 使用它；请在 **设置 → 集成** 中安装。详见
[命令行与 agent skill](#command-line-and-agent-skill)。

</details>

<details>
<summary><b>GenOffice 会收集数据吗？</b></summary>

官方打包版本默认会发送有限的使用统计，你可以随时在 设置 → 通用 中关闭上报。
统计数据绝不包含文档内容、文件名、文件路径、账号身份或电子邮件地址。完整的事件
和数据披露见 [GenOffice 隐私政策](../../PRIVACY.md)。

</details>

## 安全

进程安全态势（渲染进程沙箱、IPC 校验、外部链接拦截）以及针对 AI 生成内容的
威胁模型，见 [SECURITY.md](../../SECURITY.md)。

## 致谢

没有以下开源项目，就不会有 GenOffice：

- [Electron](https://www.electronjs.org/) —— 每个应用的桌面运行时。
- [Univer](https://github.com/dream-num/univer)（Apache-2.0）—— Sheets 所扩展的
  电子表格 UI 内核。
- [PDFium](https://pdfium.googlesource.com/pdfium/)（BSD-3-Clause，经
  [@embedpdf/pdfium](https://github.com/embedpdf/embed-pdf-viewer) 打包）——
  真正的 PDF 文字与图片编辑背后的内容流引擎。
- [pdf.js](https://github.com/mozilla/pdf.js)（Apache-2.0）和
  [pdf-lib](https://github.com/Hopding/pdf-lib)（MIT）—— PDF 渲染与文档组装。
- [Tiptap](https://tiptap.dev/) / [ProseMirror](https://prosemirror.net/) ——
  Docs 和 Markdown 中的块编辑器。
- [CodeMirror](https://codemirror.net/)（MIT）—— HTML 中的源码编辑器。
- [Konva](https://konvajs.org/) —— Slides 以及 Sheets 图表的画布渲染。
- [HarfBuzz](https://github.com/harfbuzz/harfbuzz)（wasm）—— 复杂文字的
  文本整形度量。
- [calamine](https://github.com/tafia/calamine) 和
  [IronCalc](https://github.com/ironcalc/IronCalc) —— Rust xlsx sidecar 的
  读取层与计算层。
- [libeot](https://github.com/umanwizard/libeot)（MPL-2.0）—— 面向嵌入式
  PowerPoint 字体的 MicroType Express 解码器，已移植为 TypeScript。
- [React](https://react.dev/)（MIT）—— 各个应用的 UI 层。
- [Mermaid](https://mermaid.js.org/)（MIT）和 [KaTeX](https://katex.org/)
  （MIT）—— Markdown 与 Docs 中的图表与数学公式。
- [opentype.js](https://opentype.js.org/)（MIT）—— 用于度量信息解析与字形
  查找的字体解析库。
- [JSZip](https://stuk.github.io/jszip/)（MIT）和
  [fast-xml-parser](https://github.com/NaturalIntelligence/fast-xml-parser)
  （MIT）—— OOXML 容器与 XML 层。
- [Fluent UI System Icons](https://github.com/microsoft/fluentui-system-icons)
  （MIT）—— 各功能区所用的图标集。
- [electron-updater](https://www.electron.build/)（MIT）—— 应用内更新。
- Liberation、Carlito、Caladea 和 Noto CJK 字体（OFL/Apache-2.0）—— 随应用
  打包的文档字体。

`npm run notices` 会重新生成随应用打包的第三方许可证摘要
（`tools/gen-third-party-notices.mjs`）；所有运行时依赖均为
MIT/Apache-2.0/BSD-3-Clause/OFL 许可。

## 许可证

GenOffice 采用 [Apache License 2.0](../../LICENSE) 许可，仅有一处例外：`ee/`
目录为未来的企业模块预留，受 [GenOffice 企业许可证](../../ee/LICENSE) 约束。

GenOffice 和 Genspark 的名称与标识是 Mainfunc, Inc. 的商标。Apache-2.0 许可证
并不授予使用它们的权利（见第 6 节）；分叉项目请使用自己的品牌。
