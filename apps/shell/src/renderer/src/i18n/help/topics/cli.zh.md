# 命令行与代理

每个安装版本都带一个 `genoffice` 命令，它驱动的是窗口所用的同一套引擎——同样的解析器、同样的写入器、同样的渲染器。应用保存的文件和命令写出的文件是同一个文件；应用 AI 面板能通过的检查，命令也能通过。

```sh
genoffice --help          # every command
genoffice guide slides    # the op reference, for writing your own
```

## 取得命令

macOS 和 Windows 上它位于应用包内。想直接按名字调用，先运行一次 `genoffice install-cli`：它会把包内的二进制文件软链接到 `/usr/local/bin`，Windows 上则链接到你的用户 `PATH`。

## 值得记住的命令

| 命令                | 作用 |
| ------------------ | ---- |
| `open`             | 在应用里打开文档；应用没在运行时会先启动。 |
| `convert`          | 用应用自己的引擎在格式之间转换。 |
| `create`           | 从结构化内容创建文档。 |
| `render`           | 每页一张 PNG，按渲染器排出的版式输出。 |
| `pdf`              | 逐页读取 PDF 的文字层，不需要启动应用进程。 |
| `info`             | 文档的元数据与结构摘要。 |
| `search`           | 通过应用里配置的提供方做网页或图片搜索。 |
| `image` / `media`  | 生成图片，或描述图片、视频、音频文件并向它们提问。 |
| `merge`            | 填充 `.docx`、`.pptx` 或 `.xlsx` 模板里的 `{{key}}` 占位符。 |
| `capabilities`     | 报告本机配置了哪些云端功能。 |
| `guide`            | 操作参考与设计指南，由执行器校验所依据的同一份定义生成，因此不会与 `apply` 接受的内容产生偏差。`--json` 会带上每个操作的 schema 返回。 |
| `install-cli`      | 把 `genoffice` 放进 `PATH`。 |
| `skill`            | 列出本机找到的编程代理，并在其中安装或更新 GenOffice 技能。 |
| `mcp`              | 把每条命令作为 Model Context Protocol 工具提供。见**连接编程代理**。 |

## 编辑：docs、sheets、slides

`genoffice docs`、`genoffice sheet` 和 `genoffice slides` 都要先接一个子命令——`read`、`apply`、`check`，slides 另有 `audit`、`render`、`replace`——然后通过应用所用的同一条写入路径读写。它们共享一套词汇：一个 **op** 就是一次编辑，一个 **spec** 就是按顺序应用的一组 op。

```sh
genoffice sheet read report.xlsx
genoffice slides apply deck.pptx --ops ops.json --dry-run
genoffice docs apply brief.docx --ops ops.json
```

`--dry-run` 只报告一批操作会做什么，不写入任何内容，这是让一份 spec 落地前先自查最省事的办法。应用里的 AI 面板跑的就是这些 op，所以凡是你能让它做的事，都能写成脚本。

## Model Context Protocol

`genoffice mcp` 把每条命令作为 MCP 工具提供；`genoffice mcp install <agent|all>` 则把它注册到编程代理自己的配置里。这一侧见**连接编程代理**。

## 命令不会做的事

它只负责读写文件。它不是应用：没有窗口，应用内的更新对话框也不适用。凡是需要窗口的事——AI 面板、对渲染后幻灯片的 QC 检查——都得等你打开文件。`genoffice render` 可以在没有窗口的情况下拿到像素，`genoffice slides` 也能自己检查一份演示文稿的版面。
