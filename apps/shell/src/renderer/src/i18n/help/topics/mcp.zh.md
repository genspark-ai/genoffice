# 连接编程代理

GenOffice 说 Model Context Protocol，所以编程代理可以通过应用所用的同一套引擎来读写并渲染你的文档。代理不是靠猜文件格式：它拿到的类型化 op schema，与执行器校验所依据的是同一份定义。

## 注册

一般情况下一条命令就够：

```sh
genoffice mcp install all
```

它会找出本机上的编程代理——Claude Code、Codex、Cursor、Gemini CLI、Copilot CLI、OpenCode、Windsurf——把 stdio 服务器条目写进各自的配置文件，其余部分保持原样。

```sh
genoffice mcp list             # where each agent stands
genoffice mcp install cursor    # just one
genoffice mcp uninstall cursor  # take it back out
```

如果某个代理装在不常见的位置，加上 `--dir <path>`；`--force` 会重写已经存在的条目。

## 自行运行

如果客户端在另一台机器上，就改用 HTTP 提供服务：

```sh
genoffice mcp --http 8765 --token <secret>
```

`--host <addr>` 决定它监听的地址。把同一个令牌传给客户端。

走 HTTP 时文件也能传：`PUT /files/<name>` 上传一个文件，每个工具都接受 `http(s)` URL 来代替路径，输出则以下载 URL 返回，文件足够小时也可以作为内嵌资源。op、spec 和 Markdown 在两种方式下都是内联传递的。

## 代理能拿到什么

每条命令都是一个工具。其中值得一看的有：

- **`docs`、`sheet`、`slides`** — 通过应用自己的写入路径读写文件，每次一个 **op**。新建一份演示文稿要走 `deck_start`、`deck_page`、`deck_build`。
- **`render`** — 每页一张 PNG，由应用的渲染器排版，这样代理能直接看幻灯片，而不是靠猜。
- **`pdf`** — 逐页读取 PDF 的文字层，不需要应用进程。
- **`info`** — 元数据与结构摘要，通常是拿到一个陌生文件后最省事的第一通调用。
- **`search`、`image`、`media`** — 应用里配置好的提供方，代理不必再自带密钥。
- **`merge`** — 填充一个 `{{key}}` 模板。

## Schema 与更小的预算

`apply` 和 `create` 会用逐 op 的类型化 schema 来声明它们的 `ops`、`cells` 和 `data` 参数，schema 由 `genoffice guide <domain> --json` 生成。这样很精确，但体积不小。上下文窗口吃紧的客户端可以改要朴素数组：

```sh
genoffice mcp --compact-schemas
GENOFFICE_MCP_COMPACT_SCHEMAS=1 genoffice mcp
```

## 技能

不了解 op 词汇表的代理只能靠猜。`genoffice skill` 会把它找到的代理都装上一份 GenOffice 技能，里面带着参考手册和设计指南——和 `genoffice guide` 打印出来的是同一份材料。

## 它不是什么

MCP 服务器负责读写文件。它不是窗口：没有 AI 面板，应用内的更新对话框也不适用。哪一步需要窗口，就把文件打开。
