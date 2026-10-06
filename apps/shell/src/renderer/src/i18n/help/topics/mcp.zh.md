# 连接编程代理

GenOffice 说 Model Context Protocol，所以编程代理可以通过应用所用的同一套引擎来读写并渲染你的文档。代理不是靠猜文件格式：它拿到的类型化 op schema，与执行器校验所依据的是同一份定义。

## 在应用里注册

**设置 ▸ 集成**就是做这件事的地方。这个面板分成两半，你只用其中一半、或者两半都用，都行。

**技能。**本机上找到的每个编程代理各占一行——Claude Code、Codex、Cursor、Gemini CLI、Copilot CLI、OpenCode、Windsurf——每行都有**安装**、**更新**和**卸载**，另外还有**安装到其他目录…**、**下载 skill（zip）**与**复制路径**。如果你的助手不在列表里，就让 GenOffice 指向它读取 `SKILL.md` 的那个目录，或者把 zip 存下来，让助手自己装。技能与 MCP 可以并存：助手挑一个用，而它们做的事情完全一样。

**MCP。**提供两条路线：**由助手启动（推荐）**，你把显示出来的配置加到客户端里，助手自己把服务器拉起来；以及**本地 HTTP 服务**，由应用替你运行。无论走哪一条，助手最后都是通过 GenOffice 干活，而你一条命令都不用敲。

## 从命令行注册

同样的事，在终端里也能做——这是高级路径，当代理装在面板找不到的地方，就该用它：

```sh
genoffice mcp install all
```

它会找出本机上的编程代理，把 stdio 服务器条目写进各自的配置文件，其余部分保持原样。

```sh
genoffice mcp list             # where each agent stands
genoffice mcp install cursor    # just one
genoffice mcp uninstall cursor  # take it back out
```

装在不常见位置的代理要加上 `--dir <path>`；`--force` 会重写已经存在的条目。

服务器能接受的东西都在一屏之内——注册、移除与列举这三种形式，用 HTTP 提供服务，以及两个 schema 开关：

![genoffice mcp --help 的真实输出：install、uninstall 与 list 三种形式，以及 --http、--host、--token、--compact-schemas、--dir 和 --force 选项](img/mcp.png)

## 不用助手，自行运行

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

## 为什么要技能

不了解 op 词汇表的代理只能靠猜。技能里带着参考手册和设计指南——和 `genoffice guide` 打印出来的是同一份材料——这样助手写下的 op，它的规格是真的读过的。可以从上面那个面板安装，也可以用终端里的 `genoffice skill`。

## 它在应用里能触达什么

服务器不只限于磁盘上的文件。GenOffice 运行期间，代理还可以通过窗口来干活：

- **`open_in_genoffice`** 在一个标签页里打开文件，并把它切到前台。
- **`open_documents`** 列出你已经打开的每个文档——id、类型、路径，以及有没有未保存的改动——然后读取某个文档的当前内容或把它关掉，默认先保存，除非你让它丢弃。
- **内容类工具**把那个 id（或路径）当作 `document` 参数，于是改动就落在你已经开着的那个标签页里，窗口也会切过去显示它。

有两件事仍然够不着：没有 AI 面板，应用内的更新对话框也不适用。

## 本地 HTTP 服务面板

**本地 HTTP 服务**下，应用可以自己运行服务器，而不是把它丢给助手：一个启用开关、一个**端口**输入框、一个**运行中 / 未运行**的状态指示，以及**后台生成**（直接把文档写到某个路径，不打开界面），再加上一份可复制的**客户端配置示例**。展开**高级**后，会多出两个连接网址——Streamable HTTP 和旧的 SSE——一个**健康检查**网址，以及一个**日志**开关，把服务器和工具的活动记录到本地文件，可以在那里**打开**、**刷新**或**清除**。它只监听 localhost。
