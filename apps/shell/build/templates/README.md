# 模板中心资源目录

这里存放模板中心（`window.aiOffice.templatesList()` / `newFromTemplate()`）随应用发布的
真实 Office 文件。内容由 `src/` 下可审阅的源文件生成：

- `src/manifest.json` — 模板清单（id / 标题 / 描述 / 类型 / 分类 / 文件名），`tools/gen-templates.mjs`
  校验后输出为 `templates.json`，主进程的 `templates-ipc.ts` 只读后者。
- `src/<id>.md` — docx 模板的 Markdown 源（`genoffice create --type docx --from`）。
- `src/<id>.json` + `src/<id>.ops.json` — xlsx 模板的表格数据（含公式）与样式操作
  （`create --type xlsx --from` + `sheet apply --ops`）。
- `src/<id>.json`（deck spec）— pptx 模板的幻灯片规格（`create --type pptx --spec`）。
- `src/<id>.md` — 直接复制的 Markdown 模板。

生成产物（本目录顶层的 `.docx/.xlsx/.pptx/.md` 与 `templates.json`）与源文件一同提交进
仓库，修改源后运行：

```
node tools/gen-templates.mjs
```

打包时 `electron-builder.cjs` 把整个目录映射到 `resources/templates`。
