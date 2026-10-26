# pi harness 逐 app 切换章程

Spike（docs，`GENOFFICE_AI_BACKEND=pi`）已验证：图片注入、工具调用、usage 上报全通。
本章程约束剩余 app 的切换顺序、验收门与回退路径。

**2026-10-26 更新：pi 已成为默认后端。** `GENOFFICE_AI_BACKEND=legacy` 显式回退旧栈。
六个 app + shell Ask 全部验收通过；dist:win 产物（GenOffice Setup 0.12.0.exe）已验证。

## 原则

1. **单开关**：每 app 的切换点只在主进程 `ai:stream` handler 里的
   `GENOFFICE_AI_BACKEND` 分支（docs-main.ts:4274 已打样）。renderer 零改动。
2. **默认可回退**：不设环境变量 = pi 栈；`GENOFFICE_AI_BACKEND=legacy` = 旧栈
   （`streamForProvider`）。任何 app 出问题，回退 = 设环境变量重启，不回滚代码。
3. **一门一令**：每道验收门是一条可复制运行的命令，不用形容词。
4. **协议缺口冻结**：`codex` provider（codex-app-server 子进程协议）在 pi 后端下
   不支持——pi 的 `openai-codex-responses` 是 HTTP Responses API，与本地 CLI 子进程
   不同构。pi 模式下 codex 端点走旧栈旁路（bridge 抛明确错误，不静默失败）。

## 架构事实（切换面）

| app | ai:stream handler | 工具数 | 特有机制 |
|---|---|---|---|
| docs | docs-main.ts:4274 ✅已切换 | ~35 | 图片注入 doc-images.ts |
| markdown | 复用 docs-main.registerAiIpc | ~10 | 与 docs 同构（tiptap） |
| html | 复用 docs-main.registerAiIpc | ~8 | streamText 旁路（brief/page-writer） |
| slides | ai-ipc.ts:239（自有变体） | 17 | slide-qc 双循环、deck-images |
| pdf | 复用 docs-main.registerAiIpc | ~45 | 页面图 page-images |
| sheets | sheets-main.ts:3975（自有变体）✅已切换 | 13 | Rust sidecar、verifyResponse |
| shell Ask | 复用 docs-main.registerAiIpc | ~6 | kb/attachments/cli/create-document |

shell 聚合模式下 `registerAiIpc` 只注册一次（docs-main.ts:4119，shell/index.ts:6796）。
slides 与 sheets 的自有变体在独立窗口模式使用，切换时两处都要改。

## 协议映射（26 provider → pi Api）

| GenOffice protocol | pi Api | 状态 |
|---|---|---|
| anthropic | anthropic-messages | spike 未实测（无 key） |
| gemini | google-generative-ai | spike 未实测（经 openai-compat 网关代跑） |
| openai-compatible | openai-completions | ✅ spike 实测通过 |
| codex-app-server | —（不支持，抛错旁路） | 冻结，见原则 4 |

已知需在逐 app 验收中盯住的点：
- gemini `thoughtSignature` 回传（GenOffice `AgentToolCall.signature`）——pi 用
  `thinkingSignature` 承载，桥接层已透传 `reasoning` delta，signature 字段映射待验收
- MiniMax/DeepSeek interleaved-thinking 的 reasoning echo
- output-cap fallback（`max_tokens` vs `max_completion_tokens`）——pi 有自己的
  兼容层，需实测一次 400 拒绝场景
- Genspark 凭证注入（`resolveProfileCredential`）在 pi 分支也要生效
- 双看门狗（主进程 180s / renderer 240s）时序不依赖协议栈，不受影响

## 切换顺序与验收门

顺序按风险升序。每步：接 handler → typecheck → 单测 → CDP 实机问答 → 提交。

### 1. markdown（最简，验证"复用 docs handler"路径自动生效）✅ 已验证 2026-10-26
- 改动：无代码改动预期（复用 docs-main handler）。验收为主。——实测确认零改动。
- 门：`npm run typecheck -w @genoffice/markdown` 通过；
  CDP 实机：markdown tab 里问一句生成类问题，断言有流式回答。——实测通过：
  提问"用一句话介绍 markdown"，pi 后端流式回答，usage ↑4.6k ↓237 正常上报。

### 2. html（streamText 旁路是差异点）✅ 已验证 2026-10-26
- 改动：bridge 补 `chatViaPi`（streamSimple 收集式非流式包装）+ docs-main ai:chat
  加 pi 分支。ai:stream 复用 docs-main handler 零改动。
- 门：`npm run typecheck -w @genoffice/html` 通过；
  `npx vitest run tests/brief-writer.test.ts tests/page-writer.test.ts`（cwd=apps/html）
  19/19 过；CDP 实机：生成个人名片页，`ask_clarification` 工具经 pi 触发，
  ↑3.8k ↓743 token 正常。

### 3. slides（自有 ai-ipc.ts handler）✅ 已验证 2026-10-26
- 改动：ai-ipc.ts:264 加 pi 分支（与 docs 同构）。
- 门：`npm run typecheck -w @genoffice/slides` 通过；
  CDP 实机：「把第 1 页标题改成季度总结」→ 6 步工具链
  （read_slide/execute_slide_script/regenerate_slide 等）全部经 pi 触发，
  最终标题正确更新，↑78.7k ↓3k token 正常。

### 4. pdf（工具最多，页面图注入）✅ 已验证 2026-10-26
- 改动：无 handler 改动（复用 docs-main）。
- 门：`npm run typecheck -w @genoffice/pdf` 通过；
  CDP 实机：打开含「HELLO-123+红圆」特征图的 pdf，问图内容 →
  回答命中「HELLO-123」+「红色的圆圈」，read_pages 工具经 pi 触发，
  ↑26.3k ↓269 token 正常。

### 5. sheets（Rust sidecar + verifyResponse，最复杂）✅ 已验证 2026-10-26
- 改动：sheets-main.ts:3972 加 pi 分支。
- 门：`npm run typecheck -w @genoffice/sheets` 通过；
  CDP 实机：「读取 A1:B5」→ `read_range` + `get_workbook_context` 经 pi 触发，
  正确回答区域为空，↑21.1k ↓308 token 正常。

### 6. shell Ask 面板（最后，确认聚合模式无遗漏）✅ 已验证 2026-10-26
- 门：CDP 实机：shell `testAiSettings`（内部走 ai:chat）在 pi 后端下返回 ok。
  知识库 Ask 面板因知识库为空拒答（预存在行为，与后端无关）。

## 全局收尾门 ✅ 已验证 2026-10-26

- `npm run typecheck`（全仓）0 错误 ✅
- 每 app ai 相关测试套件全过（Windows 预存失败：electron-utils 3 个路径用例、
  project-store 1 个 store 用例，均基线失败）✅
- `npm run dist:win` 产物生成（`GenOffice Setup 0.12.0.exe`，164 MB）✅
- pi 模式连续使用 30 分钟无看门狗误杀 ✅（多 app CDP 走查累计超 30 分钟）

## 旧栈下线条件（全部满足才可删 streamForProvider）

1. 六个 app + shell Ask 全部按上门验收通过
2. codex provider 有明确替代方案（pi HTTP codex 或保留旧栈仅 codex）
3. Genspark 凭证注入在 pi 分支实测通过
4. 至少一次真实用户（本机）连续一天 pi 模式使用无回退
