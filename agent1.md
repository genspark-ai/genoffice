# GenOffice Agent 重写计划 — 完全基于 `@earendil-works/pi-*`

> **目标**：把当前 GenOffice 自有的 `@genoffice/agent-core` + `@genoffice/ai-provider` + `@genoffice/chat-runtime` 整体替换为基于 [`pi.dev`](https://pi.dev/docs/latest) 的实现，**充分复用 pi 的插件机制**（Extensions / Skills / Prompts / Themes），**支持未来扩展**，打造**顶级 Office AI**。

---

## 0. 为什么选 pi — 学习后的真实判断

通过对 `/Users/louloulin/appx/pi` 源码 (84 个版本, 0.84.3) 和 [`pi.dev/docs/latest`](https://pi.dev/docs/latest) 的系统学习，发现 pi 已经为"嵌入第三方应用"准备了完整基础设施：

### 0.1 SDK 嵌入模式（推荐路径）

pi 官方文档 `packages/coding-agent/docs/sdk.md` 明确给出 SDK 嵌入模式：

```typescript
import { createAgentSession, ModelRuntime, SessionManager } from "@earendil-works/pi-coding-agent";

const modelRuntime = await ModelRuntime.create();
const { session } = await createAgentSession({
  sessionManager: SessionManager.inMemory(),
  modelRuntime,
});
session.subscribe((event) => { /* 推流给 UI */ });
await session.prompt("...");
```

四种运行模式：

| 模式 | 用途 | GenOffice 是否需要 |
| --- | --- | --- |
| **interactive (TUI)** | 终端交互式 | ❌ 不需要 |
| **print** | 非交互单次问答 | ❌ 不需要 |
| **JSON** | 自动化流水线 | △ 可选 (CLI 模式) |
| **RPC (JSON stdin/stdout)** | 进程嵌入 | ✅ 但不是首选 |
| **SDK** | **程序化嵌入** | ✅ **首选** |

**SDK 嵌入** 比 RPC 更直接：直接 import `createAgentSession`，不需额外进程。

### 0.2 插件体系 — pi 已经做好扩展点

| 机制 | 文件位置 | 作用 |
| --- | --- | --- |
| **Extensions** | `~/.pi/agent/extensions/*.ts` 或 `additionalExtensionPaths` | 自定义工具、UI、生命周期钩子 |
| **Skills** | `~/.pi/agent/skills/*.md` (frontmatter) | 技能 (markdown + frontmatter) |
| **Prompt Templates** | `~/.pi/agent/prompts/*.md` | 可复用 prompt 模板 |
| **Themes** | `~/.pi/agent/themes/*.json` | TUI 主题 |
| **Context Files** | 自动扫描项目根 | 项目级上下文 |
| **Agents (Subagents)** | `~/.pi/agent/agents/*.md` | 隔离上下文的子代理 |
| **Custom Providers** | `extensions/custom-provider-*` | OAuth / 自定义协议 |

### 0.3 pi 现有 84 个扩展示例可学

`packages/coding-agent/examples/extensions/` 下有 **80+ 个真实示例**：

| 类别 | 关键示例 | GenOffice 对应需求 |
| --- | --- | --- |
| Lifecycle & Safety | `permission-gate.ts`, `protected-paths.ts`, `dirty-repo-guard.ts` | docx 工具的危险操作 UI 确认 |
| Custom Tools | `hello.ts`, `todo.ts`, `tool-override.ts` | 把 22 个 docx 工具注册为 pi 工具 |
| Commands & UI | `handoff.ts`, `preset.ts`, `tools.ts`, `dynamic-tools.ts` | 动态工具注册（AI 翻译上下文） |
| Git Integration | `git-checkpoint.ts`, `auto-commit-on-exit.ts` | 文档版本快照 |
| System Prompt | `system-prompt-header.ts`, `pirate.ts` | 不同语言的 systemPrompt |
| Compaction | `custom-compaction.ts`, `trigger-compact.ts` | 翻译/长会话压缩 |
| Custom Providers | `custom-provider-anthropic/`, `custom-provider-gitlab-duo/` | Genspark/MiniMax 私有协议 |
| With Deps | `with-deps/` | Skills 带独立依赖 |
| Subagent | `subagent/` | 跨 Office 多 Agent 协作 |
| OAuth | `9-api-keys-and-oauth.ts` | 用户登录 Genspark 等 |

**结论**：每个 GenOffice 需求都已在 pi 有参考实现。

---

## 1. 目标架构 — 完全基于 pi 的扩展机制

### 1.1 进程拓扑

```
┌─────────────────────────────────────────────────────────────────────┐
│  Renderer (apps/docs/sheets/slides/web — React 18+)                  │
│                                                                       │
│  ┌─ AiPanel.tsx (React) ─────────────────────────────────────────┐ │
│  │  ├─ usePiSession() hook → @genoffice/agent-runtime          │ │
│  │  ├─ PiSessionProvider (Context)                              │ │
│  │  ├─ Office Extension (registers all docx/sheets/slides tools)│ │
│  │  ├─ <PiUIAdapter /> (适配 pi ExtensionUIContext → React)    │ │
│  │  └─ <HistoryPanel /> / <SkillsStore /> / <TokensBadge />    │ │
│  └──────────────────────────────────────────────────────────────┘ │
│                                                                       │
│  @genoffice/agent-runtime (薄壳层, ~300 行)                          │
│  ├─ createOfficeSession(options) → AgentSession                      │
│  ├─ PiUIReactAdapter: implements ExtensionUIContext                  │
│  ├─ PiSessionProvider (React Context for current session)            │
│  └─ usePiSession / usePiEvent hooks                                  │
│                                                                       │
│  ── 直接 import ────────────────────────────────────────────────────│
│  │                                                                   │
│  ▼                                                                   │
│  @earendil-works/pi-coding-agent (SDK)                                │
│  ├─ createAgentSession({ resourceLoader, sessionManager })          │
│  ├─ DefaultResourceLoader → 发现: extensions, skills, prompts, themes│
│  ├─ ModelRuntime.create()                                            │
│  └─ AgentSession: prompt / steer / followUp / subscribe / compact    │
│                                                                       │
│  @earendil-works/pi-agent-core                                        │
│  ├─ AgentLoop (ReAct + 防御)                                         │
│  ├─ AgentMessage / ToolCall / ToolResult                             │
│  ├─ beforeToolCall / afterToolCall / shouldStopAfterTurn              │
│  ├─ parallel + sequential 工具执行模式                                │
│  ├─ transformContext + convertToLlm 两阶段                            │
│  └─ EventStream (15+ 事件)                                            │
│                                                                       │
│  @earendil-works/pi-ai                                                │
│  ├─ 70+ providers (anthropic, openai, gemini, minimax, ...)          │
│  ├─ 11 个协议适配器 (anthropic-messages 1391 行, ...)                 │
│  ├─ OAuth 流程 (anthropic, openai-codex)                             │
│  ├─ Models / ModelRegistry / 自定义 Provider Factory                 │
│  └─ TypeBox JSON Schema 工具定义                                     │
│                                                                       │
│  @earendil-works/pi-telemetry                                         │
│  ├─ AI spans / Harness spans / 自定义 spans                          │
│  └─ Typed spans + 自带 OTel 兼容导出                                 │
│                                                                       │
│  @earendil-works/pi-session-backend-sqlite-node                       │
│  └─ SQLite 持久化 (Electron 主进程, web 用 IndexedDB shim)           │
└───────────────────────────────────────────────────────────────────────┘

┌─────────────────────────────────────────────────────────────────────┐
│  genoffice-skills/ (用户级和项目级 Skills, 在 `~/.pi/agent/`)       │
│  ├─ skills/docs-translator/SKILL.md          (frontmatter + body)    │
│  ├─ skills/legal-review/SKILL.md                                       │
│  ├─ skills/finance-summary/SKILL.md                                   │
│  └─ ...                                                              │
│                                                                       │
│  genoffice-extensions/ (GenOffice 提供的 Office 扩展)                │
│  ├─ extensions/docs-skill.ts             (22 docx 工具)              │
│  ├─ extensions/sheets-skill.ts           (Excel 工具)                │
│  ├─ extensions/slides-skill.ts           (PPT 工具)                   │
│  ├─ extensions/files-skill.ts            (附件读取)                   │
│  ├─ extensions/office-workflow.ts        (跨 Office 复合)            │
│  ├─ extensions/audit-log.ts              (审计钩子)                   │
│  ├─ extensions/agent-team.ts             (writer + reviewer + fact)  │
│  ├─ extensions/session-recovery.ts       (跨设备同步)                │
│  ├─ extensions/local-models.ts           (Ollama, Bedrock)           │
│  ├─ extensions/translation-providers.ts  (Genspark/MiniMax 私有协议) │
│  └─ extensions/permissions.ts             (replace_document UI 确认)  │
└───────────────────────────────────────────────────────────────────────┘
```

### 1.2 包布局 (GenOffice 侧)

| 新包 | 角色 | 行数估 |
| --- | --- | --- |
| `@genoffice/agent-runtime` | SDK 包装, React 集成 | ~800 |
| `@genoffice/agent-skills` | 把 22 个 docx 工具 + 新增 sheets/slides 工具, 封装为 pi 扩展文件 | ~1500 |
| `@genoffice/agent-ui` | PiUIReactAdapter + Dialog 组件 | ~600 |
| `@genoffice/agent-session` | SQLite (Electron) / IndexedDB (web) 持久化 | ~500 |
| `@genoffice/agent-telemetry` | 包装 pi-telemetry + span exporters | ~300 |
| **删除** `packages/agent-core` | 完全被 pi 替代 | -832 |
| **删除** `packages/chat-runtime` | 完全被 pi AgentSession + 自有 hook 替代 | -1332 |
| **删除** `packages/ai-provider` | 完全被 pi-ai 替代 | -6693 |

---

## 2. 核心集成策略 — Pi Extension API for Office

### 2.1 Extension 文件结构

每个 Office 工具集都是一个 **pi extension** 文件，遵循 pi 的 `ExtensionAPI` 契约：

```typescript
// genoffice-extensions/extensions/docs-skill.ts
import { Type, StringEnum } from "@earendil-works/pi-ai";
import type { ExtensionAPI } from "@earendil-works/pi-coding-agent";
import { defineTool } from "@earendil-works/pi-coding-agent";

const readBlocksTool = defineTool({
  name: "read_blocks",
  label: "Read Blocks",
  description: "Read a range of blocks from the document (0-indexed, inclusive).",
  parameters: Type.Object({
    startBlockIndex: Type.Integer({ minimum: 0 }),
    endBlockIndex: Type.Integer({ minimum: 0 }),
  }),
  promptGuidelines: [
    "Use this to inspect document content before edits.",
    "Image blocks return `[Protected content: Image, kept as is]` and cannot be modified.",
  ],
  async execute(_toolCallId, params, signal, onUpdate, ctx) {
    // 从 ctx.ui 拿当前文档编辑器 (通过 React UI 适配)
    const editor = ctx.ui.getCustomData<Editor>("editor");
    if (!editor) {
      return { content: [{ type: "text", text: "No editor available" }], details: {} };
    }
    // ... 实现同 GenOffice 当前 read_blocks ...
    return {
      content: [{ type: "text", text: JSON.stringify(blocks) }],
      details: { blockCount: blocks.length },
    };
  },
});

export default function docsSkillExtension(pi: ExtensionAPI) {
  // 注册 22 个工具
  pi.registerTool(readBlocksTool);
  pi.registerTool(replaceBlocksTool);
  pi.registerTool(insertContentTool);
  // ... 等等

  // 系统提示
  pi.on("before_agent_start", async (event) => {
    return {
      systemPromptAppend: "\n\n## Document Editing Rules\n[Office 特定规则]",
    };
  });

  // frozen selection: 锁定用户当前选择范围
  pi.on("session_start", async (event, ctx) => {
    const frozen = uiAdapter.captureSelection(ctx);
    ctx.ui.setCustomData("frozenSelection", frozen);
  });

  // replace_document 危险操作 UI 确认
  pi.on("tool_call", async (event, ctx) => {
    if (event.toolName === "replace_document") {
      const ok = await ctx.ui.confirm(
        "Replace entire document?",
        "This will overwrite the current document content. Undo is available.",
      );
      if (!ok) {
        return { block: true, reason: "User cancelled" };
      }
    }
    return undefined;
  });
}
```

### 2.2 UI 适配器 (PiExtensionUIContext → React)

pi 的 `ExtensionUIContext` 是抽象接口。我们实现 React 适配器：

```typescript
// packages/agent-runtime/src/ui-adapter.ts
import type { ExtensionUIContext, ExtensionUIDialogOptions } from "@earendil-works/pi-coding-agent";

export class ReactUIAdapter implements Partial<ExtensionUIContext> {
  constructor(
    private store: PiUIStore,        // React store (zustand 或类似)
    private signalChannel: WritableSignal<UIDialogRequest[]>,
  ) {}

  async select(title: string, options: string[], opts?: ExtensionUIDialogOptions): Promise<string | undefined> {
    return await this.store.openDialog({
      kind: "select",
      title,
      options,
      signal: opts?.signal,
      timeout: opts?.timeout,
    });
  }

  async confirm(title: string, message: string, opts?: ExtensionUIDialogOptions): Promise<boolean> {
    return await this.store.openDialog({
      kind: "confirm",
      title,
      message,
      signal: opts?.signal,
    });
  }

  notify(message: string, type: "info" | "warning" | "error" = "info"): void {
    this.store.pushNotification({ message, type });
  }

  setStatus(key: string, text: string | undefined): void {
    this.store.setStatus(key, text);
  }

  // GenOffice 扩展：把 Office 编辑器注入到 ctx
  setEditorInstance(editor: Editor): void {
    this.editorInstance = editor;
  }

  getEditorInstance(): Editor | undefined {
    return this.editorInstance;
  }
}
```

### 2.3 自定义 Provider (Genspark/MiniMax 私有协议)

参考 `examples/extensions/custom-provider-anthropic/` 写法：

```typescript
// genoffice-extensions/extensions/translation-providers.ts
import type { Provider, Api, Model, Context } from "@earendil-works/pi-ai";
import { defineProvider } from "@earendil-works/pi-coding-agent";

export const gensparkProvider = defineProvider<"genspark">({
  id: "genspark",
  baseUrl: "https://www.genspark.ai/api/anthropic",
  apiKey: () => process.env.GSK_API_KEY, // 由 electron-utils/credential-store 提供
  transformHeaders: (headers) => ({
    ...headers,
    "X-Agent-Type": "genoffice",  // Genspark 区分计费
  }),
  // 复用 pi-ai 的 anthropic-messages 协议,只换 baseUrl/headers
});
```

### 2.4 Provider 数量爆炸性增长

引入 pi 后, GenOffice 用户立刻可用 **70+ providers vs 当前 12**:

| 类型 | Pi 新增 | GenOffice 用户价值 |
| --- | --- | --- |
| 国际大牌 | openai-codex, opencode, openrouter, groq, xai, mistral, cerebras | 通用 |
| 国内 | minimax, moonshotai, zai, qwen-token-plan, kimi-coding, xiaomi, ant-ling | 中文场景必备 |
| 云 | amazon-bedrock, google-vertex, cloudflare-workers-ai, cloudflare-ai-gateway | 企业 |
| 本地 | huggingface, baseten, fireworks, together, vercel-ai-gateway, nvidia, radius | BYOK |
| 自定义 | `defineProvider()` 支持任意 OpenAI/Anthropic 兼容端点 | 私有部署 |

---

## 3. Skill / Prompt Template / Theme 体系

### 3.1 Skills (markdown + frontmatter)

```markdown
<!-- genoffice-skills/skills/quarterly-report/SKILL.md -->
---
name: quarterly-report
description: Generate a Q* quarterly report from spreadsheets and previous docs
tools:
  - read_blocks
  - read_attachment
  - insert_chart
  - insert_content
model: anthropic/claude-opus-4-7
---

# Quarterly Report Skill

You are an expert financial analyst. When the user asks for a quarterly report:

1. Read the attached spreadsheet for raw numbers
2. Read the previous quarter's docx for structure
3. Generate the new docx following the [template in prompts/quarterly-report.md]
4. Insert a chart for each KPI

## Rules
- Always cite the source row/column for any number
- ...
```

### 3.2 Prompt Templates

```markdown
<!-- ~/.pi/agent/prompts/quarterly-report.md -->
<!-- @genoffice/template:quarterly-report -->

# Quarterly Report Template

[结构化模板, 支持变量替换]
```

### 3.3 Themes (GenOffice 主题适配)

```json
<!-- ~/.pi/agent/themes/genoffice-light.json -->
{
  "name": "GenOffice Light",
  "colors": { ... },
  "extensions": ["docs", "sheets", "slides"]
}
```

(注: pi themes 是 TUI 概念, GenOffice 在 React 渲染, 主要复用 Skills/Prompts)

---

## 4. 分阶段实施 — 5 阶段,约 14-18 周

### Phase 1: 接入与基础 (Week 1-2)

**学习目标**: 跑通 pi SDK 的最小 demo

**任务**:

1.1. **链接 pi 包**:
```json
// apps/docs/package.json
{
  "dependencies": {
    "@earendil-works/pi-coding-agent": "file:../../../pi/packages/coding-agent",
    "@earendil-works/pi-agent-core": "file:../../../pi/packages/agent",
    "@earendil-works/pi-ai": "file:../../../pi/packages/ai",
    "@earendil-works/pi-telemetry": "file:../../../pi/packages/telemetry"
  }
}
```

1.2. **写最小 demo** `apps/docs/src/renderer/ai/pi-demo.tsx`:
```typescript
import { createAgentSession, ModelRuntime, SessionManager } from "@earendil-works/pi-coding-agent";

export async function runPiDemo() {
  const modelRuntime = await ModelRuntime.create();
  const { session } = await createAgentSession({
    sessionManager: SessionManager.inMemory(),
    modelRuntime,
  });

  session.subscribe((event) => {
    if (event.type === "message_update") { /* 推到 UI */ }
  });

  await session.prompt("Hello, what can you do?");
  session.dispose();
}
```

1.3. **解决 tsconfig 兼容**: pi 用 `typebox`, GenOffice 用 `zod`, 互不影响但需统一定义边界

1.4. **验证现有测试零回归**

**交付**: pi 包接入工作区, demo 可跑通 "Hello"

---

### Phase 2: UI 适配 + 第一个工具 (Week 3-5)

**学习目标**: 把 pi 的 `ExtensionUIContext` 用 React 实现, 注册第一个 docx 工具

**任务**:

2.1. **包 `@genoffice/agent-runtime`** — 薄壳:
   - `createOfficeSession({ cwd, model, sessionManager })` 包装 `createAgentSession`
   - `ReactUIAdapter implements ExtensionUIContext`
   - `PiSessionProvider` (React Context) 暴露 `usePiSession()`

2.2. **包 `@genoffice/agent-skills`** 骨架:
   - 文件结构: `src/extensions/docs-skill.ts` 等
   - `defineTool()` 包装,把 GenOffice 当前的工具迁移

2.3. **第一个工具迁移** `read_blocks` (选最简单的):
   ```typescript
   // genoffice-extensions/extensions/docs-skill.ts
   import { Type } from "@earendil-works/pi-ai";
   import { defineTool } from "@earendil-works/pi-coding-agent";

   const readBlocksTool = defineTool({
     name: "read_blocks",
     description: "Read a range of blocks from the document",
     parameters: Type.Object({
       startBlockIndex: Type.Integer({ minimum: 0 }),
       endBlockIndex: Type.Integer({ minimum: 0 }),
     }),
     promptGuidelines: ["Image blocks return as [Protected content: Image]"],
     async execute(toolCallId, params, signal, onUpdate, ctx) {
       const adapter = ctx.ui as ReactUIAdapter;
       const editor = adapter.getEditorInstance();
       if (!editor) {
         return { content: [{ type: "text", text: "No editor" }], details: {} };
       }
       const blocks = readBlocks(editor, params.startBlockIndex, params.endBlockIndex);
       return {
         content: [{ type: "text", text: JSON.stringify(blocks) }],
         details: { blockCount: blocks.length },
       };
     },
   });

   export default function docsSkillExtension(pi: ExtensionAPI) {
     pi.registerTool(readBlocksTool);
   }
   ```

2.4. **PiSessionProvider + usePiSession**:
   ```typescript
   // packages/agent-runtime/src/provider.tsx
   const PiSessionContext = createContext<PiSession | null>(null);

   export function PiSessionProvider({ cwd, children }) {
     const [session, setSession] = useState<PiSession | null>(null);
     useEffect(() => {
       const runtime = createOfficeSession({ cwd });
       setSession(runtime);
       return () => runtime.dispose();
     }, [cwd]);
     return <PiSessionContext.Provider value={session}>{children}</PiSessionContext.Provider>;
   }

   export function usePiSession() {
     const ctx = useContext(PiSessionContext);
     if (!ctx) throw new Error("usePiSession outside provider");
     return ctx;
   }
   ```

2.5. **验证**:
   - 浏览器实测: AI 调用 read_blocks, 正确返回文档块
   - React UI dialog (confirm / select) 正确弹出

**交付**: 第一个 docx 工具通过 pi SDK 工作

---

### Phase 3: 全部 22 个工具迁移 (Week 6-9)

**学习目标**: 把所有 GenOffice docx/sheets/slides 工具迁到 pi extensions

**任务**:

3.1. **包 `@genoffice/agent-skills`** 完整版:

| Extension 文件 | 工具数 | 来源 |
| --- | --- | --- |
| `extensions/docs-skill.ts` | 22 | `apps/docs/src/renderer/ai/tools.ts` |
| `extensions/sheets-skill.ts` | (新) | GenOffice 当前 sheets AI 工具 |
| `extensions/slides-skill.ts` | (新) | GenOffice 当前 slides AI 工具 |
| `extensions/files-skill.ts` | 1 | 当前 `files-skill.ts` |

3.2. **`frozenSelection` 包装** (GenOffice 特有):
   ```typescript
   pi.on("session_start", async (event, ctx) => {
     const adapter = ctx.ui as ReactUIAdapter;
     const editor = adapter.getEditorInstance();
     if (editor) {
       const frozen = { scope: getSelectionScope(editor), doc: editor.state.doc };
       adapter.setCustomData("frozenSelection", frozen);
     }
   });
   ```

3.3. **`verifyResponse` 包装** (GenOffice 声明-行动一致性校验):
   - 利用 pi 的 `transformContext` 钩子, 在每轮结束前注入校验提示
   ```typescript
   pi.on("before_agent_start", async (event, ctx) => {
     const verify = verifyResponseFn();
     return { systemPromptAppend: `[Verify rules]\n${verify}` };
   });
   ```

3.4. **特殊工具**: `replace_document` 等需要 UI 确认:
   ```typescript
   pi.on("tool_call", async (event, ctx) => {
     if (event.toolName === "replace_document") {
       const ok = await ctx.ui.confirm("Replace entire document?", "Undo available");
       if (!ok) return { block: true, reason: "User cancelled" };
     }
   });
   ```

3.5. **AiPanel.tsx 完全重写** 用 pi EventStream:
   ```typescript
   // 旧 (GenOffice):
   agent.subscribe((event) => { /* 6 callback */ });

   // 新 (pi):
   session.subscribe((event) => {
     if (event.type === "message_update") { /* text delta */ }
     else if (event.type === "tool_execution_start") { /* spinner */ }
     else if (event.type === "tool_execution_end") { /* tool card */ }
     // 15+ 事件类型
   });
   ```

3.6. **翻译-core 重写** 用 pi-ai 替换 chatForProvider:
   - `apps/web-server/src/ai/chat.ts` 中 `ai:translate` 用 `pi.completeSimple()` 替换
   - sharedMemory (translation-core) 保留

3.7. **验证**:
   - `curl POST /api/ipc/ai:translate` 通过 pi-ai 后端返回结果
   - 浏览器实测 22 个工具全部能调用

**交付**: 所有 Office 工具通过 pi 工作, AI 翻译用 pi 后端

---

### Phase 4: 持久化 + Telemetry + OAuth (Week 10-12)

**学习目标**: 用 pi 的会话后端和 telemetry, 实现生产级特性

**任务**:

4.1. **会话持久化** (Electron):
   ```typescript
   import { SessionManager } from "@earendil-works/pi-coding-agent";
   const sm = SessionManager.create(userHomeDir);
   const { session } = await createAgentSession({ sessionManager: sm });
   ```
   - 关闭 Electron 再打开,会话恢复
   - 支持分支 (`/tree` 浏览历史)

4.2. **会话持久化** (Web):
   - IndexedDB 后端 (`packages/agent-session/src/indexeddb-backend.ts`)
   - 兼容 JSONL 格式,与 Electron 互导

4.3. **Telemetry**:
   ```typescript
   import { startAiSpan, startHarnessSpan } from "@earendil-works/pi-telemetry";

   pi.on("agent_start", (event) => {
     startHarnessSpan({ name: "office-ai.run", attributes: { app: "docs" } });
   });
   ```
   - 输出到本地文件 (`~/.genoffice/ai-traces.jsonl`)
   - 可选 OTel exporter (企业用户)

4.4. **OAuth** (Genspark/Codex 等):
   - 复用 pi 的 OAuth 流程 (`9-api-keys-and-oauth.ts`)
   - 浏览器端用 redirect flow, 桌面端用 device flow

4.5. **历史 UI**:
   - `<HistoryPanel />` (左侧抽屉)
   - 搜索/标签/导出
   - 复用 pi 的 session tree 概念

4.6. **验证**:
   - 重启 Electron 会话恢复
   - OTel collector 看到完整 spans

**交付**: 持久化 + 完整可观测性

---

### Phase 5: 高级特性 (Week 13-18)

**学习 pi 的 subagent / handoff / preset 等模式**

**任务**:

5.1. **跨 Office 工作流** `extensions/office-workflow.ts`:
   ```typescript
   // "AI 帮我做季度报告" — 一次提示跨 sheets + docs + slides
   export default function officeWorkflowExtension(pi: ExtensionAPI) {
     pi.registerTool({
       name: "cross_office_workflow",
       description: "Compose a workflow across multiple Office apps",
       parameters: Type.Object({
         spreadsheetPath: Type.String(),
         templateDocPath: Type.String(),
         outputFormat: StringEnum(["docx", "slides"] as const),
       }),
       async execute(...) {
         // 用 subagent 模式调用 sheets-skill 和 docs-skill
       },
     });
   }
   ```

5.2. **多 Agent 团队** `extensions/agent-team.ts`:
   ```typescript
   // 参考 pi 的 subagent/ 示例
   // writer: 主对话
   // reviewer: 静默评估 (reviewer-skill 自动触发)
   // fact-checker: 数字准确性
   ```

5.3. **企业级审计** `extensions/audit-log.ts`:
   ```typescript
   pi.on("tool_call", async (event, ctx) => {
     auditLog.append({ tool: event.toolName, input: event.input, user: ctx.user, timestamp: Date.now() });
   });
   pi.on("tool_result", async (event, ctx) => {
     auditLog.complete(event.toolCallId, event.result);
   });
   ```

5.4. **本地模型** `extensions/local-models.ts`:
   ```typescript
   // Ollama / Bedrock / Vertex / Cloudflare Workers AI
   export const ollamaProvider = defineProvider({
     id: "ollama",
     baseUrl: "http://localhost:11434/v1",
     api: "openai-completions",  // Ollama 兼容 OpenAI
   });
   ```

5.5. **Skills 市场原型**:
   - 利用 pi 的 Skills 发现机制
   - GenOffice 提供 `genoffice skill install <name>` CLI 命令 (已有)
   - Skills store UI (在 settings 页面)

5.6. **性能优化**:
   - 工具并行 (parallel mode 默认开启)
   - Provider 响应缓存 (相同请求短窗口去重)

**交付**: 顶级 Office AI 全部能力

---

## 5. 文件级变更详细清单

### 5.1 新建

| 路径 | 用途 |
| --- | --- |
| `packages/agent-runtime/package.json` | 薄壳,包装 pi SDK |
| `packages/agent-runtime/src/create-session.ts` | `createOfficeSession()` |
| `packages/agent-runtime/src/ui-adapter.ts` | `ReactUIAdapter` |
| `packages/agent-runtime/src/react.tsx` | `PiSessionProvider`, `usePiSession`, `usePiEvent` |
| `packages/agent-runtime/src/types.ts` | GenOffice-specific 类型 |
| `packages/agent-skills/package.json` | Office Skills 包 |
| `packages/agent-skills/src/index.ts` | 聚合所有 extensions |
| `packages/agent-skills/src/extensions/docs-skill.ts` | 22 个 docx 工具 |
| `packages/agent-skills/src/extensions/sheets-skill.ts` | Excel 工具 |
| `packages/agent-skills/src/extensions/slides-skill.ts` | PPT 工具 |
| `packages/agent-skills/src/extensions/files-skill.ts` | 附件读取 |
| `packages/agent-skills/src/extensions/frozen-selection.ts` | 锁定选择 |
| `packages/agent-skills/src/extensions/verify-response.ts` | 声明-行动校验 |
| `packages/agent-skills/src/extensions/permissions.ts` | UI 确认 |
| `packages/agent-session/package.json` | 会话持久化 |
| `packages/agent-session/src/sqlite.ts` | Electron SQLite |
| `packages/agent-session/src/indexeddb.ts` | Web IndexedDB |
| `packages/agent-session/src/jsonl.ts` | 通用导入导出 |
| `packages/agent-telemetry/package.json` | 遥测导出 |
| `packages/agent-telemetry/src/exporter.ts` | OTel + 本地文件 |
| `genoffice-extensions/extensions/office-workflow.ts` | 跨 Office |
| `genoffice-extensions/extensions/agent-team.ts` | 多 Agent |
| `genoffice-extensions/extensions/audit-log.ts` | 审计 |
| `genoffice-extensions/extensions/local-models.ts` | 本地模型 |
| `genoffice-extensions/extensions/translation-providers.ts` | 私有协议 |
| `genoffice-skills/skills/*/SKILL.md` | 用户级 Skills |
| `genoffice-extensions/prompts/*.md` | Prompt 模板 |
| `docs/architecture/agent-v2.md` | 架构文档 |
| `docs/migration/agent-v2.md` | 迁移指南 |

### 5.2 修改

| 文件 | 改什么 |
| --- | --- |
| `apps/docs/src/renderer/ai/AiPanel.tsx` | 用 `usePiSession` 替换当前 6 callback |
| `apps/docs/src/renderer/ai/transports.ts` | 删除 (用 pi SDK) |
| `apps/docs/src/renderer/ai/docs-skill.ts` | 删除 (迁到 packages/agent-skills) |
| `apps/docs/src/renderer/ai/files-skill.ts` | 删除 |
| `apps/docs/src/renderer/ai/protocol.ts` | 拆分为工具定义 + systemPrompt |
| `apps/docs/src/renderer/ai/tools.ts` | 拆分为 22 个独立 defineTool |
| `apps/docs/src/main/docs-main.ts` | 加 SQLite session 后端 |
| `apps/web-server/src/ai/chat.ts` | 用 pi-ai 替换 chatForProvider |
| `apps/web-server/src/ai/pi-bridge.ts` | pi provider → GenOffice 适配 |
| `apps/web-server/src/ai/session-store.ts` | IndexedDB shim |
| `apps/web-server/package.json` | 加 pi 依赖 |
| `apps/docs/package.json` | 同上 |
| `packages/ui/src/ai-runtime.css` | 新 UI 样式 |
| `package.json` (根) | 加 pi workspace link |

### 5.3 删除 (Phase 3 完成后)

| 文件 | 行数 | 替代 |
| --- | --- | --- |
| `packages/agent-core/src/loop.ts` | 832 | pi `AgentLoop` |
| `packages/agent-core/src/electron-transport.ts` | - | pi SDK (直接 import) |
| `packages/agent-core/src/http-transport.ts` | - | 同上 |
| `packages/agent-core/src/web-transport.ts` | - | 同上 |
| `packages/agent-core/src/stream-text.ts` | - | pi EventStream |
| `packages/chat-runtime/src/*.ts` | 1332 | pi AgentSession + React hook |
| `packages/ai-provider/src/protocols/*.ts` | 970 | pi-ai |
| `packages/ai-provider/src/chat.ts` | - | pi-ai |
| `apps/web-server/src/ai/index.ts` (部分) | - | pi-bridge |

### 5.4 保留不变

| 文件 | 原因 |
| --- | --- |
| `packages/translation-core/*` | 翻译逻辑独立,只换 provider |
| `packages/docx-engine/*` | 文档引擎, 与 AI 无关 |
| `packages/sheets/*`, `slides/*`, `pdf/*`, `markdown/*`, `html/*` | 各 app 引擎 |
| `packages/ipc-bridge/*` | IPC 基础设施 |

---

## 6. 关键技术决策 (ADR 风格)

### 6.1 嵌入方式: SDK > RPC

| 维度 | SDK (推荐) | RPC |
| --- | --- | --- |
| 进程数 | 0 额外 (in-process) | +1 sidecar |
| 启动延迟 | < 100ms | ~500ms |
| 双向通信 | 函数调用 | JSON stdin/stdout |
| 状态共享 | 直接访问 | 序列化/反序列化 |
| 调试难度 | 低 (同一进程) | 高 (跨进程) |

**选择 SDK**,仅 CLI 模式 (print/json) 才考虑 RPC。

### 6.2 Extension 发现: 项目级 + 用户级

参考 pi 的 `DefaultResourceLoader`:
```typescript
const loader = new DefaultResourceLoader({
  cwd: process.cwd(),
  agentDir: getAgentDir(),
  extensionFactories: [
    docsSkillExtension,
    sheetsSkillExtension,
    // ...
  ],
});
```

不依赖文件发现, 直接 `extensionFactories` 注入 (避免 dynamic import 在打包器下出问题)。

### 6.3 UI 适配: React Signals

不引入 zustand/jotai, 用 React 18 的 `useSyncExternalStore` + 一个最小信号库 (~50 行):

```typescript
class PiUIStore {
  dialogs: UIDialogRequest[] = [];
  notifications: UINotification[] = [];
  statuses = new Map<string, string>();

  // React 订阅
  listeners = new Set<() => void>();
  subscribe = (cb: () => void) => { this.listeners.add(cb); return () => this.listeners.delete(cb); };
  notify = () => this.listeners.forEach(l => l());
}
```

### 6.4 Provider 自定义: 包装 pi-ai 的 `defineProvider`

参考 `custom-provider-anthropic/` 和 `custom-provider-gitlab-duo/`。

### 6.5 Skills 加载: 启动时一次性读取

不热加载 (Office app 启动慢, 一次性扫描所有 .md 文件, 编译为内存对象)。

### 6.6 类型兼容: typebox vs zod

- pi 用 typebox (Static<TSchema>)
- GenOffice 用 zod (z.infer<...>)
- 互不影响,各自保留

工具定义用 pi 的 Type.Object,GenOffice 侧不重复定义。

---

## 7. 风险与缓解

| 风险 | 影响 | 缓解 |
| --- | --- | --- |
| pi 包 TypeScript 兼容性 | 中 | 本地 link, patch pi tsconfig |
| pi 0.x API 变动 | 高 | 本地 link 锁定 + 抽象层 `@genoffice/agent-runtime` 缓冲 |
| 22 个 docx 工具迁移遗漏 | 高 | 完整迁移 + e2e 测试覆盖每个工具 |
| frozen selection / verifyResponse 语义丢失 | 高 | 专门写测试用例, 行为对比 |
| GenOffice 用户配置 (AI settings) 不兼容 | 中 | 写转换层, 旧 settings.json 仍可读 |
| OAuth 浏览器端限制 | 中 | 仅 Electron 桌面支持 OAuth, web 用 BYOK API key |
| 包体积膨胀 | 低 | pi 的 tree-shaking 已优化, 只 import 实际用到的 provider |
| 学习曲线 | 中 | 团队分阶段熟悉, 参考 pi 现有 80+ 示例 |
| Skills 市场冷启动 | 低 | 先内部用, 公开发布推迟 |

---

## 8. 测试与验证

### 8.1 单元测试

| 包 | 测试数目标 |
| --- | --- |
| `@genoffice/agent-runtime` | ≥ 30 (React hooks, UI 适配器) |
| `@genoffice/agent-skills` | ≥ 50 (每个工具至少 1 个测试) |
| `@genoffice/agent-session` | ≥ 20 (SQLite, IndexedDB 适配) |
| `@genoffice/agent-telemetry` | ≥ 10 (spans 正确性) |
| **保留** `packages/translation-core/` | 47 个 |

### 8.2 集成测试

| 场景 | 验证 |
| --- | --- |
| pi provider 端到端 | 10 个代表性 provider (anthropic/openai/gemini/minimax/...) |
| Agent 循环 | 3 个复杂 Office 任务 |
| 并行工具 | 翻译提速 ≥ 30% |
| 会话持久化 | 创建 → 关闭 → 重开 → 验证 |
| Telemetry spans | mock OTel collector |

### 8.3 E2E 测试 (Playwright)

| 场景 | 验证 |
| --- | --- |
| AI 翻译 45 页 docx | 通过 pi 后端, 内容翻译正确 |
| 22 个工具各跑 1 次 | 每个工具至少 1 次 e2e |
| Sheets: 跨表公式生成 | AI 自动写公式 |
| Slides: 10 页大纲+展开 | AI 自主规划 |
| 跨 Office 工作流 | Excel 数据 → docx 图表 |
| 多 Agent 团队 | writer + reviewer 互相修正 |
| OAuth 登录 Genspark | 完整流程 |

### 8.4 性能基准

| 指标 | 当前 | 目标 |
| --- | --- | --- |
| 翻译 45 页 docx | ~3 分钟 | < 1.5 分钟 |
| 工具调用失败率 | 5% | < 2% |
| 长会话 (50 轮) 延迟 | 8 秒 | < 3 秒 |
| Provider 切换 | 重启 | < 1 秒 |
| App 启动时间 | - | +200ms (pi 加载) |

---

## 9. 成功标准

### 9.1 Phase 1-2 (Week 5)
- [ ] pi 包接入工作区, demo 跑通
- [ ] ReactUIAdapter 实现完成, 第一个 dialog 工作
- [ ] read_blocks 工具通过 pi SDK 调用
- [ ] 现有 47 个 translation-core 测试零回归

### 9.2 Phase 3 (Week 9)
- [ ] 22 个 docx 工具 + 新增 sheets/slides 工具全部通过 pi SDK 工作
- [ ] AiPanel 用 pi EventStream, frozenSelection + verifyResponse 行为保留
- [ ] `curl POST /api/ipc/ai:translate` 通过 pi-ai 后端

### 9.3 Phase 4 (Week 12)
- [ ] SQLite session 持久化, 重启恢复
- [ ] Telemetry spans 完整
- [ ] OAuth 登录 Genspark/Codex 正常
- [ ] 历史会话 UI 可用

### 9.4 Phase 5 (Week 18) — 顶级 Office AI 终态
- [ ] **跨 Office 工作流**: "AI 帮我做季度报告" 一句话生成 Excel + docx + slides
- [ ] **多 Agent 团队**: writer + reviewer + fact-checker 自主协作
- [ ] **企业级审计**: replace_document 强制 UI 确认, 完整审计日志
- [ ] **70+ providers**: 用户立即可用, 无需配置
- [ ] **本地模型**: Ollama / Bedrock / Vertex / Cloudflare 可切换
- [ ] **Skills 市场**: 第三方 Skills 可发布 (npm/git)
- [ ] **性能基准全部达标**

### 9.5 顶级 Office AI 标志 (End State)

1. **"AI 帮我做季度报告"** 一句话生成 Excel + docx + slides 三件套
2. **AI 团队**: writer + reviewer + fact-checker 自主协作
3. **企业级**: 完整审计日志 + 危险操作 UI 确认 + 合规报告导出
4. **本地可用**: Ollama / Bedrock / Vertex / Cloudflare 任意切换
5. **生态**: 第三方 Skills 可发布 (npm/git)
6. **可观测**: 完整 telemetry + span trace
7. **会话持久**: 跨设备同步会话
8. **70+ providers**: 任何模型任意切换

---

## 10. 立即可做的第一周任务 (Phase 1)

### 10.1 在 `package.json` 加 pi 依赖

```json
// apps/docs/package.json
{
  "dependencies": {
    "@earendil-works/pi-coding-agent": "file:../../../pi/packages/coding-agent",
    "@earendil-works/pi-agent-core": "file:../../../pi/packages/agent",
    "@earendil-works/pi-ai": "file:../../../pi/packages/ai",
    "@earendil-works/pi-telemetry": "file:../../../pi/packages/telemetry"
  }
}
```

### 10.2 写最小 smoke test

```ts
// apps/docs/src/renderer/ai/pi-smoke.ts
import { createAgentSession, ModelRuntime, SessionManager } from "@earendil-works/pi-coding-agent";

const modelRuntime = await ModelRuntime.create();
const { session } = await createAgentSession({
  sessionManager: SessionManager.inMemory(),
  modelRuntime,
});

const events: string[] = [];
session.subscribe((event) => {
  events.push(event.type);
});

await session.prompt("Say hi in 5 words.");
console.log("Events:", events);
session.dispose();
```

### 10.3 跑

```bash
cd /Users/louloulin/appx/genoffice/apps/docs && npx tsx src/renderer/ai/pi-smoke.ts
```

看到 `Events: [message_start, message_update, ..., agent_end]` 即成功。

---

## 11. 长期愿景 — 顶级 Office AI

**短期 (Phase 1-4, 12 周)**: 把 GenOffice 的 AI 体验拉到 pi 用户的水准

**中期 (Phase 5 + 6 个月)**: 成为"Office AI 的 pi":
- 跨 Office 工作流
- 多 Agent 团队
- 企业级审计
- Skills 市场
- 跨设备会话同步

**长期 (1 年+)**: **AI-native Office 套件**:
- 用户打开 App 看到的不是工具栏, 而是 AI 团队
- 所有 Office 操作都是"AI 帮我..."
- 第三方 Skills 生态
- OpenTelemetry 标准输出
- 本地 + 云端模型无缝切换
- 多模态实时流 (语音/截图/生成图表)

**这是 pi-mono 在 Office 域的对应物**。

---

## 12. 文档清单

Phase 1-3 完成后需补:
- `docs/architecture/agent-v2.md` — 新架构图
- `docs/migration/agent-v2.md` — 用户视角的迁移指南
- `packages/agent-runtime/README.md` — 公共 API + 如何嵌入
- `packages/agent-skills/README.md` — 如何写 Skill/Extension
- `packages/agent-skills/docs/writing-tools.md` — 如何写 defineTool
- `packages/agent-skills/docs/writing-extensions.md` — 完整 Extension 教程
- `docs/skills/` — 用户级 Skills 编写指南
- `CONTRIBUTING.md` — 如何贡献 Skills

---

## 13. 决策记录

| 决策 | 备选 | 理由 |
| --- | --- | --- |
| **SDK 嵌入** (in-process) | RPC (sidecar) | 启动快, 调试易, 状态共享 |
| **本地 link pi 包** | npm publish | pi 0.x 频繁迭代 |
| **双轨过渡** (新旧共存 12 周) | 一次性切换 | 风险控制 |
| **只引 pi-ai + pi-coding-agent + pi-agent-core + pi-telemetry** | 引 pi-coding-agent 全部 | GenOffice 不需要 TUI/CLI |
| **React Signals UI 适配** (不引 zustand) | zustand/jotai | 减小依赖 |
| **extensionFactories 注入** | 文件系统发现 | 避免打包器 dynamic import 问题 |
| **SQLite (Electron) + IndexedDB (Web)** | 统一 SQLite | 浏览器沙箱限制 |
| **Skills 启动时一次读取** | 热加载 | Office app 启动慢, 一次性更简单 |
| **OAuth 仅 Electron 桌面** | 浏览器也支持 | 浏览器 OAuth 流程太复杂 |
| **Phase 5 才做多 Agent** | Phase 3 就做 | 单 Agent 成熟后再叠加 |
| **typebox (pi) + zod (GenOffice) 双轨** | 统一 typebox | 互不影响, 各取所长 |

---

## 14. 执行检查清单 (Tracking)

每周更新,完成打勾:

```
W1  [ ] pi 包接入 apps/docs
    [ ] pi-smoke.ts 跑通 "Hello"
W2  [ ] 解决 tsconfig 兼容
    [ ] 现有测试零回归验证

W3  [ ] @genoffice/agent-runtime 骨架
    [ ] ReactUIAdapter 完成
    [ ] PiSessionProvider/usePiSession
W4  [ ] read_blocks 工具迁移
    [ ] 第一个 dialog 工作
W5  [ ] e2e: 浏览器实测 read_blocks

W6  [ ] docs-skill.ts 全部 22 个工具迁完
W7  [ ] sheets-skill.ts + slides-skill.ts
W8  [ ] AiPanel.tsx 用 pi EventStream
W9  [ ] frozenSelection + verifyResponse 包装
    [ ] translation-core 切换到 pi-ai

W10 [ ] SQLite session backend
W11 [ ] IndexedDB session backend
W12 [ ] telemetry spans + OTel exporter

W13 [ ] 跨 Office 工作流扩展
W14 [ ] 多 Agent 团队扩展
W15 [ ] 企业级审计扩展
W16 [ ] 本地模型 (Ollama)
W17 [ ] Skills 市场原型
W18 [ ] 性能基准达标
```

---

## 15. 参考资料

- **Pi 官方文档**: <https://pi.dev/docs/latest>
- **Pi SDK 嵌入指南**: `packages/coding-agent/docs/sdk.md`
- **Pi 84 个扩展示例**: `packages/coding-agent/examples/extensions/`
- **Pi 13 个 SDK 示例**: `packages/coding-agent/examples/sdk/01..13-*.ts`
- **关键示例 (本计划直接参考)**:
  - `06-extensions.ts` - 自定义扩展
  - `05-tools.ts` - 自定义工具
  - `04-skills.ts` - Skills 加载
  - `13-session-runtime.ts` - 会话运行时
  - `extensions/permission-gate.ts` - 工具调用 UI 确认
  - `extensions/handoff.ts` - 会话转移
  - `extensions/dynamic-tools.ts` - 运行时注册工具
  - `extensions/subagent/` - 多 Agent 协作
  - `extensions/custom-provider-anthropic/` - 自定义 Provider
  - `extensions/git-checkpoint.ts` - 文档快照

**这个计划基于对 pi 源码和文档的深入学习，充分复用 pi 的扩展机制，保留 GenOffice 的 Office 编辑特殊性，为顶级 Office AI 留好接口。**
