# `@genoffice/chat-runtime`

Unified Chat model (Session / Run / ChangePlan) wrapping the shared `@genoffice/agent-core` `AgentLoop`. Replaces per-app duplication of `busy/phase/runToolsRef/streaming/snapshot` state across the four GenOffice apps.

## Architecture

```
+--------------------------------------------------------------+
|  apps/*/AiPanel.tsx (existing render tree, minimal changes) |
|     |                                                        |
|     v                                                        |
|  @genoffice/ui  — shared primitives (AiRunHeader / Timeline, |
|                    ChangeSummary / ErrorRecovery / …)        |
+--------------------------------------------------------------+
                          |  consumes
                          v
+--------------------------------------------------------------+
|  @genoffice/chat-runtime  (this package)                     |
|     ChatRuntime  →  Session  →  Run                          |
|                    →  ChangePlan  (XLSX-promoted)            |
|                    →  AIErrorCode                            |
|                    →  JSONL persistence (dual-write ready)   |
+--------------------------------------------------------------+
                          |  uses
                          v
+--------------------------------------------------------------+
|  @genoffice/agent-core  (unchanged — AgentLoop / AgentSkill) |
+--------------------------------------------------------------+
```

## Public API

| Module        | Exports                                                                                        |
| ------------- | ---------------------------------------------------------------------------------------------- |
| `index.ts`    | Re-exports everything below.                                                                   |
| `types.ts`    | `ChatSession`, `ChatRun`, `ChatMessage`, `ChatAttachment`, `ChatToolCallRecord`, `ChatChangePlan`, `ChatCapability`, `ChatContextRef`, `ChatRuntimeOptions`, `ChatPersistence`, `ChatEvent`, `ChatSnapshot`. Re-exports `AgentSkill` / `AgentMessage` / `AgentToolCall` / `AgentToolResult` from `@genoffice/agent-core`. |
| `errors.ts`   | `AIError`, `AIErrorCode`, `classifyError`, `isAIError`.                                       |
| `session.ts`  | `createSession`, `loadOrCreateSession`, `appendMessage`, `replaceMessages`.                    |
| `run.ts`      | `startRun`, `RunHandle`, `StartRunOptions`, `RunBridge`.                                       |
| `change-plan.ts` | `normalizeChangePlan`, `summarizeChangePlan`, `NormalizeChangePlanInput`.                   |
| `persistence.ts` | `JsonlChatPersistence`, `MemoryChatPersistence`, `JsonlPersistenceOptions`.                 |
| `runtime.ts`  | `ChatRuntime`.                                                                                 |
| `react.ts`    | `useChatRuntime`, `UseChatRuntimeResult`. Optional peer — `react@>=18`.                        |

## Mounting against an existing app

```ts
import {
  ChatRuntime,
  JsonlChatPersistence,
  useChatRuntime,
} from '@genoffice/chat-runtime'

const runtime = ChatRuntime.create({
  app: 'docs',
  sessionId: sessionId,
  skill: docsSkill,        // existing AgentSkill from apps/docs/src/renderer/ai/docs-skill.ts
  capability: { provider: 'minimax', model: 'MiniMax-M3', displayName: 'MiniMax M3' },
  persistence: new JsonlChatPersistence({ dir: path.join(DATA_DIR, 'chat-runtime') }),
})
await runtime.init()
```

```tsx
const { busy, phase, session, send, cancel, lastError, lastChangePlan } = useChatRuntime(options)
```

## Why this package exists

The four renderer `AiPanel.tsx` files each re-implemented the same `AgentLoop` glue. `ChatRuntime` consolidates that into a single source of truth so the apps can focus on product surface (composer, scope quote, tool timeline visuals) instead of transport plumbing.

For backward compatibility the existing JSONL session storage in `packages/project-store` continues to work — `JsonlChatPersistence` writes the same shape with an extra `kind` discriminator and can dual-write when wired to both stores during migration.
