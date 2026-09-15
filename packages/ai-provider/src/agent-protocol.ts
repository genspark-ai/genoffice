/**
 * Agent protocol types shared between the AI provider and the renderer side.
 *
 * Why this file exists (W22 deliverable — first step toward deleting agent-core):
 *   `@genoffice/ai-provider` historically declared its message / tool / image
 *   types (`AgentMessage` / `AgentToolCall` / `AgentToolDef` / `AgentImage`)
 *   via `import type { ... } from '@genoffice/agent-core'`. Those types are
 *   the wire shape between the renderer (apps/*) and the LLM provider
 *   layer — they have nothing to do with the agent-core ReAct loop that the
 *   pi-based migration replaces.
 *
 *   The plan (§1.2) marks `@genoffice/agent-core` for full deletion once
 *   nothing in the repo depends on it. This file lifts the four wire types
 *   into ai-provider itself so ai-provider no longer has to depend on
 *   agent-core for *type-only* imports.
 *
 * Stability: these types are part of the public surface of `@genoffice/ai-provider`
 * (`AiStreamRequest.messages`, `AiStreamChunk.toolCall`, etc.). They must
 * stay byte-compatible with whatever the renderer and pi-ai callers expect.
 */

export interface AgentToolDef {
  name: string
  description: string
  /** JSON Schema (object) describing the tool input */
  inputSchema: Record<string, unknown>
}

export interface AgentToolCall {
  id: string
  name: string
  input: Record<string, unknown>
  /** Parse error when the model emitted invalid input JSON; the loop feeds back an is_error result for retry instead of aborting the run */
  inputError?: string | undefined
  /** The argument stream was cut off by the token limit (stop_reason max_tokens); the loop asks the model to split the call instead of "fixing JSON" */
  truncated?: boolean | undefined
}

export interface AgentToolResult {
  id: string
  /** tool name (Gemini addresses function responses by name, not id) */
  name: string
  output: string
  isError?: boolean | undefined
}

/** inline image attached to a user turn, fed to vision-capable providers as multimodal input */
export interface AgentImage {
  /** raw base64 (no data: URL prefix) */
  base64: string
  /** e.g. "image/png" */
  mime: string
}

export type AgentMessage =
  | { role: 'user'; text: string; images?: AgentImage[] | undefined }
  /** reasoning: opaque vendor thinking captured during the turn; interleaved-thinking
   * models (e.g. MiniMax M3, DeepSeek V4) degrade in tool loops unless it is echoed back */
  | {
      role: 'assistant'
      text: string
      toolCalls?: AgentToolCall[] | undefined
      reasoning?: string | undefined
    }
  | { role: 'tool'; results: AgentToolResult[] }
