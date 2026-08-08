/** JSON-Schema described tool exposed to the model */
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
  /**
   * Gemini's per-call thinking signature, carried so it can be echoed back on
   * the next turn. Gemini rejects a follow-up whose function calls have lost it
   * (`HTTP 400 ... missing a thought_signature in functionCall parts`), which
   * only bites from the second tool round onwards. Opaque: never inspected,
   * only stored and handed back, and absent for every other provider.
   */
  thoughtSignature?: string | undefined
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
  | {
      role: 'user'
      text: string
      images?: AgentImage[] | undefined
      /**
       * This turn only exists to carry images a tool produced (see
       * `ToolExecution.images`), not anything the user attached. Compaction
       * uses it to reclaim stale renders, which the user's own attachments
       * must keep.
       */
      fromTool?: boolean | undefined
    }
  | { role: 'assistant'; text: string; toolCalls?: AgentToolCall[] | undefined }
  | { role: 'tool'; results: AgentToolResult[] }

/**
 * Side-channel display data: UI-only, never merged into messages sent to the LLM.
 * kind='images' → image grid; kind='links' → link list; kind='text' → extra text.
 */
export interface ToolDisplay {
  kind: 'images' | 'links' | 'text'
  /** entry list for images / links modes */
  items?: Array<{ url: string; title?: string; thumb?: string }>
  /** extra text for text mode */
  text?: string
}

/** outcome of one tool execution */
export interface ToolExecution {
  /** result text fed back to the model */
  output: string
  isError?: boolean
  /** true when the tool changed the underlying artifact (document / sheet / deck) */
  mutated?: boolean
  /** short human-readable label for activity UI */
  summary: string
  /**
   * Side-channel display: for UI only, never enters the LLM context.
   * Ignored when tool results are assembled into an AgentMessage.
   */
  display?: ToolDisplay
  /**
   * Images the tool produced for the model to look at — a rendered page, a
   * slide as it will actually print.
   *
   * They cannot travel in `output`, because a tool result is text on every
   * provider we support: Anthropic allows image blocks inside a tool_result,
   * the OpenAI-compatible `tool` message is a plain string, and Gemini's
   * functionResponse carries a JSON struct. Only a *user* turn accepts images
   * on all three, so the loop appends one after the tool results rather than
   * trying to smuggle them into the result itself.
   */
  images?: AgentImage[]
}

// ---- run phase (drives the in-progress status line in chat UIs) ----

export type AgentPhaseKind =
  /** request sent, waiting for the model's first content block */
  | 'requesting'
  | 'thinking'
  | 'responding'
  /** the model is streaming tool arguments (e.g. a full outline) with no visible text */
  | 'tool-input'
  | 'tool-running'

export interface AgentPhase {
  kind: AgentPhaseKind
  toolName?: string | undefined
}

// ---- LLM transport (how one model turn is streamed; app supplies the impl) ----

export interface AgentStreamRequest {
  system: string
  messages: AgentMessage[]
  tools: AgentToolDef[]
}

export interface AgentStreamCallbacks {
  onDelta(text: string): void
  /** complete parsed tool call (arguments finished streaming) */
  onToolCall(call: AgentToolCall): void
  /** Phase changes within the model stream (thinking / responding / tool-input); older transports may omit this */
  onPhase?(phase: AgentPhase): void
  /** normalized stop reason of the turn ('max_tokens' = cut off by the token limit); transports may omit this */
  onStopReason?(reason: string): void
  onDone(): void
  onError(error: string): void
}

export interface AgentStreamHandle {
  /** abort the in-flight turn; the transport must still emit onDone afterwards */
  cancel(): void
}

export interface AgentTransport {
  stream(request: AgentStreamRequest, callbacks: AgentStreamCallbacks): AgentStreamHandle
}
