import type { AgentSkill, ExecutedToolCall } from './skill'
import type {
  AgentImage,
  AgentMessage,
  AgentStreamHandle,
  AgentToolCall,
  AgentToolDef,
  AgentToolResult,
  AgentTransport,
  ToolExecution,
} from './types'

export interface ToolExecutedEvent<TSnapshot> {
  call: AgentToolCall
  execution: ToolExecution
  /**
   * Snapshot captured just before this tool ran; present only on the first
   * mutating tool of a run (hook for one-click rollback UIs).
   */
  snapshotBefore?: TSnapshot | undefined
}

export interface AgentRunResult {
  /** final assistant text of the run ('' when cut off) */
  text: string
  cancelled: boolean
  /** true when maxTurns was reached; text is the partial answer from the no-tools finalizing turn */
  turnLimit: boolean
  /** the final turn hit the token limit (stop_reason max_tokens): text is incomplete; set only when true */
  truncated?: boolean
}

export interface AgentLoopEvents<TSnapshot> {
  /** cumulative assistant text of the current turn (call per delta) */
  onText?(text: string): void
  /** a tool is about to execute (UI shows a live "running" indicator; onToolExecuted always follows) */
  onToolStart?(call: AgentToolCall): void
  onToolExecuted?(event: ToolExecutedEvent<TSnapshot>): void
  /** a turn requested tools and they ran; the loop is going back to the model */
  onTurnEnd?(): void
  onDone?(result: AgentRunResult): void
  onError?(error: string): void
}

/** Context compaction config (budget tracked in UTF-8 bytes rather than message count) */
export interface CompactionOptions {
  /** History size that triggers compaction (UTF-8 bytes, default 256KB) */
  maxBytes?: number
  /** Size of recent messages kept after compaction (bytes, default 96KB, cut at a user boundary) */
  keepRecentBytes?: number
  /** Disable LLM summarization and use only the mechanical digest (for tests/offline) */
  disableLlmSummary?: boolean
}

export interface AgentLoopOptions<TSnapshot = unknown> {
  transport: AgentTransport
  skill: AgentSkill
  events?: AgentLoopEvents<TSnapshot>
  /** hard cap on model round-trips per run (default DEFAULT_MAX_TURNS) */
  maxTurns?: number
  /** history cap in messages, trimmed at user-turn boundaries (default 40) */
  maxHistory?: number
  /** Context compaction; false disables it (enabled by default with default thresholds) */
  compaction?: CompactionOptions | false
  /** capture rollback state; invoked right before tools run (see snapshotBefore) */
  captureSnapshot?(): TSnapshot
  /** wrap instruction + skill context into the user message text */
  formatUserMessage?(instruction: string, context: string): string
  /** appended to the system prompt each turn (e.g. reply-language directive following the UI language) */
  systemSuffix?(): string
}

const COMPACT_MAX_BYTES = 256 * 1024
const COMPACT_KEEP_RECENT_BYTES = 96 * 1024
/** Pre-truncation of each tool output in the summary request (the compaction request itself must not blow up on huge outputs) */
const SUMMARIZE_TOOL_OUTPUT_MAX = 2_000
const SUMMARIZE_TIMEOUT_MS = 30_000
/** When over budget mid-run, keep the last N tool messages verbatim and truncate earlier outputs to this length */
const STALE_TOOL_KEEP_RECENT = 2
const STALE_TOOL_OUTPUT_MAX = 1_000

/** Unified turn budget across the suite's chat panels (apps may still override per loop) */
export const DEFAULT_MAX_TURNS = 100

/** Cap on consecutive turns whose tool input was all unusable (a turn that executes a call resets it); abort beyond it (keeps the model from burning turns on bad JSON) */
const MAX_INPUT_PARSE_RETRIES = 3

/**
 * Whether a property's own JSON Schema declares null as an acceptable value.
 * A tool that genuinely takes null (e.g. clearing a style value) says so in
 * the schema, so the required-field check must not reject it. Covers the
 * spellings in use: `type: ['string','null']`, OpenAPI/Gemini `nullable: true`,
 * an enum listing null, and an anyOf/oneOf branch typed null.
 */
function schemaAcceptsNull(schema: unknown): boolean {
  if (!schema || typeof schema !== 'object') return false
  const s = schema as {
    type?: unknown
    nullable?: unknown
    enum?: unknown
    anyOf?: unknown
    oneOf?: unknown
  }
  if (s.nullable === true) return true
  // `type` is a bare "null" in an anyOf/oneOf branch, a list in `type: ['string','null']`
  if (s.type === 'null') return true
  if (Array.isArray(s.type) && s.type.includes('null')) return true
  if (Array.isArray(s.enum) && s.enum.includes(null)) return true
  for (const branch of [s.anyOf, s.oneOf]) {
    if (Array.isArray(branch) && branch.some((b) => schemaAcceptsNull(b))) return true
  }
  return false
}

/**
 * Required fields the model left out of a tool call, per the tool's JSON
 * Schema. Providers turn an empty argument stream into `{}` without an
 * inputError (the model wrote prose instead of arguments, or a gateway dropped
 * the argument stream), so without this check the empty object reaches the
 * skill and fails with a tool-specific message instead of a targeted retry.
 * `null` counts as missing too: a garbled field arrives as `"ops": null`, and
 * downstream coercion (Number(null) === 0) would pass a silently wrong value
 * to the tool. A field whose schema declares null as valid is exempt.
 */
export function missingRequiredFields(
  tool: AgentToolDef | undefined,
  input: Record<string, unknown>,
): string[] {
  const required = tool?.inputSchema.required
  if (!Array.isArray(required)) return []
  const properties = tool?.inputSchema.properties
  const propSchema = (field: string): unknown =>
    properties && typeof properties === 'object' && !Array.isArray(properties)
      ? (properties as Record<string, unknown>)[field]
      : undefined
  return required.filter((field): field is string => {
    if (typeof field !== 'string') return false
    const value = input[field]
    if (value === undefined) return true
    return value === null && !schemaAcceptsNull(propSchema(field))
  })
}

/** Plain object (not null, not an array) — the only schema/value shape we walk into */
function isPlainRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}

/** Nesting cap for `items`; a value deeper than this is passed through unvalidated */
const MAX_SCHEMA_DEPTH = 4

/**
 * Why one value contradicts one property schema, or undefined when it is fine.
 * Deliberately not a JSON Schema implementation: only the keywords this repo's
 * tools actually declare are honoured, and a schema that declares none of them
 * always returns undefined so a permissive or schema-less tool keeps working.
 */
function schemaViolation(
  value: unknown,
  schema: Record<string, unknown>,
  depth: number,
): string | undefined {
  if (depth > MAX_SCHEMA_DEPTH) return undefined
  // enum: only trusted when every member is a primitive, otherwise comparing
  // would need a deep-equality walk this does not do
  const allowed = schema.enum
  if (
    Array.isArray(allowed) &&
    allowed.every((v) => v === null || (!isPlainRecord(v) && typeof v !== 'function'))
  ) {
    if (!allowed.some((v) => Object.is(v, value))) {
      return `expected one of ${allowed.map((v) => JSON.stringify(v)).join(', ')}`
    }
  }
  // type: a string, or an array of strings. Keywords we do not model are dropped
  // from the list; an all-unknown list means "unconstrained", not "invalid".
  const declared = (Array.isArray(schema.type) ? schema.type : [schema.type]).filter(
    (t): t is string => typeof t === 'string',
  )
  const typeNames = declared.filter((t) => TYPE_CHECKS[t])
  if (typeNames.length > 0 && !typeNames.some((t) => TYPE_CHECKS[t]!(value))) {
    return `expected ${typeNames.join(' or ')}`
  }
  if (typeof value === 'number' && Number.isFinite(value)) {
    // draft-04 spells exclusivity as a boolean sibling; honour it so the
    // boundary value is not rejected by an inclusive check
    if (
      typeof schema.minimum === 'number' &&
      schema.exclusiveMinimum !== true &&
      value < schema.minimum
    ) {
      return `must be >= ${schema.minimum}`
    }
    if (
      typeof schema.maximum === 'number' &&
      schema.exclusiveMaximum !== true &&
      value > schema.maximum
    ) {
      return `must be <= ${schema.maximum}`
    }
  }
  if (typeof value === 'string') {
    if (typeof schema.minLength === 'number' && value.length < schema.minLength) {
      return `must be at least ${schema.minLength} characters`
    }
    if (typeof schema.maxLength === 'number' && value.length > schema.maxLength) {
      return `must be at most ${schema.maxLength} characters`
    }
  }
  // items: only the single-schema form; a tuple (array of schemas) is out of scope
  if (Array.isArray(value) && isPlainRecord(schema.items)) {
    for (let i = 0; i < value.length; i++) {
      const bad = schemaViolation(value[i], schema.items, depth + 1)
      if (bad) return `item ${i} ${bad}`
    }
  }
  return undefined
}

/** JSON Schema primitive types this repo's tools declare; NaN is not a valid number */
const TYPE_CHECKS: Record<string, (value: unknown) => boolean> = {
  string: (v) => typeof v === 'string',
  number: (v) => typeof v === 'number' && Number.isFinite(v),
  integer: (v) => typeof v === 'number' && Number.isInteger(v),
  boolean: (v) => typeof v === 'boolean',
  array: (v) => Array.isArray(v),
  object: (v) => isPlainRecord(v),
  null: (v) => v === null,
}

/**
 * Fields the model *sent* but whose value contradicts the tool's own JSON
 * Schema ({"count": "twelve"}, {"rows": 999999999999} against a bounded
 * integer). Such a value otherwise reaches the tool, which coerces it
 * (Number("twelve") -> NaN, a capped range silently clamped) and returns a
 * confident wrong answer — where missingRequiredFields would have produced a
 * targeted retry instead.
 *
 * Only `properties` are inspected, and a field with no declared constraints is
 * never reported: a tool with an empty, missing, or non-JSON-Schema inputSchema
 * (this repo has `{}` and `{ type: 'object' }` tools) validates nothing, exactly
 * as before. Absent fields are skipped — they belong to missingRequiredFields.
 */
export function invalidArgumentFields(
  tool: AgentToolDef | undefined,
  input: Record<string, unknown>,
): string[] {
  // `?.` on inputSchema as well: a tool can arrive over IPC without one
  const properties = tool?.inputSchema?.properties
  if (!isPlainRecord(properties)) return []
  const violations: string[] = []
  for (const [field, fieldSchema] of Object.entries(properties)) {
    if (!isPlainRecord(fieldSchema)) continue
    const value = input[field]
    if (value === undefined) continue
    const violation = schemaViolation(value, fieldSchema, 0)
    if (violation) violations.push(`"${field}" ${violation}`)
  }
  return violations
}

/**
 * Degenerate-loop guards. Weak models (BYOK/local endpoints especially) can
 * repeat the exact same turn forever or keep issuing failing tool calls; with
 * a large turn budget these must abort early instead of burning it.
 */
const MAX_IDENTICAL_TURNS = 3
const MAX_ALL_ERROR_TURNS = 8

/**
 * Backoff schedule for in-place same-turn retries on empty-stream errors.
 * The "(empty stream)" suffix is a cross-layer contract with the ai-provider
 * protocols: the gateway closed the SSE stream without content, tool calls, or
 * message framing — a transient soft-failure. The turn produced nothing and
 * history is untouched, so re-sending the identical request is idempotent;
 * retrying here keeps one gateway hiccup from killing a long multi-tool run.
 */
const EMPTY_STREAM_RETRY_DELAYS_MS = [1_000, 3_000]
/**
 * A stream that closed while a tool's arguments were still streaming (buffered
 * server-side, cut by a gateway idle timeout) never delivered a tool call, so
 * history is untouched and one replay is safe; it is billed, hence one attempt.
 */
const TOOL_ARGS_DROP_MARK = 'while sending tool arguments'
const TOOL_ARGS_DROP_RETRIES = 1

const TURN_LIMIT_NOTE =
  '[System] The tool-call turn limit for this request has been reached; no more tools may be called this turn. ' +
  'Answer directly from the information already gathered; if the task is unfinished, briefly state what is done and what remains.'

export const TOOL_ABORTED_OUTPUT =
  '(the user stopped the run while this tool was still executing; its result was discarded)'

const TOOL_ABORTED = Symbol('tool-aborted')

async function awaitToolOrAbort(
  tool: ToolExecution | Promise<ToolExecution>,
  signal: AbortSignal | undefined,
): Promise<ToolExecution | typeof TOOL_ABORTED> {
  if (!signal) return tool
  if (signal.aborted) return TOOL_ABORTED
  const running = Promise.resolve(tool)
  return new Promise<ToolExecution | typeof TOOL_ABORTED>((resolve, reject) => {
    const onAbort = (): void => {
      resolve(TOOL_ABORTED)
      running.catch(() => undefined)
    }
    signal.addEventListener('abort', onAbort, { once: true })
    running.then(
      (execution) => {
        signal.removeEventListener('abort', onAbort)
        resolve(execution)
      },
      (error: unknown) => {
        signal.removeEventListener('abort', onAbort)
        reject(error)
      },
    )
  })
}

/**
 * Terminal assistant text when tools mutated the artifact (or an edits-only
 * turn was restored) and the model returned no prose. Must be non-empty so
 * provider message converters never emit empty assistant content, which breaks
 * multi-turn follow-ups (see finishTurn / restore).
 * Exported so apps can substitute a localized / tool-derived summary in the UI.
 */
export const COMPLETED_VIA_TOOLS_TEXT = '(completed tool actions; no text reply)'

/**
 * Models default to their training-cutoff year without this (e.g. web searches
 * for "... 2024"). Leads the system prompt and spells out the year: measured
 * against claude-opus-4-7 with the docs prompt, the date alone (front or tail)
 * still produced cutoff-year searches in 6/6 runs; naming the year fixed all.
 */
export function runtimePreamble(now = new Date()): string {
  const pad = (n: number) => String(n).padStart(2, '0')
  const date = `${now.getFullYear()}-${pad(now.getMonth() + 1)}-${pad(now.getDate())}`
  return `Today's date is ${date}; the current year is ${now.getFullYear()}.\n\n`
}

const SUMMARIZE_SYSTEM =
  'You are a conversation compressor. Compress this editing session between the user and the AI assistant into a concise summary so later turns can continue with context. ' +
  "Keep: the user's goals and key instructions, completed changes (which files/pages/elements were modified), important facts and data, and outstanding items. " +
  'For specific figures/statistics, mark their provenance: figures from the user or from tool results (e.g. web_search) keep their source; figures the assistant produced without a source must be marked "(unverified)" so later turns do not treat them as established facts. ' +
  'Omit: pleasantries, tool-call details, and intermediate trial and error. Use a bullet list of at most 400 words. Write the summary in the same language as the conversation. Output only the summary body, with no preamble.'

/** Prefix of the synthetic user message that carries the compacted-history summary */
const COMPACT_SUMMARY_PREFIX = '[Summary of earlier conversation'
const COMPACT_SUMMARY_HEADER = '[Summary of earlier conversation (auto-compacted)]'
const COMPACT_SUMMARY_ACK = 'Understood, continuing from the progress so far.'

/** Approximate UTF-8 byte count (ASCII 1 byte, CJK etc. 3; surrogate pairs count as 6 — slight overestimate is harmless) */
function utf8Size(s: string): number {
  let n = 0
  for (let i = 0; i < s.length; i++) {
    const c = s.charCodeAt(i)
    n += c < 0x80 ? 1 : c < 0x800 ? 2 : 3
  }
  return n
}

/** Approximate byte cost of one message (text + tool inputs/outputs + image base64) */
function messageSize(m: AgentMessage): number {
  if (m.role === 'tool') {
    return m.results.reduce((n, r) => n + utf8Size(r.output) + 40, 0)
  }
  let n = utf8Size(m.text)
  if (m.role === 'user' && m.images) {
    n += m.images.reduce((s, img) => s + img.base64.length, 0)
  }
  if (m.role === 'assistant' && m.toolCalls) {
    for (const c of m.toolCalls) {
      try {
        n += utf8Size(JSON.stringify(c.input)) + 40
      } catch {
        n += 40
      }
    }
  }
  return n
}

function historySize(messages: readonly AgentMessage[]): number {
  return messages.reduce((n, m) => n + messageSize(m), 0)
}

/** Mechanical digest when LLM summarization is unavailable: bullet list of user instructions + final replies */
function mechanicalDigest(dropped: readonly AgentMessage[]): string {
  const lines: string[] = []
  for (const m of dropped) {
    if (m.role === 'user' && !m.text.startsWith(COMPACT_SUMMARY_PREFIX)) {
      lines.push(`- User: ${m.text.slice(0, 200)}`)
    } else if (m.role === 'assistant' && m.text && !m.toolCalls?.length) {
      lines.push(`  Reply: ${m.text.slice(0, 200)}`)
    }
  }
  return lines.join('\n').slice(0, 4_000) || '(earlier conversation omitted)'
}

/**
 * Generic ReAct loop: user message -> model turn (text + tool calls) ->
 * execute tools -> feed results back -> repeat until the model answers with
 * plain text. History persists across runs, so follow-up questions work.
 */
export class AgentLoop<TSnapshot = unknown> {
  private readonly options: AgentLoopOptions<TSnapshot>
  private history: AgentMessage[] = []
  private handle: AgentStreamHandle | null = null
  private running = false
  private cancelled = false
  private turns = 0
  /** Finalizing turn after hitting the turn limit: no tools, let the model answer from what it has read */
  private finalizing = false
  private mutationSeen = false
  private inputParseFails = 0
  /** signature (text + tool calls) of the previous turn, for the identical-turn guard */
  private lastTurnSig = ''
  private identicalTurns = 0
  private allErrorTurns = 0
  private turnStopReason: string | null = null
  private turnText = ''
  private turnReasoning = ''
  private toolCalls: AgentToolCall[] = []
  /** tools actually executed during this run, fed to skill.verifyResponse */
  private executedCalls: ExecutedToolCall[] = []
  /** verifyResponse may force one extra corrective turn per run — never more */
  private verifyRetryUsed = false
  /** user message of the in-flight run; a failed run rolls it (and everything after) back out of history */
  private runUserMsg: AgentMessage | null = null
  /** invalidates stale transport callbacks after cancel/reset */
  private generation = 0
  /** per-run abort: aborted on cancel(); long tools (e.g. generate_deck) use it to break internal loops */
  private abortController: AbortController | null = null

  constructor(options: AgentLoopOptions<TSnapshot>) {
    this.options = options
  }

  get busy(): boolean {
    return this.running
  }

  get messages(): readonly AgentMessage[] {
    return this.history
  }

  /**
   * Seed the conversation with restored history (e.g. transcript reloaded from
   * disk when a document reopens), so follow-up instructions keep their context.
   * No-op unless the loop is idle with an empty history.
   * Old messages over the compaction budget fold into a mechanical digest
   * (no LLM request on restore, guaranteeing zero latency).
   */
  restore(messages: readonly AgentMessage[]): void {
    if (this.running || this.history.length > 0 || messages.length === 0) return
    // Edits-only runs persist an assistant message with no text; give it a placeholder
    // so the turn stays paired and providers never see an empty assistant content block.
    // Turn-limit notes persisted by older builds are stripped: they are stale
    // directives ("no more tools may be called") that poison every later run.
    const normalized = messages
      .filter((m) => !(m.role === 'user' && m.text === TURN_LIMIT_NOTE))
      .map((m) =>
        m.role === 'assistant' && !m.text ? { ...m, text: COMPLETED_VIA_TOOLS_TEXT } : m,
      )
    // Unanswered user messages (a failed or interrupted run persisted them without a
    // reply) must not re-enter the model context: trailing ones would pair with the
    // next instruction as one turn, adjacent ones read as a combined instruction
    const answered = normalized.filter(
      (m, i) => m.role !== 'user' || (normalized[i + 1] && normalized[i + 1]!.role !== 'user'),
    )
    // An assistant message whose tool calls never received results (a run interrupted
    // between the model's tool call and its execution) would reach the provider as
    // unpaired tool_calls and 400 the next turn. Pair each orphan call with an
    // isError result — the same signal the cancel path synthesizes — so the
    // transcript stays valid and the model can retry the call.
    const paired: AgentMessage[] = []
    for (let i = 0; i < answered.length; i++) {
      const m = answered[i]!
      paired.push(m)
      if (m.role === 'assistant' && m.toolCalls?.length && answered[i + 1]?.role !== 'tool') {
        paired.push({
          role: 'tool',
          results: m.toolCalls.map((call) => ({
            id: call.id,
            name: call.name,
            output: TOOL_ABORTED_OUTPUT,
            isError: true,
          })),
        })
      }
    }
    this.history = paired
    if (this.history.length === 0) return
    if (this.compactionEnabled()) {
      const { maxBytes, keepRecentBytes } = this.compactBudget()
      if (historySize(this.history) > maxBytes) {
        const cut = this.findCompactCut(keepRecentBytes)
        if (cut > 0) {
          const digest = mechanicalDigest(this.history.slice(0, cut))
          this.history = [
            { role: 'user', text: `${COMPACT_SUMMARY_HEADER}\n${digest}` },
            { role: 'assistant', text: COMPACT_SUMMARY_ACK },
            ...this.history.slice(cut),
          ]
        }
      }
    }
    this.trimHistory()
  }

  /** images: inline attachments for this user turn (vision input; see AgentImage) */
  run(instruction: string, images?: AgentImage[]): void {
    if (this.running || !instruction) return
    this.running = true
    this.cancelled = false
    this.turns = 0
    this.finalizing = false
    this.mutationSeen = false
    this.inputParseFails = 0
    this.lastTurnSig = ''
    this.identicalTurns = 0
    this.allErrorTurns = 0
    this.executedCalls = []
    this.verifyRetryUsed = false
    this.abortController = new AbortController()
    // buildContext and formatUserMessage are consumer-supplied and run before any
    // turn exists, so a throw here would escape run() with `running` still true:
    // the guard at the top then drops every later message silently, cancel()
    // no-ops, and only reset() frees the loop. composeSkills fans buildContext out
    // to every sub-skill, each of which reads the live document, so a document
    // mid-transition is enough to wedge the panel. The user message was never
    // pushed, so the rollback in failRun() is a no-op and only the report matters.
    try {
      const context = this.options.skill.buildContext?.() ?? ''
      const format =
        this.options.formatUserMessage ??
        ((instr: string, ctx: string) => (ctx ? `${instr}\n\n${ctx}` : instr))
      const userMsg: AgentMessage = {
        role: 'user',
        text: format(instruction, context),
        ...(images?.length ? { images } : {}),
      }
      void this.beginRun(userMsg)
    } catch (err) {
      this.failRun(err instanceof Error ? err.message : String(err))
    }
  }

  /** Compact (if needed), push the user message, then start the turn. Compaction failure doesn't block the run. */
  private async beginRun(userMsg: AgentMessage): Promise<void> {
    const generation = this.generation
    try {
      await this.maybeCompact()
    } catch {
      // Proceed with the run even if compaction fails (an over-budget history only costs more, it's still correct)
    }
    if (generation !== this.generation) return // reset during compaction
    if (this.cancelled) {
      this.running = false
      this.options.events?.onDone?.({ text: '', cancelled: true, turnLimit: false })
      return
    }
    // Leftover unanswered user message (a previous run failed before replying):
    // drop it so the model never sees two adjacent user turns as one combined instruction
    while (this.history.at(-1)?.role === 'user') this.history.pop()
    this.trimHistory()
    // Reasoning echo only matters inside a run's own tool loop; drop it from
    // finished runs so it stops costing tokens on every later request.
    this.history = this.history.map((m) =>
      m.role === 'assistant' && m.reasoning ? { ...m, reasoning: undefined } : m,
    )
    if (userMsg.role === 'user') {
      userMsg = { ...userMsg, text: sanitizeAgentPayload(userMsg.text) }
    }
    this.runUserMsg = userMsg
    this.history.push(userMsg)
    this.startTurn()
  }

  /**
   * A run failed: remove its user message and every message after it, so the
   * failed instruction can't be silently re-executed by the next run.
   */
  private rollbackFailedRun(): void {
    const msg = this.runUserMsg
    this.runUserMsg = null
    if (!msg) return
    const i = this.history.lastIndexOf(msg)
    if (i >= 0) this.history.splice(i)
  }

  // ── Context compaction: fold old conversation into a summary, keep recent messages verbatim ──

  private compactionEnabled(): boolean {
    return this.options.compaction !== false
  }

  private compactBudget(): { maxBytes: number; keepRecentBytes: number } {
    const opt = this.options.compaction === false ? undefined : this.options.compaction
    return {
      maxBytes: opt?.maxBytes ?? COMPACT_MAX_BYTES,
      keepRecentBytes: opt?.keepRecentBytes ?? COMPACT_KEEP_RECENT_BYTES,
    }
  }

  /**
   * Find the compaction cut at a user boundary: accumulate from the tail up to keepRecentBytes.
   * Returns the start index of the kept segment; if no suitable boundary exists,
   * fall back to keeping the last user turn.
   */
  private findCompactCut(keepRecentBytes: number): number {
    let kept = 0
    let cut = -1
    for (let i = this.history.length - 1; i >= 0; i--) {
      kept += messageSize(this.history[i]!)
      if (kept > keepRecentBytes && cut >= 0) break
      if (this.history[i]!.role === 'user') cut = i
    }
    if (cut < 0) {
      for (let i = this.history.length - 1; i >= 0; i--) {
        if (this.history[i]!.role === 'user') return i
      }
    }
    return cut
  }

  private async maybeCompact(): Promise<void> {
    if (!this.compactionEnabled()) return
    const { maxBytes, keepRecentBytes } = this.compactBudget()
    if (historySize(this.history) <= maxBytes) return
    const cut = this.findCompactCut(keepRecentBytes)
    if (cut <= 0) return // no foldable prefix
    const generation = this.generation
    const dropped = this.history.slice(0, cut)
    const opt = this.options.compaction === false ? undefined : this.options.compaction
    let summary: string | null = null
    if (!opt?.disableLlmSummary) summary = await this.summarizeViaLlm(dropped)
    // A reset may have cleared history or started a new conversation while
    // the summary was pending. Discard its result before touching that history.
    if (generation !== this.generation) return
    if (!summary) summary = mechanicalDigest(dropped)
    this.history = [
      { role: 'user', text: `${COMPACT_SUMMARY_HEADER}\n${summary}` },
      { role: 'assistant', text: COMPACT_SUMMARY_ACK },
      ...this.history.slice(cut),
    ]
  }

  /** Hand the folded conversation to the model for a summary; returns null on failure/timeout (falls back to the mechanical digest). */
  private summarizeViaLlm(dropped: readonly AgentMessage[]): Promise<string | null> {
    // Slim down the summary request itself: pre-truncate tool outputs, strip images
    const slim: AgentMessage[] = dropped.map((m) => {
      if (m.role === 'tool') {
        return {
          role: 'tool' as const,
          results: m.results.map((r) => ({
            ...r,
            output: r.output.slice(0, SUMMARIZE_TOOL_OUTPUT_MAX),
          })),
        }
      }
      if (m.role === 'user' && m.images?.length) return { role: 'user' as const, text: m.text }
      return m
    })
    return new Promise((resolve) => {
      let text = ''
      let settled = false
      let handle: AgentStreamHandle | null = null
      const finish = (v: string | null) => {
        if (settled) return
        settled = true
        clearTimeout(timer)
        resolve(v)
      }
      const timer = setTimeout(() => {
        finish(null)
        handle?.cancel()
      }, SUMMARIZE_TIMEOUT_MS)
      try {
        // Attach to this.handle so cancel() can abort the summary request when the user clicks stop
        handle = this.options.transport.stream(
          {
            system: SUMMARIZE_SYSTEM,
            messages: [
              ...slim,
              { role: 'user', text: 'Compress the conversation above as instructed.' },
            ],
            tools: [],
          },
          {
            onDelta: (t) => {
              if (settled) return
              text += t
            },
            onToolCall: () => {
              /* the summary turn gets no tools */
            },
            onDone: () => finish(text.trim() || null),
            onError: () => finish(null),
          },
        )
        this.handle = handle
      } catch {
        finish(null)
      }
    })
  }

  /**
   * When over budget mid-run (between tool turns), truncate stale tool outputs:
   * keep structure (tool_use/tool_result pairs intact), cut content only,
   * and keep the most recent N verbatim.
   */
  private squashStaleToolOutputs(): void {
    if (!this.compactionEnabled()) return
    const { maxBytes } = this.compactBudget()
    if (historySize(this.history) <= maxBytes) return
    let recent = 0
    for (let i = this.history.length - 1; i >= 0; i--) {
      const m = this.history[i]!
      if (m.role !== 'tool') continue
      recent++
      if (recent <= STALE_TOOL_KEEP_RECENT) continue
      m.results = m.results.map((r) =>
        r.output.length > STALE_TOOL_OUTPUT_MAX
          ? {
              ...r,
              output: `${r.output.slice(0, STALE_TOOL_OUTPUT_MAX)}\n…(output truncated: too long)`,
            }
          : r,
      )
    }
  }

  cancel(): void {
    if (!this.running) return
    this.cancelled = true
    // abort lets long tools mid-execution (internal LLM loops etc.) stop promptly
    this.abortController?.abort()
    // the transport emits onDone after aborting, which finalizes the run
    this.handle?.cancel()
  }

  /** drop the conversation (e.g. when a different document is opened) */
  reset(): void {
    this.generation++
    this.abortController?.abort()
    this.handle?.cancel()
    this.handle = null
    this.running = false
    this.cancelled = false
    this.history = []
    this.runUserMsg = null
  }

  /**
   * Terminal failure path for the callback guard sites: clear `running` (or
   * every later message is silently dropped, and even cancel() no-ops), roll
   * the failed instruction back out of history so no orphaned tool_use is left
   * behind, then report. onError is the consumer's last callback, so a throw
   * from it must not escape and wedge the loop a second time.
   */
  private failRun(message: string): void {
    this.running = false
    this.rollbackFailedRun()
    try {
      this.options.events?.onError?.(message)
    } catch {
      // nothing left to notify
    }
  }

  /**
   * finishTurn() runs consumer-supplied callbacks (onToolStart, onToolExecuted
   * with its snapshotBefore, onTurnEnd, onDone) and the captureSnapshot hook.
   * Its promise is discarded, so a throw from any of them would escape as an
   * unhandled rejection and leave `running` true forever. Same terminal-state
   * guarantee as the tools-getter guard in startTurn().
   */
  private settleTurn(): void {
    void this.finishTurn().catch((err: unknown) => {
      this.failRun(err instanceof Error ? err.message : String(err))
    })
  }

  /** Runs at run boundaries only (restore / before a new user message): a long run's tail is all assistant/tool messages, and cutting mid-run would empty the request. */
  private trimHistory(): void {
    const max = this.options.maxHistory ?? 40
    if (this.history.length <= max) return
    // cut only at a user message so tool_use/tool_result pairs stay intact
    let i = this.history.length - max
    while (i < this.history.length && this.history[i]!.role !== 'user') i++
    if (i >= this.history.length) return // no user boundary in the window: keep history over budget
    const next = this.history.slice(i)
    if (this.runUserMsg && !next.includes(this.runUserMsg)) return
    this.history = next
  }

  private startTurn(retriesUsed = 0): void {
    const generation = this.generation
    this.turnText = ''
    this.turnReasoning = ''
    this.toolCalls = []
    this.turnStopReason = null
    // Some transports emit an extra onDone after cancel — this turn may finalize only once
    let settled = false
    try {
      this.handle = this.options.transport.stream(
        {
          system:
            runtimePreamble() +
            this.options.skill.systemPrompt +
            (this.options.systemSuffix?.() ?? ''),
          messages: [...this.history],
          tools: this.finalizing ? [] : this.options.skill.tools,
        },
        {
          onDelta: (text) => {
            if (generation !== this.generation || settled) return
            this.turnText += text
            try {
              this.options.events?.onText?.(this.turnText)
            } catch (err) {
              // A transport drives onDelta from its own async event handler
              // (see the Electron IPC transport), so this throw never reaches
              // the try/catch around stream() below.
              settled = true
              this.failRun(err instanceof Error ? err.message : String(err))
            }
          },
          onReasoning: (text) => {
            if (generation !== this.generation || settled) return
            this.turnReasoning += text
          },
          onToolCall: (call) => {
            if (generation !== this.generation || settled) return
            this.toolCalls.push(call)
          },
          onStopReason: (reason) => {
            if (generation !== this.generation || settled) return
            this.turnStopReason = reason
          },
          onDone: () => {
            if (generation !== this.generation || settled) return
            settled = true
            this.settleTurn()
          },
          onError: (error) => {
            if (generation !== this.generation || settled) return
            settled = true
            // The no-partial-output guard keeps the empty-stream retry idempotent (an
            // empty stream never emits deltas, but a mislabeled error must not replay
            // a turn whose text/tool calls the UI already saw). A dropped tool-argument
            // stream may have shown text first; that text is simply re-rendered.
            const emptyDelay = EMPTY_STREAM_RETRY_DELAYS_MS[retriesUsed]
            const retryEmpty =
              emptyDelay !== undefined &&
              error.includes('(empty stream)') &&
              !this.turnText &&
              this.toolCalls.length === 0
            const retryDrop =
              retriesUsed < TOOL_ARGS_DROP_RETRIES &&
              error.includes(TOOL_ARGS_DROP_MARK) &&
              this.toolCalls.length === 0
            const delay = retryEmpty ? emptyDelay : EMPTY_STREAM_RETRY_DELAYS_MS[0]
            if ((retryEmpty || retryDrop) && !this.cancelled) {
              setTimeout(() => {
                if (generation !== this.generation) return
                // Stopped during the backoff window: finalize like a normal cancel
                if (this.cancelled) {
                  this.settleTurn()
                  return
                }
                this.startTurn(retriesUsed + 1)
              }, delay)
              return
            }
            this.running = false
            this.rollbackFailedRun()
            this.options.events?.onError?.(error)
          },
        },
      )
    } catch (err) {
      // A skill's tools getter (a duplicate name in a composed skill) can throw
      // before any callback runs: this keeps the run from staying busy forever.
      this.running = false
      this.rollbackFailedRun()
      this.options.events?.onError?.(err instanceof Error ? err.message : String(err))
    }
  }

  private async finishTurn(): Promise<void> {
    const { events, skill, captureSnapshot } = this.options
    const toolCalls = this.toolCalls

    // Claimed-action guard: before accepting a final text turn, let the skill
    // check the claims in it against the tools that actually ran this run.
    // A returned correction forces one more model turn (tools stay available,
    // so the model can perform the missing action or reword its claim).
    if (toolCalls.length === 0 && !this.cancelled && !this.finalizing) {
      // snapshot copy: the live array keeps growing if the corrective turn
      // runs more tools, and the hook must see the state at check time
      const correction =
        !this.verifyRetryUsed && this.turnText && skill.verifyResponse
          ? skill.verifyResponse(this.turnText, [...this.executedCalls])
          : null
      if (correction) {
        this.verifyRetryUsed = true
        this.history.push({ role: 'assistant', text: this.turnText })
        this.history.push({ role: 'user', text: correction })
        // No onTurnEnd here: UIs use it to seal the current assistant bubble,
        // which would keep the rejected claim visible. Without it, the
        // corrective turn's cumulative onText overwrites the bubble in place.
        this.startTurn()
        return
      }
    }

    // final turn: no tools requested, the user stopped the run, or the
    // no-tools finalizing turn after hitting the limit
    // (a cancelled turn drops its tool calls — no results would follow)
    if (toolCalls.length === 0 || this.cancelled || this.finalizing) {
      // The turn-limit note has served its purpose once the finalizing turn
      // ends. Left in history it would tell every later run "no more tools may
      // be called" — a stale directive models obey (or worse, echo verbatim
      // over and over; see public issue about BYOK models repeating it).
      if (this.finalizing) {
        for (let i = this.history.length - 1; i >= 0; i--) {
          const m = this.history[i]!
          if (m.role === 'user' && m.text === TURN_LIMIT_NOTE) {
            this.history.splice(i, 1)
            break
          }
        }
      }
      // Models often end a tool-using run with an empty text turn ("I'm done").
      // Leaving assistant text empty in history then poisons the next user
      // prompt: Anthropic rejects empty content arrays, Gemini rejects empty
      // parts, and OpenAI-compatible routes send content:null with no tool_calls —
      // all of which make follow-up turns fail or return empty again (see
      // genoffice#12 / #22: first prompt works, second shows "no summary").
      // Same normalization as restore(), applied unconditionally: cancelled and
      // read-only empty turns poison follow-ups just the same. onDone still
      // reports the raw turn text so app UIs keep their localized fallbacks
      // instead of surfacing this English placeholder.
      this.history.push({ role: 'assistant', text: this.turnText || COMPLETED_VIA_TOOLS_TEXT })
      this.running = false
      this.runUserMsg = null
      events?.onDone?.({
        text: this.turnText,
        cancelled: this.cancelled,
        turnLimit: this.finalizing,
        // set only when true so exact-shape consumers/tests stay unaffected
        ...(this.turnStopReason === 'max_tokens' && !this.cancelled ? { truncated: true } : {}),
      })
      return
    }

    // Strip turn-local execution hints (inputError/truncated) from the stored
    // history: they are not model context, and transports with strict message
    // schemas (the Electron IPC bridge) reject unknown tool-call keys when the
    // history is echoed back on the next turn. The OpenAI-compatible stream
    // paths attach `inputError: undefined` on every parsed call, so without
    // this the second turn of any custom-provider agent run fails validation.
    this.history.push({
      role: 'assistant',
      text: this.turnText,
      toolCalls: toolCalls.map(({ id, name, input, signature }) => ({
        id,
        name,
        input,
        ...(signature ? { signature } : {}),
      })),
      // interleaved-thinking models degrade in tool loops unless their reasoning is echoed back
      ...(this.turnReasoning ? { reasoning: this.turnReasoning } : {}),
    })
    const generation = this.generation
    const results: AgentToolResult[] = []
    let turnMutated = false
    // unusable-input streak is counted per turn: a batch of empty calls in one
    // turn is one failed attempt, and any executed call in the turn resets it
    let unusableInTurn = false
    let executedInTurn = false
    for (const call of toolCalls) {
      // The user hit stop while an earlier tool was running: skip remaining tools,
      // but fill in paired error results to keep tool_use/tool_result pairs valid for the next request
      if (this.cancelled) {
        results.push({
          id: call.id,
          name: call.name,
          output: '(the user stopped the run; this tool was not executed)',
          isError: true,
        })
        continue
      }
      // Unusable input (truncated by the token limit, or JSON that failed to parse):
      // don't execute; feed a targeted error back so the model retries correctly
      const tool = skill.tools.find((t) => t.name === call.name)
      const unreadable = call.truncated || call.inputError
      const missing = unreadable ? [] : missingRequiredFields(tool, call.input)
      // A value that contradicts the schema is as unusable as a missing one, and
      // is reported last: telling the model which field is absent is what it needs first.
      const invalid =
        unreadable || missing.length > 0 ? [] : invalidArgumentFields(tool, call.input)
      if (unreadable || missing.length > 0 || invalid.length > 0) {
        unusableInTurn = true
        const output = call.truncated
          ? 'Tool arguments were cut off by the output length limit; the tool was not executed. Split this operation into several smaller tool calls (less content per call) and try again.'
          : call.inputError
            ? `Tool input JSON failed to parse; the tool was not executed: ${call.inputError}\nFix the arguments (make sure quotes inside strings are escaped) and call again.`
            : missing.length > 0
              ? `Tool call ${call.name} is missing the required argument(s) ${missing.map((f) => `"${f}"`).join(', ')}; the tool was not executed. Put the arguments in the tool call itself (not in your reply text) and call again with every required field.`
              : `Tool call ${call.name} got ${invalid.length === 1 ? 'an argument' : 'arguments'} that do not match the tool's input schema — ${invalid.join('; ')}; the tool was not executed. Call again with every value of the declared type and within the declared limits (a number must be a bare JSON number, not a quoted string).`
        results.push({ id: call.id, name: call.name, output, isError: true })
        events?.onToolExecuted?.({
          call,
          execution: { output, isError: true, summary: call.name },
        })
        continue
      }
      executedInTurn = true
      events?.onToolStart?.(call)
      const snapshot = !this.mutationSeen ? captureSnapshot?.() : undefined
      let raced: ToolExecution | typeof TOOL_ABORTED
      try {
        raced = await awaitToolOrAbort(
          skill.executeTool(call, this.abortController?.signal),
          this.abortController?.signal,
        )
      } catch (e) {
        raced = {
          output: e instanceof Error ? e.message : String(e),
          isError: true,
          summary: call.name,
        }
      }
      if (generation !== this.generation) return // reset while a tool was running
      if (raced === TOOL_ABORTED) {
        const aborted: ToolExecution = {
          output: TOOL_ABORTED_OUTPUT,
          isError: true,
          summary: call.name,
        }
        this.executedCalls.push({ name: call.name, ok: false })
        results.push({ id: call.id, name: call.name, output: aborted.output, isError: true })
        events?.onToolExecuted?.({ call, execution: aborted })
        continue
      }
      const execution = raced
      this.executedCalls.push({ name: call.name, ok: !execution.isError })
      const firstMutation = !!execution.mutated && !this.mutationSeen
      if (execution.mutated) {
        this.mutationSeen = true
        turnMutated = true
      }
      results.push({
        id: call.id,
        name: call.name,
        output: execution.output,
        isError: execution.isError,
      })
      events?.onToolExecuted?.({
        call,
        execution,
        snapshotBefore: firstMutation ? snapshot : undefined,
      })
    }
    this.history.push({ role: 'tool', results })
    if (executedInTurn) this.inputParseFails = 0
    else if (unusableInTurn) this.inputParseFails++

    // Cancelled while tools were executing: finish immediately, no further model request
    if (this.cancelled) {
      this.running = false
      this.runUserMsg = null
      events?.onDone?.({ text: this.turnText, cancelled: true, turnLimit: false })
      return
    }

    // Bad-input retries hit the cap: abort instead of burning more turns
    if (this.inputParseFails >= MAX_INPUT_PARSE_RETRIES) {
      this.running = false
      this.rollbackFailedRun()
      events?.onError?.(
        `Tool input was unusable (unparseable, truncated or missing required arguments) ${MAX_INPUT_PARSE_RETRIES} times in a row; retries stopped, please send the request again`,
      )
      return
    }

    // A turn where every tool call failed makes no progress; a long streak
    // (unknown-tool loops from malformed BYOK streams, hallucinated tools)
    // would otherwise burn the whole turn budget re-erroring.
    this.allErrorTurns = results.every((r) => r.isError) ? this.allErrorTurns + 1 : 0
    if (this.allErrorTurns >= MAX_ALL_ERROR_TURNS) {
      this.running = false
      this.rollbackFailedRun()
      events?.onError?.(
        `Every tool call failed for ${MAX_ALL_ERROR_TURNS} turns in a row; the run was stopped. Please send the request again`,
      )
      return
    }

    // Identical-turn guard: a model (typically a weak BYOK/local endpoint)
    // re-emitting the same text, tool calls AND tool outputs is looping, not
    // progressing. Turns that mutated the artifact are exempt (repeating an
    // identical edit is legitimate progress), and changing outputs break the
    // streak (so poll-style tools survive).
    const turnSig = JSON.stringify([
      this.turnText,
      toolCalls.map(({ name, input }) => [name, input]),
      results.map((r) => r.output),
    ])
    if (turnSig === this.lastTurnSig && !turnMutated) {
      if (++this.identicalTurns >= MAX_IDENTICAL_TURNS) {
        this.running = false
        this.rollbackFailedRun()
        events?.onError?.(
          'The model kept repeating the exact same turn without making progress; the run was stopped. Please send the request again',
        )
        return
      }
    } else {
      this.lastTurnSig = turnSig
      this.identicalTurns = 0
    }

    this.turns++
    if (this.turns >= (this.options.maxTurns ?? DEFAULT_MAX_TURNS)) {
      // Don't throw away the context already gathered: append one no-tools turn for a partial answer
      this.finalizing = true
      this.history.push({ role: 'user', text: TURN_LIMIT_NOTE })
    }
    // Long runs (e.g. page-by-page generation) over budget mid-way: truncate stale tool outputs so each turn doesn't resend a huge payload
    this.squashStaleToolOutputs()
    events?.onTurnEnd?.()
    this.startTurn()
  }
}

/**
 * Redact secret-looking tokens from an outgoing user message so accidentally
 * pasted API keys, URL credentials, and password assignments don't reach
 * remote model APIs verbatim.
 *
 * Imported from public PR #32 (BuiltByHarshil), with the credential pattern
 * narrowed to URL userinfo (scheme://user:pass@host) so ordinary "a:b@c"
 * prose is never rewritten.
 */
export function sanitizeAgentPayload(payload: string): string {
  return (
    payload
      .replace(
        /-----BEGIN [A-Z ]*PRIVATE KEY-----[\s\S]*?-----END [A-Z ]*PRIVATE KEY-----/g,
        '[REDACTED_PRIVATE_KEY]',
      )
      // Truncated paste: header plus base64 body lines, no END marker.
      .replace(
        /-----BEGIN [A-Z ]*PRIVATE KEY-----(?:\r?\n[A-Za-z0-9+/=]+(?=\r?\n|$))*/g,
        '[REDACTED_PRIVATE_KEY]',
      )
      .replace(/\b(?:sk-|AIza|ghp_|secret_)[A-Za-z0-9_-]{16,}/g, '[REDACTED_API_KEY]')
      .replace(/\b(?:AKIA|ASIA)[0-9A-Z]{16}\b/g, '[REDACTED_API_KEY]')
      .replace(/\bxox[abeoprs]-[A-Za-z0-9-]{10,}/g, '[REDACTED_API_KEY]')
      // The scheme run is bounded for the same reason as the identifier prefix below:
      // an unbounded `[a-z0-9+.-]*` in front of a literal `://` is ambiguous, so every
      // start offset consumed the whole run and backtracked looking for the `://`
      // (measured: 96 KB of hex took ~12s in this regex alone, ~70s for the call).
      // A 30-char scheme covers every registered one; past it the match simply starts
      // mid-scheme, which still redacts the password — only the captured scheme
      // prefix is shorter.
      .replace(/([a-z][a-z0-9+.-]{0,30}:\/\/[^\s:@/]+):[^\s@/]+@/gi, '$1:[REDACTED_CREDENTIALS]@')
      .replace(
        /(password|passwd|secret_key|private_key)(\s*[:=]\s*)["'][^"']+["']/gi,
        '$1$2"[REDACTED_SECURE_TOKEN]"',
      )
      // Unquoted `password=abc123`: the value must be 6+ chars with a non-letter,
      // so "password: is in the vault" prose stays untouched.
      // The identifier prefix is bounded: an unbounded `\w*` in front of the
      // alternation overlaps it, so at every start offset the engine consumed the
      // whole word run and backtracked one character at a time to place the
      // keyword — quadratic in the length of an unbroken [A-Za-z0-9_] run. A user
      // pasting a hex dump or a base64url token froze the renderer for ~70s
      // (measured: 100 KB of hex, this function, one call). 64 chars is far more
      // than any real `my_password`-style prefix and keeps the match set identical.
      .replace(
        /(?<!\/)(\w{0,64}(?:password|passwd|secret_key|private_key))(\s*[:=]\s*)(?=[^\s"',;]*[^A-Za-z\s"',;])[^\s"',;]{6,}/gi,
        '$1$2[REDACTED_SECURE_TOKEN]',
      )
  )
}
