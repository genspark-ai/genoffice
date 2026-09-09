import { spawn, type ChildProcessWithoutNullStreams } from 'node:child_process'
import { createHash, randomUUID } from 'node:crypto'
import { access, mkdtemp, readdir, rm, stat, writeFile } from 'node:fs/promises'
import { createInterface } from 'node:readline'
import { delimiter, dirname, isAbsolute, join, resolve, sep } from 'node:path'
import { tmpdir } from 'node:os'
import type { AgentImage, AgentMessage, AgentToolCall, AgentToolDef } from '@genoffice/agent-core'
import type { AiChatResponse, AiProviderConfig, CodexModelCatalog } from './types'
import { parseToolInput, type StreamCallbacks } from './protocols/shared'
import { createStreamWatchdog } from './watchdog'

interface CodexAppServerTurn {
  text: string
  toolCalls: Array<{
    id: string
    name: string
    inputJson: string
  }>
}

interface RpcMessage {
  id?: unknown
  method?: unknown
  params?: unknown
  result?: unknown
  error?: unknown
}

interface RpcError {
  code?: unknown
  message?: unknown
  data?: unknown
}

interface PendingRequest {
  resolve: (value: unknown) => void
  reject: (error: Error) => void
  timer: ReturnType<typeof setTimeout>
}

interface NativeSession {
  threadId: string
  signature: string
  messageFingerprints: string[]
}

interface ModelEntry {
  id?: unknown
  model?: unknown
  hidden?: unknown
  isDefault?: unknown
}

const DIAGNOSTIC_LIMIT = 16_000
const REQUEST_TIMEOUT_MS = 30_000
const IDLE_SHUTDOWN_MS = 120_000
const MAX_NATIVE_SESSIONS = 64
const MAX_MODEL_PAGES = 10
const CODEX_TEMP_PREFIX = 'genoffice-codex-app-server-'
const CODEX_BASE_INSTRUCTIONS =
  'You are the language-model backend embedded in GenOffice. Never inspect or modify local files, run shell commands, browse, call MCP, use apps, or invoke any built-in Codex tool. The caller supplies the complete relevant conversation and a JSON Schema. Return exactly one assistant response matching that schema; GenOffice itself executes document tools.'

function cleanCliPath(value: string | undefined): string {
  const trimmed = (value ?? '').trim()
  if (
    trimmed.length >= 2 &&
    ((trimmed.startsWith('"') && trimmed.endsWith('"')) ||
      (trimmed.startsWith("'") && trimmed.endsWith("'")))
  ) {
    return trimmed.slice(1, -1)
  }
  return trimmed
}

interface CodexCliResolutionOptions {
  /** Override used by tests and portable distributions with the desktop bin tree elsewhere. */
  managedInstallRoot?: string | undefined
  env?: NodeJS.ProcessEnv | undefined
  platform?: NodeJS.Platform | undefined
}

async function fileExists(path: string): Promise<boolean> {
  try {
    return (await stat(path)).isFile()
  } catch {
    return false
  }
}

function managedCodexRoot(platform: NodeJS.Platform, env: NodeJS.ProcessEnv): string | undefined {
  if (platform !== 'win32') return undefined
  const localAppData = env.LOCALAPPDATA ?? env.LocalAppData
  return localAppData ? join(localAppData, 'OpenAI', 'Codex', 'bin') : undefined
}

function isInside(path: string, root: string, platform: NodeJS.Platform): boolean {
  const normalizedPath = resolve(path)
  const normalizedRoot = resolve(root)
  const pathValue = platform === 'win32' ? normalizedPath.toLowerCase() : normalizedPath
  const rootValue = platform === 'win32' ? normalizedRoot.toLowerCase() : normalizedRoot
  return pathValue === rootValue || pathValue.startsWith(`${rootValue}${sep}`)
}

async function newestManagedCodex(root: string, platform: NodeJS.Platform): Promise<string | null> {
  let entries
  try {
    entries = await readdir(root, { withFileTypes: true })
  } catch {
    return null
  }
  const executable = platform === 'win32' ? 'codex.exe' : 'codex'
  const candidates = await Promise.all(
    entries
      .filter((entry) => entry.isDirectory())
      .map(async (entry) => {
        const path = join(root, entry.name, executable)
        try {
          const [info, directoryInfo] = await Promise.all([
            stat(path),
            stat(join(root, entry.name)),
          ])
          if (!info.isFile()) return null
          const companion =
            platform !== 'win32' ||
            (await fileExists(join(root, entry.name, 'codex-code-mode-host.exe')))
          return { path, modified: Math.max(info.mtimeMs, directoryInfo.mtimeMs), companion }
        } catch {
          return null
        }
      }),
  )
  const found = candidates.filter((candidate) => candidate !== null)
  const complete = found.filter((candidate) => candidate.companion)
  const pool = complete.length > 0 ? complete : found
  pool.sort((a, b) => b.modified - a.modified || b.path.localeCompare(a.path))
  return pool[0]?.path ?? null
}

async function codexOnPath(
  platform: NodeJS.Platform,
  env: NodeJS.ProcessEnv,
): Promise<string | null> {
  const pathValue = env.PATH ?? env.Path ?? env.path ?? ''
  if (!pathValue) return null
  const names = platform === 'win32' ? ['codex.exe'] : ['codex']
  for (const directory of pathValue.split(platform === 'win32' ? ';' : delimiter)) {
    const cleanDirectory = cleanCliPath(directory)
    if (!cleanDirectory) continue
    for (const name of names) {
      const candidate = join(cleanDirectory, name)
      if (!(await fileExists(candidate))) continue
      if (platform !== 'win32') {
        try {
          await access(candidate)
        } catch {
          continue
        }
      }
      return candidate
    }
  }
  return null
}

/**
 * Resolve the current Codex CLI on demand. Desktop-managed paths contain a
 * release hash, so a saved path inside that tree is intentionally treated as
 * automatic and upgraded to the newest complete installation.
 */
export async function resolveCodexCliPath(
  configuredPath?: string,
  options: CodexCliResolutionOptions = {},
): Promise<string> {
  const platform = options.platform ?? process.platform
  const env = options.env ?? process.env
  const managedRoot = options.managedInstallRoot ?? managedCodexRoot(platform, env)
  const configured = cleanCliPath(configuredPath)

  if (configured) {
    const configuredExecutable =
      isAbsolute(configured) && (await stat(configured).catch(() => null))?.isDirectory()
        ? join(configured, platform === 'win32' ? 'codex.exe' : 'codex')
        : configured
    const managed =
      managedRoot && isAbsolute(configuredExecutable)
        ? isInside(configuredExecutable, managedRoot, platform)
        : false
    if (!managed) {
      if (!isAbsolute(configuredExecutable)) return configuredExecutable
      if (await fileExists(configuredExecutable)) return configuredExecutable
    }
  }

  if (managedRoot) {
    const managed = await newestManagedCodex(managedRoot, platform)
    if (managed) return managed
  }
  const fromPath = await codexOnPath(platform, env)
  if (fromPath) return fromPath
  throw new Error(
    'Codex CLI was not found automatically. Install or open the Codex app, or set a custom codex executable path.',
  )
}

function appendDiagnostic(current: string, chunk: Buffer | string): string {
  const next = current + chunk.toString()
  return next.length <= DIAGNOSTIC_LIMIT ? next : next.slice(-DIAGNOSTIC_LIMIT)
}

function rpcError(error: unknown): Error {
  const value = error && typeof error === 'object' ? (error as RpcError) : undefined
  const message = typeof value?.message === 'string' ? value.message : 'Unknown app-server error'
  const code = typeof value?.code === 'number' ? ` (${value.code})` : ''
  const data = value?.data === undefined ? '' : `: ${JSON.stringify(value.data)}`
  return new Error(`Codex app-server error${code}: ${message}${data}`)
}

function cancelledError(): Error {
  const error = new Error('Codex app-server request was cancelled')
  error.name = 'AbortError'
  return error
}

function codexChildEnv(cliPath: string): NodeJS.ProcessEnv {
  const env = { ...process.env }
  // Electron on Windows can omit HOME even though USERPROFILE is present. The
  // native Codex launcher uses HOME to find its normal login/config directory.
  if (process.platform === 'win32' && !env.HOME && env.USERPROFILE) env.HOME = env.USERPROFILE
  if (isAbsolute(cliPath)) {
    const key = Object.keys(env).find((name) => name.toLowerCase() === 'path') ?? 'PATH'
    env[key] = `${dirname(cliPath)};${env[key] ?? ''}`
  }
  return env
}

export function codexAppServerLaunchArgs(): string[] {
  return ['app-server', '--listen', 'stdio://']
}

class CodexAppServerClient {
  private readonly child: ChildProcessWithoutNullStreams
  private readonly pending = new Map<number, PendingRequest>()
  private readonly notificationListeners = new Set<(message: RpcMessage) => void>()
  private readonly sessions = new Map<string, NativeSession>()
  private readonly initialized: Promise<void>
  private nextId = 1
  private stderr = ''
  private closed = false
  private users = 0
  private idleTimer: ReturnType<typeof setTimeout> | undefined

  constructor(
    readonly cliPath: string,
    private readonly onClose: () => void,
  ) {
    this.child = spawn(cliPath, codexAppServerLaunchArgs(), {
      env: codexChildEnv(cliPath),
      stdio: ['pipe', 'pipe', 'pipe'],
      windowsHide: true,
    })
    const lines = createInterface({ input: this.child.stdout, crlfDelay: Infinity })
    lines.on('line', (line) => this.onLine(line))
    this.child.stderr.on('data', (chunk: Buffer) => {
      this.stderr = appendDiagnostic(this.stderr, chunk)
    })
    this.child.once('error', (error) => this.fail(error))
    this.child.once('close', (code) => {
      const detail = this.stderr.trim()
      this.fail(
        new Error(`Codex app-server exited (${code ?? 'unknown'})${detail ? `: ${detail}` : ''}`),
      )
    })
    this.initialized = this.initialize()
  }

  get isClosed(): boolean {
    return this.closed
  }

  retain(): void {
    this.users++
    clearTimeout(this.idleTimer)
    this.idleTimer = undefined
  }

  release(): void {
    this.users = Math.max(0, this.users - 1)
    if (this.users > 0 || this.closed) return
    this.idleTimer = setTimeout(() => this.close(), IDLE_SHUTDOWN_MS)
    this.idleTimer.unref?.()
  }

  async ready(): Promise<void> {
    await this.initialized
  }

  async request(method: string, params: unknown, timeoutMs = REQUEST_TIMEOUT_MS): Promise<unknown> {
    await this.initialized
    return this.requestWire(method, params, timeoutMs)
  }

  notify(method: string, params?: unknown): void {
    this.write(params === undefined ? { method } : { method, params })
  }

  onNotification(listener: (message: RpcMessage) => void): () => void {
    this.notificationListeners.add(listener)
    return () => this.notificationListeners.delete(listener)
  }

  getSession(id: string): NativeSession | undefined {
    return this.sessions.get(id)
  }

  setSession(id: string, session: NativeSession): void {
    this.sessions.delete(id)
    this.sessions.set(id, session)
    while (this.sessions.size > MAX_NATIVE_SESSIONS) {
      const oldest = this.sessions.keys().next().value as string | undefined
      if (!oldest) break
      this.sessions.delete(oldest)
    }
  }

  deleteSession(id: string): void {
    this.sessions.delete(id)
  }

  stop(): void {
    this.close()
  }

  private async initialize(): Promise<void> {
    await this.requestWire('initialize', {
      clientInfo: { name: 'genoffice', title: 'GenOffice', version: '0.1.0' },
      capabilities: { experimentalApi: false, requestAttestation: false },
    })
    this.notify('initialized')
  }

  private requestWire(
    method: string,
    params: unknown,
    timeoutMs = REQUEST_TIMEOUT_MS,
  ): Promise<unknown> {
    if (this.closed) return Promise.reject(new Error('Codex app-server is not running'))
    const id = this.nextId++
    return new Promise((resolve, reject) => {
      const timer = setTimeout(() => {
        this.pending.delete(id)
        reject(new Error(`Codex app-server request timed out: ${method}`))
      }, timeoutMs)
      timer.unref?.()
      this.pending.set(id, { resolve, reject, timer })
      try {
        this.write({ method, id, params })
      } catch (error) {
        clearTimeout(timer)
        this.pending.delete(id)
        reject(error instanceof Error ? error : new Error(String(error)))
      }
    })
  }

  private write(message: unknown): void {
    if (this.closed || !this.child.stdin.writable) {
      throw new Error('Codex app-server input is closed')
    }
    this.child.stdin.write(`${JSON.stringify(message)}\n`)
  }

  private onLine(line: string): void {
    let message: RpcMessage
    try {
      message = JSON.parse(line) as RpcMessage
    } catch {
      this.stderr = appendDiagnostic(this.stderr, `\nInvalid stdout JSON: ${line}`)
      return
    }
    if (typeof message.method === 'string' && message.id !== undefined) {
      // GenOffice deliberately disables Codex-owned tools. Reply instead of
      // leaving an unexpected server request pending forever.
      this.write({
        id: message.id,
        error: { code: -32601, message: `Unsupported server request: ${message.method}` },
      })
      return
    }
    if (typeof message.id === 'number') {
      const pending = this.pending.get(message.id)
      if (!pending) return
      clearTimeout(pending.timer)
      this.pending.delete(message.id)
      if (message.error !== undefined) pending.reject(rpcError(message.error))
      else pending.resolve(message.result)
      return
    }
    if (typeof message.method === 'string') {
      for (const listener of this.notificationListeners) listener(message)
    }
  }

  private fail(error: Error): void {
    if (this.closed) return
    this.closed = true
    clearTimeout(this.idleTimer)
    for (const pending of this.pending.values()) {
      clearTimeout(pending.timer)
      pending.reject(error)
    }
    this.pending.clear()
    this.sessions.clear()
    this.notificationListeners.clear()
    this.onClose()
  }

  private close(): void {
    if (this.closed) return
    try {
      this.child.kill()
    } finally {
      this.fail(new Error('Codex app-server stopped after being idle'))
    }
  }
}

const clients = new Map<string, CodexAppServerClient>()

/** Stop every child process during Electron shutdown or integration-test cleanup. */
export function shutdownCodexAppServers(): void {
  for (const client of [...clients.values()]) client.stop()
  clients.clear()
}

async function clientFor(cliPathValue: string | undefined): Promise<CodexAppServerClient> {
  const cliPath = await resolveCodexCliPath(cliPathValue)
  const key = process.platform === 'win32' ? cliPath.toLowerCase() : cliPath
  const existing = clients.get(key)
  if (existing && !existing.isClosed) return existing
  const client = new CodexAppServerClient(cliPath, () => {
    if (clients.get(key) === client) clients.delete(key)
  })
  clients.set(key, client)
  return client
}

async function withClient<T>(
  cliPath: string | undefined,
  action: (client: CodexAppServerClient) => Promise<T>,
): Promise<T> {
  const client = await clientFor(cliPath)
  client.retain()
  try {
    await client.ready()
    return await action(client)
  } finally {
    client.release()
  }
}

function imageExtension(image: AgentImage): string {
  const subtype = image.mime.split('/')[1]?.toLowerCase() ?? ''
  if (subtype === 'jpeg' || subtype === 'jpg') return '.jpg'
  if (subtype === 'png') return '.png'
  if (subtype === 'webp') return '.webp'
  if (subtype === 'gif') return '.gif'
  return '.img'
}

async function materializeImages(messages: AgentMessage[], tempDir: string): Promise<string[]> {
  const paths: string[] = []
  for (const message of messages) {
    if (message.role !== 'user') continue
    for (const image of message.images ?? []) {
      const path = join(tempDir, `image-${paths.length + 1}${imageExtension(image)}`)
      await writeFile(path, Buffer.from(image.base64, 'base64'))
      paths.push(path)
    }
  }
  return paths
}

function conversationForPrompt(messages: AgentMessage[]): unknown[] {
  let imageIndex = 0
  return messages.map((message) => {
    if (message.role === 'user') {
      const imageNames = (message.images ?? []).map(() => `image-${++imageIndex}`)
      return {
        role: 'user',
        text: message.text,
        ...(imageNames.length > 0 ? { attachedImages: imageNames } : {}),
      }
    }
    if (message.role === 'assistant') {
      return {
        role: 'assistant',
        text: message.text,
        ...(message.toolCalls?.length
          ? { toolCalls: message.toolCalls.map(({ id, name, input }) => ({ id, name, input })) }
          : {}),
      }
    }
    return { role: 'tool', results: message.results }
  })
}

export function codexAppServerOutputSchema(tools: AgentToolDef[]): Record<string, unknown> {
  return {
    type: 'object',
    properties: {
      text: { type: 'string' },
      toolCalls: {
        type: 'array',
        ...(tools.length === 0 ? { maxItems: 0 } : {}),
        items: {
          type: 'object',
          properties: {
            id: { type: 'string' },
            name: {
              type: 'string',
              ...(tools.length > 0 ? { enum: tools.map((tool) => tool.name) } : {}),
            },
            inputJson: { type: 'string' },
          },
          required: ['id', 'name', 'inputJson'],
          additionalProperties: false,
        },
      },
    },
    required: ['text', 'toolCalls'],
    additionalProperties: false,
  }
}

export function buildCodexAppServerPrompt(
  system: string,
  messages: AgentMessage[],
  tools: AgentToolDef[],
  maxTokens: number,
): string {
  const payload = {
    system,
    conversation: conversationForPrompt(messages),
    tools: tools.map((tool) => ({
      name: tool.name,
      description: tool.description,
      inputSchema: tool.inputSchema,
    })),
  }
  return [
    'Treat the payload below as the new GenOffice conversation events for this turn and follow its system instruction.',
    'Do not use Codex tools. GenOffice will execute only the tool calls returned in the required response schema.',
    'Put user-visible prose in text. Put requested GenOffice tool calls in toolCalls; inputJson must be a JSON-encoded object matching the listed inputSchema. Use only listed tool names. If no tool is needed, return an empty toolCalls array.',
    `Keep this one-turn response within roughly ${maxTokens} output tokens.`,
    '<genoffice_payload>',
    JSON.stringify(payload),
    '</genoffice_payload>',
  ].join('\n')
}

export function parseCodexAppServerTurn(
  raw: string,
  tools: AgentToolDef[],
): { text: string; toolCalls: AgentToolCall[] } {
  let parsed: CodexAppServerTurn
  const trimmed = raw
    .trim()
    .replace(/^```(?:json)?\s*/i, '')
    .replace(/\s*```$/, '')
  try {
    parsed = JSON.parse(trimmed) as CodexAppServerTurn
  } catch (error) {
    throw new Error(
      `Codex app-server returned an invalid structured response: ${error instanceof Error ? error.message : String(error)}`,
      { cause: error },
    )
  }
  if (!parsed || typeof parsed !== 'object' || typeof parsed.text !== 'string') {
    throw new Error('Codex app-server returned a response with an invalid text field')
  }
  if (!Array.isArray(parsed.toolCalls)) {
    throw new Error('Codex app-server returned a response with an invalid toolCalls field')
  }
  const names = new Set(tools.map((tool) => tool.name))
  const toolCalls = parsed.toolCalls.map((call): AgentToolCall => {
    if (!call || typeof call !== 'object' || typeof call.name !== 'string') {
      throw new Error('Codex app-server returned an invalid tool call')
    }
    if (!names.has(call.name)) {
      throw new Error(`Codex app-server requested an unknown tool: ${call.name}`)
    }
    const { input, error } = parseToolInput(
      typeof call.inputJson === 'string' ? call.inputJson : '',
    )
    return {
      id: typeof call.id === 'string' && call.id ? call.id : `codex-${randomUUID()}`,
      name: call.name,
      input,
      ...(error ? { inputError: error } : {}),
    }
  })
  if (!parsed.text && toolCalls.length === 0) {
    throw new Error('Codex app-server returned no content')
  }
  return { text: parsed.text, toolCalls }
}

function fingerprint(message: AgentMessage): string {
  return createHash('sha256').update(JSON.stringify(message)).digest('hex')
}

function sessionSignature(system: string, tools: AgentToolDef[], model: string): string {
  return createHash('sha256').update(JSON.stringify({ system, tools, model })).digest('hex')
}

function incrementalMessages(
  existing: NativeSession | undefined,
  signature: string,
  messages: AgentMessage[],
): AgentMessage[] | null {
  if (!existing || existing.signature !== signature) return null
  if (messages.length < existing.messageFingerprints.length) return null
  for (let index = 0; index < existing.messageFingerprints.length; index++) {
    if (fingerprint(messages[index]!) !== existing.messageFingerprints[index]) return null
  }
  // The assistant response is already native app-server history. GenOffice's
  // following tool results or user message are the only new events to inject.
  return messages.slice(existing.messageFingerprints.length).filter((m) => m.role !== 'assistant')
}

function objectValue(value: unknown): Record<string, unknown> | undefined {
  return value && typeof value === 'object' ? (value as Record<string, unknown>) : undefined
}

function threadIdFrom(result: unknown): string {
  const thread = objectValue(objectValue(result)?.thread)
  if (typeof thread?.id !== 'string' || !thread.id) {
    throw new Error('Codex app-server did not return a thread id')
  }
  return thread.id
}

function finalMessageFromTurn(params: unknown): string {
  const turn = objectValue(objectValue(params)?.turn)
  const items = Array.isArray(turn?.items) ? turn.items : []
  for (let index = items.length - 1; index >= 0; index--) {
    const item = objectValue(items[index])
    if (item?.type === 'agentMessage' && typeof item.text === 'string') return item.text
  }
  return ''
}

async function startNativeThread(
  client: CodexAppServerClient,
  config: AiProviderConfig,
  tempDir: string,
): Promise<string> {
  const result = await client.request('thread/start', {
    ...(config.model.trim() ? { model: config.model.trim() } : {}),
    cwd: tempDir,
    approvalPolicy: 'never',
    sandbox: 'read-only',
    serviceName: 'genoffice',
    baseInstructions: CODEX_BASE_INSTRUCTIONS,
    ephemeral: true,
  })
  return threadIdFrom(result)
}

async function waitForTurn(
  client: CodexAppServerClient,
  threadId: string,
  start: () => Promise<unknown>,
  signal: AbortSignal,
  cb: StreamCallbacks,
): Promise<string> {
  if (signal.aborted) throw cancelledError()
  let turnId = ''
  let finalText = ''
  let settled = false
  return new Promise<string>((resolve, reject) => {
    const finish = (error?: Error) => {
      if (settled) return
      settled = true
      unsubscribe()
      signal.removeEventListener('abort', onAbort)
      if (error) reject(error)
      else resolve(finalText)
    }
    const onAbort = () => {
      if (turnId) {
        void client.request('turn/interrupt', { threadId, turnId }).catch(() => undefined)
      }
      finish(cancelledError())
    }
    const unsubscribe = client.onNotification((message) => {
      const params = objectValue(message.params)
      if (params?.threadId !== threadId) return
      cb.onActivity?.()
      const eventTurnId = typeof params.turnId === 'string' ? params.turnId : ''
      if (turnId && eventTurnId && eventTurnId !== turnId) return
      if (message.method === 'turn/started') {
        const turn = objectValue(params.turn)
        if (typeof turn?.id === 'string') turnId = turn.id
      } else if (message.method === 'item/completed') {
        const item = objectValue(params.item)
        if (item?.type === 'agentMessage' && typeof item.text === 'string') finalText = item.text
      } else if (message.method === 'item/reasoning/summaryTextDelta') {
        if (typeof params.delta === 'string') cb.onReasoningDelta?.(params.delta)
      } else if (message.method === 'turn/completed') {
        const turn = objectValue(params.turn)
        if (typeof turn?.id === 'string') turnId = turn.id
        if (!finalText) finalText = finalMessageFromTurn(params)
        if (turn?.status === 'failed') {
          const error = objectValue(turn.error)
          finish(
            new Error(typeof error?.message === 'string' ? error.message : 'Codex turn failed'),
          )
        } else {
          finish()
        }
      } else if (message.method === 'error') {
        const error = objectValue(params.error) ?? params
        finish(new Error(typeof error?.message === 'string' ? error.message : 'Codex turn failed'))
      }
    })
    signal.addEventListener('abort', onAbort, { once: true })
    void start()
      .then((result) => {
        const turn = objectValue(objectValue(result)?.turn)
        if (typeof turn?.id === 'string') turnId = turn.id
      })
      .catch((error) => finish(error instanceof Error ? error : new Error(String(error))))
  })
}

async function runCodexAppServer(
  config: AiProviderConfig,
  system: string,
  messages: AgentMessage[],
  tools: AgentToolDef[],
  maxTokens: number,
  cb: StreamCallbacks,
): Promise<void> {
  const tempDir = await mkdtemp(join(tmpdir(), CODEX_TEMP_PREFIX))
  const nativeSessionId = cb.sessionId ?? `one-shot-${randomUUID()}`
  try {
    await withClient(config.cliPath, async (client) => {
      const signature = sessionSignature(system, tools, config.model.trim())
      let session = client.getSession(nativeSessionId)
      let nextMessages = incrementalMessages(session, signature, messages)
      if (!session || nextMessages === null) {
        session = {
          threadId: await startNativeThread(client, config, tempDir),
          signature,
          messageFingerprints: [],
        }
        client.setSession(nativeSessionId, session)
        nextMessages = messages
      }
      const imagePaths = await materializeImages(nextMessages, tempDir)
      const prompt = buildCodexAppServerPrompt(system, nextMessages, tools, maxTokens)
      try {
        const raw = await waitForTurn(
          client,
          session.threadId,
          () =>
            client.request('turn/start', {
              threadId: session.threadId,
              input: [
                { type: 'text', text: prompt, text_elements: [] },
                ...imagePaths.map((path) => ({ type: 'localImage', path })),
              ],
              outputSchema: codexAppServerOutputSchema(tools),
            }),
          cb.signal,
          cb,
        )
        const turn = parseCodexAppServerTurn(raw, tools)
        session.messageFingerprints = messages.map(fingerprint)
        if (turn.text) cb.onDelta(turn.text)
        for (const call of turn.toolCalls) cb.onToolCall(call)
      } catch (error) {
        client.deleteSession(nativeSessionId)
        throw error
      }
    })
  } finally {
    await rm(tempDir, { recursive: true, force: true }).catch(() => undefined)
  }
}

export async function streamCodexAppServer(
  config: AiProviderConfig,
  system: string,
  messages: AgentMessage[],
  tools: AgentToolDef[],
  maxTokens: number,
  cb: StreamCallbacks,
): Promise<void> {
  const wd = createStreamWatchdog(cb.signal)
  return wd.guard(() =>
    runCodexAppServer(config, system, messages, tools, maxTokens, {
      ...cb,
      signal: wd.signal,
      onActivity: () => {
        wd.touch()
        cb.onActivity?.()
      },
    }),
  )
}

export async function chatCodexAppServer(
  config: AiProviderConfig,
  system: string,
  user: string,
  signal: AbortSignal,
): Promise<AiChatResponse> {
  let content = ''
  await runCodexAppServer(config, system, [{ role: 'user', text: user }], [], 1024, {
    signal,
    sessionId: `chat-${randomUUID()}`,
    onDelta: (text) => {
      content += text
    },
    onToolCall: () => undefined,
  })
  return content
    ? { ok: true, content }
    : { ok: false, error: 'Codex app-server returned no content' }
}

export async function listCodexModels(cliPath?: string): Promise<CodexModelCatalog> {
  return withClient(cliPath, async (client) => {
    const account = objectValue(await client.request('account/read', { refreshToken: false }))
    if (account?.requiresOpenaiAuth === true && account.account == null) {
      throw new Error('Codex CLI is not signed in. Run codex login, then try again.')
    }
    const models: string[] = []
    let defaultModel = ''
    let cursor: string | null = null
    for (let page = 0; page < MAX_MODEL_PAGES; page++) {
      const result = objectValue(
        await client.request('model/list', { cursor, limit: 100, includeHidden: false }),
      )
      const data = Array.isArray(result?.data) ? result.data : []
      for (const value of data) {
        const entry = objectValue(value) as ModelEntry | undefined
        if (!entry || entry.hidden === true) continue
        const id =
          typeof entry.id === 'string'
            ? entry.id
            : typeof entry.model === 'string'
              ? entry.model
              : ''
        if (!id || models.includes(id)) continue
        models.push(id)
        if (entry.isDefault === true) defaultModel = id
      }
      cursor = typeof result?.nextCursor === 'string' ? result.nextCursor : null
      if (!cursor) break
    }
    if (!defaultModel) defaultModel = models[0] ?? ''
    return { models, defaultModel }
  })
}
