import { execFile, spawn, type ChildProcess } from 'node:child_process'
import {
  chmodSync,
  closeSync,
  createWriteStream,
  existsSync,
  mkdirSync,
  openSync,
  readdirSync,
  readSync,
  renameSync,
  rmSync,
  statSync,
} from 'node:fs'
import { createServer } from 'node:net'
import { cpus, freemem, homedir } from 'node:os'
import { dirname, join } from 'node:path'
import { promisify } from 'node:util'
import {
  DEFAULT_LOCAL_VLM_SETTINGS,
  LOCAL_VLM_DEFAULT_MODEL,
  LOCAL_VLM_LLAMA_RELEASES_API,
  LOCAL_VLM_MIN_RAM_MB,
  LOCAL_VLM_MODEL_FILENAME,
  LOCAL_VLM_MODEL_URL,
  LOCAL_VLM_MMPROJ_FILENAME,
  LOCAL_VLM_MMPROJ_URL,
  setLocalVlmUsable,
  type LocalVlmDownloadProgress,
  type LocalVlmSettings,
  type LocalVlmStatus,
} from './local-vlm-meta'

const execFileAsync = promisify(execFile)

const GGUF_MAGIC = Buffer.from('GGUF', 'ascii')
const DEFAULT_STARTUP_TIMEOUT_MS = 120_000
const DEFAULT_IDLE_UNLOAD_MS = 10 * 60_000
const HEALTH_POLL_MS = 250

export interface LlamaReleaseAsset {
  name: string
  browser_download_url: string
  size?: number
}

export interface LocalVlmRuntimeBridge {
  status(): LocalVlmStatus
  isRunnable(): boolean
  shouldFallbackFromAgy(): boolean
  ensureRunning(signal?: AbortSignal): Promise<{ baseUrl: string; model: string }>
  testConnection(signal?: AbortSignal): Promise<{ ok: boolean; error?: string; latencyMs?: number }>
  stop(): Promise<void>
}

let activeRuntime: LocalVlmRuntimeBridge | null = null

export function registerLocalVlmRuntime(runtime: LocalVlmRuntimeBridge | null): void {
  activeRuntime = runtime
  setLocalVlmUsable(runtime ? runtime.isRunnable() : null)
}

export function getLocalVlmRuntime(): LocalVlmRuntimeBridge | null {
  return activeRuntime
}

/**
 * Resolves optimal CPU threads for llama.cpp on x86/ARM hybrid architectures.
 * Over-allocating logical threads across both P-cores and E-cores causes severe
 * thread barrier stalls and memory bus thrashing during AVX2 GEMM.
 * Aligning to physical P-cores (typically 4-6 threads) maximizes tokens/sec.
 */
export function resolveOptimalCpuThreads(userThreads = 0): number {
  if (userThreads > 0) return userThreads
  try {
    const cpuList = cpus()
    const count = cpuList?.length ?? 4
    if (count >= 12) return 6
    if (count >= 8) return 4
    if (count >= 4) return 4
    return Math.max(1, count)
  } catch {
    return 4
  }
}

function expandHome(value: string): string {
  if (value === '~') return homedir()
  if (value.startsWith('~/') || value.startsWith('~\\')) return join(homedir(), value.slice(2))
  return value
}

/** Validates that a file exists, is non-empty, and starts with the 4-byte `GGUF` magic header. */
export function isValidGgufFile(filePath: string): { valid: boolean; sizeBytes?: number } {
  try {
    const stat = statSync(filePath)
    if (!stat.isFile() || stat.size < 16) return { valid: false }
    const fd = openSync(filePath, 'r')
    try {
      const header = Buffer.alloc(4)
      const bytesRead = readSync(fd, header, 0, 4, 0)
      if (bytesRead < 4 || !header.equals(GGUF_MAGIC)) return { valid: false }
      return { valid: true, sizeBytes: stat.size }
    } finally {
      closeSync(fd)
    }
  } catch {
    return { valid: false }
  }
}

function findInDirectoryRecursive(dir: string, targetName: string, depth = 0): string | null {
  if (depth > 4 || !existsSync(dir)) return null
  let entries
  try {
    entries = readdirSync(dir, { withFileTypes: true })
  } catch {
    return null
  }
  for (const entry of entries) {
    const full = join(dir, entry.name)
    if (entry.isFile() && entry.name.toLowerCase() === targetName.toLowerCase()) {
      return full
    }
  }
  for (const entry of entries) {
    if (entry.isDirectory()) {
      const nested = findInDirectoryRecursive(join(dir, entry.name), targetName, depth + 1)
      if (nested) return nested
    }
  }
  return null
}

function findExecutableOnPath(
  name: string,
  env: NodeJS.ProcessEnv = process.env,
  platform: NodeJS.Platform = process.platform,
): string | null {
  const pathVar = env.PATH ?? env.Path ?? ''
  if (!pathVar) return null
  const sep = platform === 'win32' ? ';' : ':'
  const candidates = platform === 'win32' ? [`${name}.exe`, name] : [name]
  for (const dir of pathVar.split(sep)) {
    const cleanDir = dir.trim().replace(/^"+|"+$/g, '')
    if (!cleanDir) continue
    for (const file of candidates) {
      const full = join(cleanDir, file)
      try {
        if (statSync(full).isFile()) return full
      } catch {
        // continue
      }
    }
  }
  return null
}

export function selectLlamaReleaseAsset(
  assets: readonly LlamaReleaseAsset[],
  platform: NodeJS.Platform = process.platform,
  arch: string = process.arch,
): LlamaReleaseAsset | null {
  const patterns: RegExp[] = []
  if (platform === 'win32') {
    if (arch === 'arm64') {
      patterns.push(/bin-win-cpu-arm64\.zip$/i, /bin-win-.*arm64\.zip$/i)
    } else {
      patterns.push(/bin-win-cpu-x64\.zip$/i, /bin-win-avx2-x64\.zip$/i, /bin-win-noavx-x64\.zip$/i)
    }
  } else if (platform === 'darwin') {
    if (arch === 'arm64') {
      patterns.push(/bin-macos-arm64\.(tar\.gz|zip)$/i)
    } else {
      patterns.push(/bin-macos-x64\.(tar\.gz|zip)$/i)
    }
  } else if (platform === 'linux') {
    if (arch === 'arm64') {
      patterns.push(/bin-ubuntu-arm64\.(tar\.gz|zip)$/i, /bin-linux-arm64\.(tar\.gz|zip)$/i)
    } else {
      patterns.push(
        /bin-ubuntu-x64\.(tar\.gz|zip)$/i,
        /bin-ubuntu-vulkan-x64\.(tar\.gz|zip)$/i,
        /bin-linux-x64\.(tar\.gz|zip)$/i,
      )
    }
  }
  for (const pattern of patterns) {
    const match = assets.find((asset) => pattern.test(asset.name))
    if (match) return match
  }
  return null
}

async function freeLocalPort(): Promise<number> {
  return new Promise((resolve, reject) => {
    const server = createServer()
    server.once('error', reject)
    server.listen(0, '127.0.0.1', () => {
      const address = server.address()
      const port = typeof address === 'object' && address ? address.port : 0
      server.close((err) => {
        if (err) reject(err)
        else if (!port) reject(new Error('Could not allocate a local loopback port'))
        else resolve(port)
      })
    })
  })
}

function killServerProcess(proc: ChildProcess, platform: NodeJS.Platform = process.platform): void {
  if (!proc.pid || proc.exitCode !== null || proc.signalCode !== null) return
  if (platform === 'win32') {
    try {
      spawn('taskkill', ['/PID', String(proc.pid), '/T', '/F'], {
        stdio: 'ignore',
        windowsHide: true,
      })
      return
    } catch {
      // fall through
    }
  }
  try {
    proc.kill('SIGTERM')
  } catch {
    // ignore
  }
}

export interface LocalVlmManagerDeps {
  userDataPath: string
  getSettings?: (() => LocalVlmSettings) | undefined
  onStatusChange?: ((status: LocalVlmStatus) => void) | undefined
  fetchFn?: typeof fetch
  spawnFn?: typeof spawn
  extractArchive?: (archivePath: string, destDir: string) => Promise<void>
  allocatePort?: () => Promise<number>
  freeRamMB?: () => number
  platform?: NodeJS.Platform
  arch?: string
  env?: NodeJS.ProcessEnv
  startupTimeoutMs?: number
  idleUnloadMs?: number
}

async function defaultExtractArchive(archivePath: string, destDir: string): Promise<void> {
  mkdirSync(destDir, { recursive: true })
  await execFileAsync('tar', ['-xf', archivePath, '-C', destDir], { windowsHide: true })
}

export class LocalVlmManager implements LocalVlmRuntimeBridge {
  private readonly userDataPath: string
  private readonly getSettingsFn: () => LocalVlmSettings
  private readonly onStatusChange: ((status: LocalVlmStatus) => void) | undefined
  private readonly fetchFn: typeof fetch
  private readonly spawnFn: typeof spawn
  private readonly extractArchive: (archivePath: string, destDir: string) => Promise<void>
  private readonly allocatePort: () => Promise<number>
  private readonly freeRamMB: () => number
  private readonly platform: NodeJS.Platform
  private readonly arch: string
  private readonly env: NodeJS.ProcessEnv
  private readonly startupTimeoutMs: number
  private readonly idleUnloadMs: number

  private settingsOverride: LocalVlmSettings | null = null
  private proc: ChildProcess | null = null
  private activePort: number | undefined
  private startInflight: Promise<{ baseUrl: string; model: string }> | null = null
  private downloadInflight: Promise<LocalVlmStatus> | null = null
  private downloadAbort: AbortController | null = null
  private downloadProgress: LocalVlmDownloadProgress | undefined
  private lastError: string | undefined
  private idleTimer: ReturnType<typeof setTimeout> | null = null
  private disposed = false

  constructor(deps: LocalVlmManagerDeps) {
    this.userDataPath = deps.userDataPath
    this.getSettingsFn = deps.getSettings ?? (() => DEFAULT_LOCAL_VLM_SETTINGS)
    this.onStatusChange = deps.onStatusChange
    this.fetchFn = deps.fetchFn ?? fetch
    this.spawnFn = deps.spawnFn ?? spawn
    this.extractArchive = deps.extractArchive ?? defaultExtractArchive
    this.allocatePort = deps.allocatePort ?? freeLocalPort
    this.freeRamMB = deps.freeRamMB ?? (() => Math.round(freemem() / (1024 * 1024)))
    this.platform = deps.platform ?? process.platform
    this.arch = deps.arch ?? process.arch
    this.env = deps.env ?? process.env
    this.startupTimeoutMs = deps.startupTimeoutMs ?? DEFAULT_STARTUP_TIMEOUT_MS
    this.idleUnloadMs = deps.idleUnloadMs ?? DEFAULT_IDLE_UNLOAD_MS
  }

  private currentSettings(): LocalVlmSettings {
    return this.settingsOverride ?? this.getSettingsFn()
  }

  updateSettings(next: LocalVlmSettings): void {
    const prev = this.currentSettings()
    this.settingsOverride = next
    const pathsChanged =
      prev.modelPath !== next.modelPath ||
      prev.mmprojPath !== next.mmprojPath ||
      prev.runtimePath !== next.runtimePath ||
      prev.ctxSize !== next.ctxSize ||
      prev.threads !== next.threads ||
      prev.gpuLayers !== next.gpuLayers
    if (pathsChanged && this.proc) {
      void this.stop()
    }
    this.emitStatus()
  }

  modelsDir(): string {
    return join(this.userDataPath, 'models', 'lfm3-vl-2b')
  }

  runtimeDir(): string {
    return join(this.userDataPath, 'runtimes', 'llama')
  }

  defaultModelPath(): string {
    return join(this.modelsDir(), LOCAL_VLM_MODEL_FILENAME)
  }

  defaultMmprojPath(): string {
    return join(this.modelsDir(), LOCAL_VLM_MMPROJ_FILENAME)
  }

  resolveModelPath(settings: LocalVlmSettings = this.currentSettings()): string {
    return settings.modelPath ? expandHome(settings.modelPath) : this.defaultModelPath()
  }

  resolveMmprojPath(settings: LocalVlmSettings = this.currentSettings()): string {
    return settings.mmprojPath ? expandHome(settings.mmprojPath) : this.defaultMmprojPath()
  }

  resolveRuntimePath(settings: LocalVlmSettings = this.currentSettings()): string | undefined {
    const custom = settings.runtimePath.trim()
    if (custom) {
      const expanded = expandHome(custom)
      if (/[./\\~]/.test(expanded)) {
        try {
          if (statSync(expanded).isFile()) return expanded
        } catch {
          return undefined
        }
        return undefined
      }
      return findExecutableOnPath(expanded, this.env, this.platform) ?? undefined
    }
    const exeName = this.platform === 'win32' ? 'llama-server.exe' : 'llama-server'
    const managed = findInDirectoryRecursive(this.runtimeDir(), exeName)
    if (managed) return managed
    return findExecutableOnPath('llama-server', this.env, this.platform) ?? undefined
  }

  isRunnable(): boolean {
    const s = this.currentSettings()
    if (!s.enabled) return false
    const model = isValidGgufFile(this.resolveModelPath(s))
    const mmproj = isValidGgufFile(this.resolveMmprojPath(s))
    const runtime = this.resolveRuntimePath(s)
    return model.valid && mmproj.valid && Boolean(runtime)
  }

  shouldFallbackFromAgy(): boolean {
    const s = this.currentSettings()
    return s.enabled && s.autoFallbackFromAgy && this.isRunnable()
  }

  status(): LocalVlmStatus {
    const settings = this.currentSettings()
    const modelPath = this.resolveModelPath(settings)
    const mmprojPath = this.resolveMmprojPath(settings)
    const modelCheck = isValidGgufFile(modelPath)
    const mmprojCheck = isValidGgufFile(mmprojPath)
    const runtimePath = this.resolveRuntimePath(settings)
    const availableRamMB = this.freeRamMB()
    const installed = modelCheck.valid && mmprojCheck.valid
    const runtimeInstalled = Boolean(runtimePath)
    setLocalVlmUsable(settings.enabled && installed && runtimeInstalled)
    return {
      installed,
      modelInstalled: modelCheck.valid,
      mmprojInstalled: mmprojCheck.valid,
      runtimeInstalled,
      running: Boolean(this.proc && this.activePort && this.proc.exitCode === null),
      starting: Boolean(this.startInflight),
      ...(this.activePort ? { port: this.activePort } : {}),
      modelPath,
      mmprojPath,
      ...(runtimePath ? { runtimePath } : {}),
      ...(modelCheck.sizeBytes !== undefined ? { modelBytes: modelCheck.sizeBytes } : {}),
      ...(mmprojCheck.sizeBytes !== undefined ? { mmprojBytes: mmprojCheck.sizeBytes } : {}),
      availableRamMB,
      ramSufficient: availableRamMB >= LOCAL_VLM_MIN_RAM_MB,
      ...(this.downloadProgress ? { download: this.downloadProgress } : {}),
      ...(this.lastError ? { lastError: this.lastError } : {}),
      settings,
    }
  }

  private emitStatus(): void {
    this.onStatusChange?.(this.status())
  }

  private resetIdleTimer(): void {
    if (this.idleTimer) clearTimeout(this.idleTimer)
    if (this.idleUnloadMs <= 0) return
    this.idleTimer = setTimeout(() => {
      void this.stop()
    }, this.idleUnloadMs)
    this.idleTimer.unref?.()
  }

  async ensureRunning(signal?: AbortSignal): Promise<{ baseUrl: string; model: string }> {
    if (this.disposed) throw new Error('Local VLM runtime has been shut down')
    const settings = this.currentSettings()
    if (!settings.enabled) {
      throw new Error('Local AI (LFM3-VL-2B) is disabled in Settings')
    }
    if (this.proc && this.proc.exitCode === null && this.activePort) {
      this.resetIdleTimer()
      return {
        baseUrl: `http://127.0.0.1:${this.activePort}/v1`,
        model: LOCAL_VLM_DEFAULT_MODEL,
      }
    }
    if (!this.startInflight) {
      this.startInflight = this.startServer(settings).finally(() => {
        this.startInflight = null
        this.emitStatus()
      })
      this.emitStatus()
    }
    if (!signal) return this.startInflight
    if (signal.aborted) {
      throw Object.assign(new Error('Aborted'), { name: 'AbortError' })
    }
    return new Promise((resolve, reject) => {
      const onAbort = () => reject(Object.assign(new Error('Aborted'), { name: 'AbortError' }))
      signal.addEventListener('abort', onAbort, { once: true })
      this.startInflight!.then(
        (res) => {
          signal.removeEventListener('abort', onAbort)
          resolve(res)
        },
        (err) => {
          signal.removeEventListener('abort', onAbort)
          reject(err)
        },
      )
    })
  }

  private async startServer(
    settings: LocalVlmSettings,
  ): Promise<{ baseUrl: string; model: string }> {
    const modelPath = this.resolveModelPath(settings)
    const mmprojPath = this.resolveMmprojPath(settings)
    const runtimePath = this.resolveRuntimePath(settings)
    if (!runtimePath) {
      throw new Error(
        'llama-server runtime is not installed. Download it in Settings → AI Model → Local AI (LFM3-VL-2B).',
      )
    }
    if (!isValidGgufFile(modelPath).valid || !isValidGgufFile(mmprojPath).valid) {
      throw new Error(
        'LFM3-VL-2B Q4_K_M GGUF or mmproj file is missing. Download it in Settings → AI Model → Local AI (LFM3-VL-2B).',
      )
    }
    const port = await this.allocatePort()
    const args = [
      '--model',
      modelPath,
      '--mmproj',
      mmprojPath,
      '--host',
      '127.0.0.1',
      '--port',
      String(port),
      '--ctx-size',
      String(settings.ctxSize),
      '--alias',
      LOCAL_VLM_DEFAULT_MODEL,
    ]
    const threadCount = resolveOptimalCpuThreads(settings.threads)
    args.push('--threads', String(threadCount))
    args.push('--threads-batch', String(threadCount))
    if (settings.gpuLayers !== 0) {
      args.push('-ngl', String(settings.gpuLayers))
    }

    let stderrTail = ''
    // Blocker 4: build a minimal environment for the child process so that Cloud API keys
    // (GEMINI_API_KEY, ANTHROPIC_API_KEY, OPENAI_API_KEY …) stored in the parent env are
    // never inherited by the untrusted llama-server subprocess.
    const SAFE_ENV_KEYS = [
      'PATH',
      'Path',
      'USERPROFILE',
      'HOME',
      'TEMP',
      'TMP',
      'TMPDIR',
      'LANG',
      'LC_ALL',
      'SystemRoot',
      'WINDIR',
      'COMSPEC',
      'TERM',
    ]
    const safeEnv: NodeJS.ProcessEnv = {}
    for (const key of SAFE_ENV_KEYS) {
      if (this.env[key] !== undefined) safeEnv[key] = this.env[key]
    }
    const child = this.spawnFn(runtimePath, args, {
      stdio: ['ignore', 'pipe', 'pipe'],
      windowsHide: true,
      env: safeEnv,
    })
    child.once('error', (err: NodeJS.ErrnoException) => {
      const code = err.code ?? ''
      if (code === 'ENOENT') {
        this.lastError = `llama-server không tìm thấy tại: ${runtimePath}`
      } else if (code === 'EACCES') {
        this.lastError = `Không có quyền chạy llama-server: ${runtimePath}`
      } else {
        this.lastError = `Lỗi khởi động llama-server: ${err.message}`
      }
    })
    this.proc = child
    this.activePort = port
    this.lastError = undefined

    child.stderr?.on('data', (chunk: Buffer | string) => {
      stderrTail = `${stderrTail}${String(chunk)}`.slice(-2000)
    })
    child.once('exit', () => {
      if (this.proc === child) {
        this.proc = null
        this.activePort = undefined
        if (this.idleTimer) {
          clearTimeout(this.idleTimer)
          this.idleTimer = null
        }
        this.emitStatus()
      }
    })

    const deadline = Date.now() + this.startupTimeoutMs
    const healthUrl = `http://127.0.0.1:${port}/health`
    while (Date.now() < deadline) {
      if (child.exitCode !== null || child.signalCode !== null) {
        const msg = `llama-server exited early (code ${child.exitCode ?? child.signalCode})${stderrTail ? `: ${stderrTail.trim()}` : ''}`
        this.lastError = msg
        this.proc = null
        this.activePort = undefined
        throw new Error(msg)
      }
      try {
        const response = await this.fetchFn(healthUrl, { signal: AbortSignal.timeout(2_000) })
        if (response.ok) {
          this.resetIdleTimer()
          this.emitStatus()
          return { baseUrl: `http://127.0.0.1:${port}/v1`, model: LOCAL_VLM_DEFAULT_MODEL }
        }
      } catch {
        // still starting
      }
      await new Promise((r) => setTimeout(r, HEALTH_POLL_MS))
    }

    killServerProcess(child, this.platform)
    this.proc = null
    this.activePort = undefined
    const timeoutErr = 'Timed out waiting for local LFM3-VL-2B llama-server to become healthy'
    this.lastError = timeoutErr
    throw new Error(timeoutErr)
  }

  async stop(): Promise<void> {
    if (this.idleTimer) {
      clearTimeout(this.idleTimer)
      this.idleTimer = null
    }
    const child = this.proc
    this.proc = null
    this.activePort = undefined
    if (child) {
      killServerProcess(child, this.platform)
    }
    this.emitStatus()
  }

  async dispose(): Promise<void> {
    this.disposed = true
    this.cancelDownload()
    await this.stop()
  }

  async testConnection(
    signal?: AbortSignal,
  ): Promise<{ ok: boolean; error?: string; latencyMs?: number }> {
    const started = Date.now()
    try {
      const { baseUrl } = await this.ensureRunning(signal)
      const response = await this.fetchFn(`${baseUrl}/models`, {
        signal: signal ?? AbortSignal.timeout(10_000),
      })
      if (!response.ok) {
        return { ok: false, error: `Local llama-server returned HTTP ${response.status}` }
      }
      return { ok: true, latencyMs: Date.now() - started }
    } catch (error) {
      return { ok: false, error: error instanceof Error ? error.message : String(error) }
    }
  }

  cancelDownload(): boolean {
    if (!this.downloadAbort) return false
    this.downloadAbort.abort()
    return true
  }

  async downloadAll(options?: {
    includeRuntime?: boolean
    includeModel?: boolean
  }): Promise<LocalVlmStatus> {
    if (this.downloadInflight) return this.downloadInflight
    const includeRuntime = options?.includeRuntime ?? true
    const includeModel = options?.includeModel ?? true
    this.downloadAbort = new AbortController()
    const signal = this.downloadAbort.signal
    this.lastError = undefined

    const run = (async (): Promise<LocalVlmStatus> => {
      try {
        if (includeRuntime && !this.resolveRuntimePath()) {
          await this.downloadLlamaRuntime(signal)
        }
        if (includeModel) {
          const modelDest = this.resolveModelPath()
          if (!isValidGgufFile(modelDest).valid) {
            await this.downloadFileAtomic(LOCAL_VLM_MODEL_URL, modelDest, 'model', signal)
            if (!isValidGgufFile(modelDest).valid) {
              rmSync(modelDest, { force: true })
              throw new Error('Downloaded model file is not a valid GGUF checkpoint')
            }
          }
          const mmprojDest = this.resolveMmprojPath()
          if (!isValidGgufFile(mmprojDest).valid) {
            await this.downloadFileAtomic(LOCAL_VLM_MMPROJ_URL, mmprojDest, 'mmproj', signal)
            if (!isValidGgufFile(mmprojDest).valid) {
              rmSync(mmprojDest, { force: true })
              throw new Error('Downloaded mmproj file is not a valid GGUF projector')
            }
          }
        }
        return this.status()
      } catch (error) {
        this.lastError =
          error instanceof Error && error.name === 'AbortError'
            ? 'Download cancelled'
            : error instanceof Error
              ? error.message
              : String(error)
        throw error
      } finally {
        this.downloadProgress = undefined
        this.downloadAbort = null
        this.downloadInflight = null
        this.emitStatus()
      }
    })()

    this.downloadInflight = run
    return run
  }

  private async downloadLlamaRuntime(signal: AbortSignal): Promise<void> {
    let releaseAssets: LlamaReleaseAsset[] = []
    try {
      const res = await this.fetchFn(LOCAL_VLM_LLAMA_RELEASES_API, {
        headers: { Accept: 'application/vnd.github+json', 'User-Agent': 'GenOffice' },
        signal,
      })
      if (res.ok) {
        const release = (await res.json()) as { assets?: LlamaReleaseAsset[] }
        releaseAssets = release.assets ?? []
      }
    } catch {
      // GitHub API unreachable or rate-limited; fall back to release page / stable tag
    }

    if (releaseAssets.length === 0) {
      try {
        const htmlRes = await this.fetchFn('https://github.com/ggml-org/llama.cpp/releases', {
          headers: { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)' },
          signal,
        })
        if (htmlRes.ok) {
          const html = await htmlRes.text()
          const regex = /href="([^"]*\/download\/[^"]*\/([^"]+\.(?:zip|tar\.gz)))"/g
          let match: RegExpExecArray | null
          while ((match = regex.exec(html)) !== null) {
            const url = match[1].startsWith('http') ? match[1] : `https://github.com${match[1]}`
            const name = match[2]
            releaseAssets.push({ name, browser_download_url: url })
          }
        }
      } catch {
        // Fall back to stable hardcoded releases below
      }
    }

    if (releaseAssets.length === 0) {
      const tag = 'b12445'
      const base = `https://github.com/ggml-org/llama.cpp/releases/download/${tag}`
      releaseAssets = [
        {
          name: `llama-${tag}-bin-win-cpu-x64.zip`,
          browser_download_url: `${base}/llama-${tag}-bin-win-cpu-x64.zip`,
        },
        {
          name: `llama-${tag}-bin-win-avx2-x64.zip`,
          browser_download_url: `${base}/llama-${tag}-bin-win-avx2-x64.zip`,
        },
        {
          name: `llama-${tag}-bin-win-cpu-arm64.zip`,
          browser_download_url: `${base}/llama-${tag}-bin-win-cpu-arm64.zip`,
        },
        {
          name: `llama-${tag}-bin-macos-arm64.zip`,
          browser_download_url: `${base}/llama-${tag}-bin-macos-arm64.zip`,
        },
        {
          name: `llama-${tag}-bin-macos-x64.zip`,
          browser_download_url: `${base}/llama-${tag}-bin-macos-x64.zip`,
        },
        {
          name: `llama-${tag}-bin-ubuntu-x64.zip`,
          browser_download_url: `${base}/llama-${tag}-bin-ubuntu-x64.zip`,
        },
        {
          name: `llama-${tag}-bin-ubuntu-arm64.zip`,
          browser_download_url: `${base}/llama-${tag}-bin-ubuntu-arm64.zip`,
        },
      ]
    }

    const stableMap: Record<string, LlamaReleaseAsset> = {
      'win32-x64': {
        name: 'llama-b12445-bin-win-cpu-x64.zip',
        browser_download_url:
          'https://github.com/ggml-org/llama.cpp/releases/download/b12445/llama-b12445-bin-win-cpu-x64.zip',
      },
      'win32-arm64': {
        name: 'llama-b12445-bin-win-cpu-arm64.zip',
        browser_download_url:
          'https://github.com/ggml-org/llama.cpp/releases/download/b12445/llama-b12445-bin-win-cpu-arm64.zip',
      },
      'darwin-arm64': {
        name: 'llama-b12445-bin-macos-arm64.zip',
        browser_download_url:
          'https://github.com/ggml-org/llama.cpp/releases/download/b12445/llama-b12445-bin-macos-arm64.zip',
      },
      'darwin-x64': {
        name: 'llama-b12445-bin-macos-x64.zip',
        browser_download_url:
          'https://github.com/ggml-org/llama.cpp/releases/download/b12445/llama-b12445-bin-macos-x64.zip',
      },
      'linux-x64': {
        name: 'llama-b12445-bin-ubuntu-x64.zip',
        browser_download_url:
          'https://github.com/ggml-org/llama.cpp/releases/download/b12445/llama-b12445-bin-ubuntu-x64.zip',
      },
      'linux-arm64': {
        name: 'llama-b12445-bin-ubuntu-arm64.zip',
        browser_download_url:
          'https://github.com/ggml-org/llama.cpp/releases/download/b12445/llama-b12445-bin-ubuntu-arm64.zip',
      },
    }
    const stableFallback = stableMap[`${this.platform}-${this.arch}`]

    const asset = selectLlamaReleaseAsset(releaseAssets, this.platform, this.arch) ?? stableFallback
    if (!asset) {
      throw new Error(`No prebuilt llama.cpp binary found for ${this.platform}/${this.arch}`)
    }
    const runtimeDir = this.runtimeDir()
    const archivePath = join(runtimeDir, asset.name)
    await this.downloadFileAtomic(asset.browser_download_url, archivePath, 'runtime', signal)
    try {
      await this.extractArchive(archivePath, runtimeDir)
    } finally {
      rmSync(archivePath, { force: true })
    }
    const exeName = this.platform === 'win32' ? 'llama-server.exe' : 'llama-server'
    const exe = findInDirectoryRecursive(runtimeDir, exeName)
    if (!exe) {
      throw new Error(`Extracted llama.cpp archive did not contain ${exeName}`)
    }
    if (this.platform !== 'win32') {
      try {
        chmodSync(exe, 0o755)
      } catch {
        // ignore
      }
    }
  }

  private async downloadFileAtomic(
    url: string,
    destPath: string,
    stage: LocalVlmDownloadProgress['stage'],
    signal: AbortSignal,
  ): Promise<void> {
    mkdirSync(dirname(destPath), { recursive: true })
    const partPath = `${destPath}.part`
    rmSync(partPath, { force: true })
    const fileLabel = destPath.split(/[\\/]/).pop() ?? destPath

    const response = await this.fetchFn(url, {
      headers: { 'User-Agent': 'GenOffice' },
      redirect: 'follow',
      signal,
    })
    if (!response.ok || !response.body) {
      throw new Error(`Download failed for ${fileLabel} (HTTP ${response.status})`)
    }
    const totalHeader = response.headers.get('content-length')
    const totalBytes = totalHeader && Number(totalHeader) > 0 ? Number(totalHeader) : undefined
    let receivedBytes = 0
    this.downloadProgress = {
      active: true,
      stage,
      file: fileLabel,
      receivedBytes: 0,
      ...(totalBytes !== undefined ? { totalBytes, percent: 0 } : {}),
    }
    this.emitStatus()

    const writer = createWriteStream(partPath)
    const reader = response.body.getReader()
    let lastEmit = 0
    try {
      for (;;) {
        const { done, value } = await reader.read()
        if (done) break
        if (value && value.byteLength > 0) {
          receivedBytes += value.byteLength
          await new Promise<void>((resolve, reject) => {
            if (!writer.write(Buffer.from(value), (err) => (err ? reject(err) : resolve()))) {
              writer.once('drain', resolve)
            }
          })
          const now = Date.now()
          if (now - lastEmit >= 200) {
            lastEmit = now
            this.downloadProgress = {
              active: true,
              stage,
              file: fileLabel,
              receivedBytes,
              ...(totalBytes !== undefined
                ? {
                    totalBytes,
                    percent: Math.min(100, Math.round((receivedBytes / totalBytes) * 100)),
                  }
                : {}),
            }
            this.emitStatus()
          }
        }
      }
      await new Promise<void>((resolve, reject) => {
        writer.end((err?: Error | null) => (err ? reject(err) : resolve()))
      })
      renameSync(partPath, destPath)
    } catch (error) {
      writer.destroy()
      rmSync(partPath, { force: true })
      throw error
    }
  }
}
