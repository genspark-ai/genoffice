/**
 * Wrapper around gsk (Genspark CLI, @genspark/cli) — search / image generation /
 * media analysis / upload / transcription.
 *
 * Execution: the main process spawns the CLI's JS entry with
 * ELECTRON_RUN_AS_NODE=1 so Electron itself acts as Node (the variable is a
 * no-op in a plain Node environment), avoiding the Windows problem where
 * .cmd files cannot be passed to execFile.
 *
 * Auth: gskApiKey() below. When not logged in, hasGskAuth() returns false and
 * callers should fall back to other implementations (e.g. Serper search).
 */

import { execFile } from 'node:child_process'
import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import { homedir } from 'node:os'
import { join } from 'node:path'
import { createRequire } from 'node:module'
import {
  COPYRIGHT_HOSTS,
  asRecord,
  firstItem,
  safeHost,
  type ImageSearchResult,
  type WebSearchResult,
} from './shared'
import { genofficeApiKey } from './genoffice-auth'

const SEARCH_TIMEOUT_MS = 60_000
const GENERATE_TIMEOUT_MS = 600_000
const MAX_BUFFER = 32 * 1024 * 1024

// ── CLI resolution & auth ───────────────────────────────────────────

let cachedEntry: string | null | undefined

/** JS entry of @genspark/cli (null if not found). Can be overridden via GSK_CLI_PATH. */
export function resolveGskEntry(): string | null {
  if (process.env.GSK_CLI_PATH) return process.env.GSK_CLI_PATH
  if (cachedEntry !== undefined) return cachedEntry
  try {
    const require = createRequire(import.meta.url)
    cachedEntry = require.resolve('@genspark/cli/dist/index.js')
  } catch {
    // The packaged app has no node_modules; electron-builder extraResources
    // copies the CLI into Resources/gsk/
    const resourcesPath = (process as { resourcesPath?: string }).resourcesPath
    const packed = resourcesPath
      ? join(resourcesPath, 'gsk', 'node_modules', '@genspark', 'cli', 'dist', 'index.js')
      : null
    cachedEntry = packed && existsSync(packed) ? packed : null
  }
  return cachedEntry
}

let compatPath: string | null | undefined

/**
 * Under Electron-as-Node, Commander.js sees process.versions.electron and
 * parses argv with Electron conventions (skipping one argument fewer), so it
 * treats the entry path as a command and prints help. Workaround: --require a
 * script that deletes versions.electron. In a plain Node environment
 * (tests/tsx) no handling is needed and an empty array is returned.
 */
function electronCompatArgs(): string[] {
  if (!process.versions.electron) return []
  if (compatPath === undefined) {
    try {
      const dir = join(homedir(), '.genoffice', 'bin')
      mkdirSync(dir, { recursive: true })
      compatPath = join(dir, 'electron-compat.js')
      writeFileSync(compatPath, 'delete process.versions.electron;\n')
    } catch {
      compatPath = null
    }
  }
  return compatPath ? ['--require', compatPath] : []
}

/**
 * API key for Genspark LLM proxy / tool_cli auth; '' when not logged in.
 * Priority: GSK_API_KEY env → GenOffice's own key (bills to us via its
 * key_name) → shared gsk CLI login (bills to the Claw bucket).
 */
export function gskApiKey(): string {
  if (process.env.GSK_API_KEY) return process.env.GSK_API_KEY
  const own = genofficeApiKey()
  if (own) return own
  try {
    const configPath = join(homedir(), '.genspark-tool-cli', 'config.json')
    if (!existsSync(configPath)) return ''
    const config = JSON.parse(readFileSync(configPath, 'utf-8')) as { api_key?: string }
    return config.api_key ?? ''
  } catch {
    return ''
  }
}

/**
 * Whether gsk is usable (CLI installed and logged in / has a key). Callers use this to decide fallback.
 * Set AI_SEARCH_DISABLE_GSK=1 to force-disable (test isolation / force Serper).
 */
export function hasGskAuth(): boolean {
  if (process.env.AI_SEARCH_DISABLE_GSK === '1') return false
  return !!gskApiKey() && resolveGskEntry() !== null
}

// ── Child-process proxy plumbing ────────────────────────────────────

let explicitProxyUrl = ''

/**
 * Proxy for spawned gsk CLI children. The Electron main process routes its own
 * fetch through undici's global dispatcher (see the apps' proxy bootstraps),
 * but that never reaches child processes: a packaged app launched from
 * Finder/Explorer inherits no proxy env vars, so `gsk login` etc. would dial
 * genspark.ai directly and fail on networks that require the system proxy
 * (browser + renderer work, only the CLI children break). The bootstraps call
 * this with the proxy they resolved so gskChildEnv() can forward it.
 */
export function setGskProxyUrl(url: string): void {
  explicitProxyUrl = url
}

/**
 * env for gsk CLI children: Electron-as-Node plus proxy forwarding. The CLI
 * uses Node's built-in fetch, which ignores proxy env vars unless
 * NODE_USE_ENV_PROXY=1 (Node >= 24 — Electron 41 ships Node 24). Only
 * http(s):// proxies are forwarded; undici's env proxy cannot speak SOCKS.
 */
export function gskChildEnv(base: NodeJS.ProcessEnv = process.env): NodeJS.ProcessEnv {
  const env: NodeJS.ProcessEnv = { ...base, ELECTRON_RUN_AS_NODE: '1' }
  const proxy = [
    explicitProxyUrl,
    base.HTTPS_PROXY,
    base.https_proxy,
    base.HTTP_PROXY,
    base.http_proxy,
    base.ALL_PROXY,
    base.all_proxy,
  ].find((v) => v && /^https?:\/\//.test(v))
  if (proxy) {
    // scrub inherited variants: undici's env proxy prefers the lowercase
    // names, so a leftover https_proxy/all_proxy would override the selection
    delete env.https_proxy
    delete env.http_proxy
    delete env.ALL_PROXY
    delete env.all_proxy
    env.NODE_USE_ENV_PROXY = '1'
    env.HTTPS_PROXY = proxy
    env.HTTP_PROXY = proxy
  }
  return env
}

// ── Low-level execution ─────────────────────────────────────────────

/**
 * gsk output may have [INFO] log lines mixed in before the JSON; scan from the
 * end for the first line starting with { or [ and parse from there.
 */
export function parseGskOutput(stdout: string): unknown {
  const trimmed = stdout.trim()
  try {
    return JSON.parse(trimmed)
  } catch {
    /* fall through to line-by-line scan */
  }
  const lines = trimmed.split('\n')
  for (let i = lines.length - 1; i >= 0; i--) {
    const line = lines[i]!.trim()
    if (line.startsWith('{') || line.startsWith('[')) {
      try {
        return JSON.parse(lines.slice(i).join('\n'))
      } catch {
        continue
      }
    }
  }
  throw new Error(`No JSON found in gsk output: ${stdout.slice(0, 300)}`)
}

function runGsk(args: string[], timeoutMs: number, signal?: AbortSignal): Promise<unknown> {
  const entry = resolveGskEntry()
  if (!entry) return Promise.reject(new Error('@genspark/cli is not installed'))
  // inject the resolved key so the CLI bills the same identity as our direct HTTP calls
  const key = gskApiKey()
  return new Promise((resolve, reject) => {
    execFile(
      process.execPath,
      [...electronCompatArgs(), entry, ...args, '--output', 'json'],
      {
        timeout: timeoutMs,
        maxBuffer: MAX_BUFFER,
        env: {
          ...gskChildEnv(),
          ...(key ? { GSK_API_KEY: key } : {}),
        },
        ...(signal ? { signal } : {}),
      },
      (err, stdout, stderr) => {
        if (err) {
          // gsk's real failure reason (auth/network/quota) is in stderr; append it to ease debugging
          const errText = (stderr || '').toString().trim().slice(0, 500)
          reject(errText ? new Error(`${err.message} | stderr: ${errText}`) : err)
          return
        }
        try {
          const result = parseGskOutput(String(stdout))
          const rec = asRecord(result)
          if (rec.status && rec.status !== 'ok') {
            reject(new Error(`gsk returned an error: ${rec.message ?? rec.status}`))
            return
          }
          resolve(result)
        } catch (parseErr) {
          reject(parseErr)
        }
      },
    )
  })
}

// ── Search ──────────────────────────────────────────────────────────

/** Parses the `gsk search` response shape data.organic_results[{title,link,snippet}] (exported for tests) */
export function parseGskWebSearch(
  raw: unknown,
  maxResults: number,
): { results: WebSearchResult[]; answer?: string } {
  const data = asRecord(asRecord(raw).data ?? raw)
  const organic: unknown[] = Array.isArray(data.organic_results) ? data.organic_results : []
  const results: WebSearchResult[] = organic.slice(0, maxResults).map((item) => {
    const o = asRecord(item)
    return {
      title: String(o.title ?? ''),
      url: String(o.link ?? ''),
      snippet: String(o.snippet ?? ''),
    }
  })
  const answer = typeof data.answer === 'string' && data.answer ? data.answer : undefined
  return answer !== undefined ? { results, answer } : { results }
}

export async function gskWebSearch(
  query: string,
  maxResults = 6,
): Promise<{ results: WebSearchResult[]; answer?: string }> {
  const raw = await runGsk(['search', query], SEARCH_TIMEOUT_MS)
  return parseGskWebSearch(raw, maxResults)
}

/** Parses the `gsk img-search` response shape data[{image_url,title,source,link,width,height}] (exported for tests) */
export function parseGskImageSearch(raw: unknown, maxResults: number): ImageSearchResult[] {
  const dataRaw = asRecord(raw).data
  const data: unknown[] = Array.isArray(dataRaw) ? dataRaw : []
  const images: ImageSearchResult[] = []
  for (const item of data) {
    const img = asRecord(item)
    const imageUrl = String(img.image_url ?? img.imageUrl ?? '')
    if (!imageUrl) continue
    if (COPYRIGHT_HOSTS.some((d) => imageUrl.toLowerCase().includes(d))) continue
    const width = Number(img.width)
    const height = Number(img.height)
    const entry: ImageSearchResult = {
      title: String(img.title ?? ''),
      imageUrl,
      sourceUrl: String(img.link ?? ''),
      source: String(img.source ?? safeHost(img.link)),
    }
    if (Number.isFinite(width) && width > 0) entry.width = width
    if (Number.isFinite(height) && height > 0) entry.height = height
    images.push(entry)
    if (images.length >= maxResults) break
  }
  return images
}

export async function gskImageSearch(query: string, maxResults = 8): Promise<ImageSearchResult[]> {
  const raw = await runGsk(['img-search', query], SEARCH_TIMEOUT_MS)
  return parseGskImageSearch(raw, maxResults)
}

// ── Image generation ────────────────────────────────────────────────

export interface GskGenerateImageOptions {
  /** Image description (English works better; text that must appear in the image stays verbatim) */
  prompt: string
  /** Generation model; defaults to the CLI default (nano-banana-2). Special-purpose: fal-bria-rmbg background removal, fal-ai/recraft-clarity-upscale upscaling, flux-pro/outpaint outpainting, fal-ai/image-editing/text-removal watermark text removal */
  model?: string
  /** Reference/edit-target image URLs (local paths also supported; the CLI uploads them automatically) */
  referenceImageUrls?: string[]
  /** 1:1 | 4:3 | 16:9 | 9:16 | 3:4 | 2:3 | 3:2 | auto */
  aspectRatio?: string
  /** auto | 0.5k | 1k | 2k | 3k | 4k */
  imageSize?: string
}

export interface GskGeneratedImage {
  /** Watermark-free image URL (preferred); can be downloaded and inserted directly */
  url: string
  taskId: string
}

/** Parses `gsk img` results: prefers data.generated_images[].image_urls_nowatermark (exported for tests) */
export function parseGskGeneratedImage(raw: unknown): GskGeneratedImage {
  const data = asRecord(asRecord(raw).data ?? raw)
  const img = asRecord(firstItem(data.generated_images ?? data.images ?? []))
  const url = firstItem(img.image_urls_nowatermark) ?? firstItem(img.image_urls) ?? img.url ?? ''
  if (!url) {
    throw new Error(`gsk image generation returned no result: ${JSON.stringify(raw).slice(0, 200)}`)
  }
  return { url: String(url), taskId: String(img.task_id ?? '') }
}

export async function gskGenerateImage(
  options: GskGenerateImageOptions,
  signal?: AbortSignal,
): Promise<GskGeneratedImage> {
  const args = ['img', options.prompt]
  if (options.model) args.push('-m', options.model)
  if (options.referenceImageUrls?.length) args.push('--image_urls', ...options.referenceImageUrls)
  if (options.aspectRatio) args.push('--aspect_ratio', options.aspectRatio)
  if (options.imageSize) args.push('--image_size', options.imageSize)
  const raw = await runGsk(args, GENERATE_TIMEOUT_MS, signal)
  const result = parseGskGeneratedImage(raw)
  // Bare genspark file URLs (/api/files/) return 403; swap for a signed direct link with a token
  // so later plain fetch downloads (e.g. insert_web_image) need no auth
  result.url = await gskResolveDownloadUrl(result.url)
  return result
}

/**
 * Swaps a genspark file wrapper URL for an anonymously downloadable signed direct link (`gsk download`).
 * Non-genspark file URLs are returned as-is; on swap failure the URL is also returned as-is
 * (callers handle download failures themselves).
 */
export async function gskResolveDownloadUrl(url: string): Promise<string> {
  if (!url.includes('/api/files/')) return url
  try {
    const raw = asRecord(await runGsk(['download', url], SEARCH_TIMEOUT_MS))
    const downloadUrl = asRecord(raw.data).download_url ?? raw.download_url
    return downloadUrl ? String(downloadUrl) : url
  } catch {
    return url
  }
}

// ── Cloud single-slide generation (Genspark slide_generate) ─────────

/**
 * Calls the tool_cli HTTP endpoint directly so structured params
 * (deck_context/images) are sent as JSON rather than through the CLI's
 * string-based argument passing.
 */
const GSK_TOOL_CLI_BASE = 'https://www.genspark.ai/api/tool_cli'
const SLIDE_GENERATE_TIMEOUT_MS = 240_000

export interface GskSlideGenerateOptions {
  /** Content and layout brief for this page */
  brief: string
  title?: string
  /** Deck-level visual system / typography / palette rules (Style Skill text) */
  styleSkill?: string
  /** Deck topic, page index, total pages, neighboring-page context */
  deckContext?: Record<string, unknown>
  /** HTTPS image candidates for this page */
  images?: { url: string; caption?: string }[]
  width?: number
  height?: number
  /** Slides model tier: ultra = opus-class model, standard (server default) = lighter model */
  tier?: 'standard' | 'ultra'
  signal?: AbortSignal
}

/** Final NDJSON line = the result; heartbeat lines have no status field */
export function parseToolCliNdjson(text: string): Record<string, unknown> {
  const lines = text.trim().split('\n')
  for (let i = lines.length - 1; i >= 0; i--) {
    const line = lines[i]!.trim()
    if (!line.startsWith('{')) continue
    try {
      const obj: unknown = JSON.parse(line)
      if (obj && typeof obj === 'object' && 'status' in obj) return obj as Record<string, unknown>
    } catch {
      continue
    }
  }
  throw new Error(`No result line in tool_cli response: ${text.slice(0, 300)}`)
}

async function toolCliPost(
  path: string,
  body: unknown,
  timeoutMs: number,
  signal?: AbortSignal,
): Promise<unknown> {
  const key = gskApiKey()
  if (!key) throw new Error('Not logged in to Genspark (gsk login)')
  const controller = new AbortController()
  const timer = setTimeout(() => controller.abort(), timeoutMs)
  const onAbort = () => controller.abort()
  signal?.addEventListener('abort', onAbort, { once: true })
  try {
    const resp = await fetch(`${GSK_TOOL_CLI_BASE}${path}`, {
      method: 'POST',
      // X-Agent-Type splits GenOffice usage out of the proxy's "Claw" billing bucket
      headers: {
        'X-Api-Key': key,
        'Content-Type': 'application/json',
        'X-Agent-Type': 'genoffice',
      },
      body: JSON.stringify(body),
      signal: controller.signal,
    })
    const text = await resp.text()
    if (!resp.ok) throw new Error(`tool_cli ${path} HTTP ${resp.status}: ${text.slice(0, 200)}`)
    const result = parseToolCliNdjson(text)
    if (result.status !== 'ok') {
      throw new Error(`tool_cli ${path} failed: ${result.message ?? result.status}`)
    }
    return result.data
  } finally {
    clearTimeout(timer)
    signal?.removeEventListener('abort', onAbort)
  }
}

/**
 * Generates one editable slide in the cloud (brief → HTML → one-slide PPTX)
 * and returns the downloaded PPTX bytes plus the model that produced it.
 */
export async function gskSlideGenerate(
  options: GskSlideGenerateOptions,
): Promise<{ bytes: Uint8Array; model: string }> {
  const { signal } = options
  const body: Record<string, unknown> = { brief: options.brief }
  if (options.title) body.title = options.title
  if (options.styleSkill) body.style_skill = options.styleSkill
  if (options.deckContext) body.deck_context = options.deckContext
  if (options.images?.length) body.images = options.images
  if (options.width) body.width = options.width
  if (options.height) body.height = options.height
  if (options.tier) body.tier = options.tier
  const data = asRecord(
    await toolCliPost('/slide_generate', body, SLIDE_GENERATE_TIMEOUT_MS, signal),
  )
  const pptxUrl = data.pptx_url
  if (!pptxUrl)
    throw new Error(`slide_generate returned no pptx_url: ${JSON.stringify(data).slice(0, 200)}`)
  // The bare file-wrapper URL 403s on plain GET; swap for a signed direct link first
  const dl = asRecord(
    await toolCliPost('/file/download', { file_wrapper_url: pptxUrl }, SEARCH_TIMEOUT_MS, signal),
  )
  const downloadUrl = dl.download_url
  if (!downloadUrl) throw new Error('file/download returned no download_url')
  const resp = await fetch(String(downloadUrl), signal ? { signal } : undefined)
  if (!resp.ok) throw new Error(`PPTX download failed: HTTP ${resp.status}`)
  return { bytes: new Uint8Array(await resp.arrayBuffer()), model: String(data.model ?? '') }
}

// ── File conversion (PDF → DOCX) ────────────────────────────────────

/** Extracts the download link from file_convert's markdown result text (exported for tests) */
export function parseGskConvertResult(raw: unknown): string {
  const data = asRecord(asRecord(raw).data ?? raw)
  const text = typeof data.result === 'string' ? data.result : ''
  const url = /\((https?:\/\/[^)\s]+)\)/.exec(text)?.[1] ?? /https?:\/\/\S+/.exec(text)?.[0]
  if (!url) {
    throw new Error(`file_convert returned no link: ${JSON.stringify(raw).slice(0, 200)}`)
  }
  return url
}

/**
 * Uploads a local PDF and converts it to DOCX in the cloud (`gsk convert`,
 * costs 5 credits); returns the DOCX bytes.
 */
export async function gskConvertPdfToDocx(
  filePath: string,
  signal?: AbortSignal,
): Promise<Uint8Array> {
  const wrapperUrl = await gskUpload(filePath)
  const raw = await runGsk(['convert', wrapperUrl], GENERATE_TIMEOUT_MS, signal)
  const link = parseGskConvertResult(raw)
  const downloadUrl = await gskResolveDownloadUrl(link)
  const resp = await fetch(downloadUrl, signal ? { signal } : undefined)
  if (!resp.ok) throw new Error(`DOCX download failed: HTTP ${resp.status}`)
  return new Uint8Array(await resp.arrayBuffer())
}

// ── Media analysis / transcription ──────────────────────────────────

/** Best-effort text extraction from gsk analysis-type command output (shape varies by task, so be lenient; exported for tests) */
export function extractGskText(raw: unknown): string {
  const data = asRecord(raw).data ?? raw
  if (typeof data === 'string') return data
  if (data && typeof data === 'object') {
    for (const key of [
      'analysis',
      'analysis_result',
      'result',
      'content',
      'text',
      'answer',
      'transcript',
    ]) {
      const v = (data as Record<string, unknown>)[key]
      if (typeof v === 'string' && v.trim()) return v
    }
  }
  return JSON.stringify(data ?? raw)
}

export interface GskAnalyzeMediaOptions {
  /** Media URLs or local paths (image/audio/video; the CLI uploads local files automatically) */
  mediaUrls: string[]
  /** Analysis requirements (in English): what info to extract and what it's for */
  requirements: string
}

export async function gskAnalyzeMedia(
  options: GskAnalyzeMediaOptions,
  signal?: AbortSignal,
): Promise<string> {
  const args = [
    'media-analyze',
    '--media_urls',
    ...options.mediaUrls,
    '--requirements',
    options.requirements,
  ]
  const raw = await runGsk(args, GENERATE_TIMEOUT_MS, signal)
  return extractGskText(raw)
}

export interface GskTranscribeOptions {
  /** Audio URLs or local paths */
  audioUrls: string[]
  /** Prompt (context / proper nouns; can improve recognition quality) */
  prompt?: string
  /** whisper-1 (default, with timestamps) | gemini-3-flash-preview | elevenlabs_scribe_v2 */
  model?: string
}

export async function gskTranscribe(
  options: GskTranscribeOptions,
  signal?: AbortSignal,
): Promise<string> {
  const args = ['transcribe', '--audio_urls', ...options.audioUrls]
  if (options.prompt) args.push('--prompt', options.prompt)
  if (options.model) args.push('-m', options.model)
  const raw = await runGsk(args, GENERATE_TIMEOUT_MS, signal)
  return extractGskText(raw)
}

// ── File upload / login ─────────────────────────────────────────────

/** Uploads a local file and returns a file wrapper URL (usable as input to other gsk commands) */
export async function gskUpload(filePath: string): Promise<string> {
  const raw = asRecord(await runGsk(['upload', filePath], GENERATE_TIMEOUT_MS))
  const dataRec = asRecord(raw.data)
  const url = dataRec.file_wrapper_url ?? raw.url ?? dataRec.url
  if (!url) throw new Error(`gsk upload did not return a URL: ${JSON.stringify(raw).slice(0, 200)}`)
  return String(url)
}

export interface GskLoginInfo {
  email: string
  plan: string
  creditBalance?: number
}

/** Current login info; null when not logged in */
export async function gskLoginInfo(): Promise<GskLoginInfo | null> {
  try {
    const raw = await runGsk(['login-info'], SEARCH_TIMEOUT_MS)
    const d = asRecord(asRecord(raw).data ?? raw)
    if (!d.email) return null
    const info: GskLoginInfo = {
      email: String(d.email),
      plan: String(d.plan ?? d.personal_plan ?? ''),
    }
    const balance = Number(d.credit_balance)
    if (Number.isFinite(balance)) info.creditBalance = balance
    return info
  } catch {
    return null
  }
}
