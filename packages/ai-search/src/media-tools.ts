/**
 * generate_image / analyze_media for the five editors' main processes: one
 * place that reads ai-settings.json live, routes to the BYOK media provider
 * when one is configured, and otherwise to the Genspark CLI behind the usual
 * login + cloud-tools gate. BYOK providers answer with bytes; those land in
 * the local generated-image store and come back as a file:// URL that the
 * insert pipelines' fetchRemoteImage accepts.
 */

import { existsSync, readFileSync, realpathSync, statSync } from 'node:fs'
import { basename, dirname, extname, sep } from 'node:path'
import {
  activeMediaConfig,
  analyzeMediaWithProvider,
  cloudToolsEnabled,
  defaultAiSettings,
  generateImageWithProvider,
  resolveAiSettings,
  type AiSettings,
  type LegacyAiSettings,
  type MediaBlob,
} from '@genoffice/ai-provider'
// deep imports: the package root re-exports Electron-bound modules, and this file also runs in the genoffice CLI
import { readGeneratedImage, storeGeneratedImage } from '@genoffice/electron-utils/generated-images'
import {
  ResponseTooLargeError,
  fetchRemoteImage,
  readBodyCapped,
} from '@genoffice/electron-utils/remote-image'
import { fetchWithSsrfGuard } from '@genoffice/electron-utils/safe-remote-url'
import { gskAnalyzeMedia, gskGenerateImage, hasGskAuth, type GskGenerateImageOptions } from './gsk'

export const GSK_NOT_LOGGED_IN_ERROR =
  'Genspark account is not logged in on this machine; ask the user to log in first'
export const GSK_TOOLS_OFF_ERROR =
  'Genspark cloud tools are turned off in Settings (AI Model); enable them or configure an image provider under Settings (AI Media) to use this tool'

/** 200 MB: enough for a long clip through the Gemini Files API, small enough to hold in memory */
const MAX_MEDIA_BYTES = 200 * 1024 * 1024

/** Per-request ceiling across every reference of one tool call: MAX_MEDIA_BYTES bounds a single
 *  item, and without a total the same cap could be multiplied by the item count. */
const MAX_MEDIA_TOTAL_BYTES = 200 * 1024 * 1024

/** Per-request item ceiling: a media tool call is a handful of references, never a data dump. */
const MAX_MEDIA_ITEMS = 12

/** How many references are decoded at once, so peak memory is a small multiple of the total cap
 *  rather than the whole request. */
const MEDIA_LOAD_CONCURRENCY = 3

/** The production budget, exported so callers and tests can reason about the ceilings. */
export const MEDIA_BUDGET = {
  maxItems: MAX_MEDIA_ITEMS,
  maxItemBytes: MAX_MEDIA_BYTES,
  maxTotalBytes: MAX_MEDIA_TOTAL_BYTES,
  concurrency: MEDIA_LOAD_CONCURRENCY,
} as const

/** the only load failure that may hand the request back to Genspark; validation failures never do */
export class MediaTooLargeError extends Error {}

/** Subclasses MediaTooLargeError so the existing "too big, try Genspark" fallback applies. */
export class MediaBudgetExceededError extends MediaTooLargeError {}

/** The item ceiling in one place, so the Genspark route rejects like the BYOK one. */
function assertMediaItemCount(count: number, maxItems: number): void {
  if (count > maxItems) {
    throw new MediaBudgetExceededError(
      `Too many media items in one request (${count}, limit ${maxItems}); analyze them in smaller batches`,
    )
  }
}

const MIME_BY_EXT: Record<string, string> = {
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.jpeg': 'image/jpeg',
  '.gif': 'image/gif',
  '.webp': 'image/webp',
  '.bmp': 'image/bmp',
  '.heic': 'image/heic',
  '.mp4': 'video/mp4',
  '.mov': 'video/quicktime',
  '.webm': 'video/webm',
  '.m4v': 'video/mp4',
  '.mkv': 'video/x-matroska',
  '.mp3': 'audio/mpeg',
  '.wav': 'audio/wav',
  '.m4a': 'audio/mp4',
  '.aac': 'audio/aac',
  '.ogg': 'audio/ogg',
  '.flac': 'audio/flac',
}

const DATA_URL_RE = /^data:([^;,]+);base64,([\s\S]*)$/

/** Decoded size a data URL will produce, computed from the base64 length alone so the check runs
 *  before anything is allocated. Whitespace inside the payload only inflates the estimate. */
function dataUrlDecodedSize(ref: string): number {
  if (!ref.startsWith('data:')) return 0
  const match = DATA_URL_RE.exec(ref)
  if (!match) return 0
  const b64 = match[2] ?? ''
  return Math.floor((b64.length * 3) / 4)
}

export function readAiSettingsFile(path: string): AiSettings {
  let stored: Partial<AiSettings> & LegacyAiSettings = {}
  try {
    if (existsSync(path)) stored = JSON.parse(readFileSync(path, 'utf-8'))
  } catch {
    /* corrupted settings file: defaults */
  }
  return resolveAiSettings(stored, defaultAiSettings())
}

type Gate = { error: string } | null

/** no allowlist configured: the extension check alone decides */
const NO_MEDIA_ROOTS: readonly string[] = []

/**
 * Directories a bare local media path may be read from. Deliberately no
 * default: cwd is "/" in a packaged macOS app and tmpdir() is world-readable,
 * so every host opts in with its own list (empty = extension check only).
 */
export function localMediaRoots(...roots: (string | undefined | null)[]): string[] {
  return roots.filter((root): root is string => Boolean(root))
}

/** The Electron allowlist: the open document's directory plus the app's attachment staging dir. */
export function documentMediaRoots(
  docPath: string | undefined | null,
  attachmentDir: string | undefined | null,
): string[] {
  return localMediaRoots(docPath ? dirname(docPath) : undefined, attachmentDir)
}

/**
 * Real path of `ref`, or null when missing or outside `roots`. Both sides go
 * through realpath so a planted symlink cannot walk out, and containment is
 * on whole segments so /tmp/root-evil does not pass as /tmp/root.
 */
function verifiedLocalPath(ref: string, roots: readonly string[]): string | null {
  let real: string
  try {
    real = realpathSync(ref)
  } catch {
    return null
  }
  if (roots.length === 0) return real
  for (const root of roots) {
    let realRoot: string
    try {
      realRoot = realpathSync(root)
    } catch {
      continue
    }
    const prefix = realRoot.endsWith(sep) ? realRoot : `${realRoot}${sep}`
    if (real === realRoot || real.startsWith(prefix)) return real
  }
  return null
}

/** True when `ref` is a readable path; with no `roots` the allowlist does not apply. */
export function isLocalMediaPathAllowed(
  ref: string,
  roots: readonly string[] = NO_MEDIA_ROOTS,
): boolean {
  return verifiedLocalPath(ref, roots) !== null
}

/** the Genspark route's preconditions; null when it may proceed */
function gskGate(settings: AiSettings, notLoggedInError: string): Gate {
  if (!hasGskAuth()) return { error: notLoggedInError }
  if (!cloudToolsEnabled(settings)) return { error: GSK_TOOLS_OFF_ERROR }
  return null
}

function errorText(err: unknown): string {
  return err instanceof Error ? err.message : String(err)
}

/**
 * Resolves a tool-supplied media reference to bytes: an https URL (SSRF-guarded),
 * a file:// URL from the generated-image store, or a local media file
 * (attachments). Only media extensions are read locally, and only from `roots`
 * when the caller supplied any, so the model cannot ship arbitrary files to a vendor.
 */
export async function loadMediaReference(
  ref: string,
  roots: readonly string[] = NO_MEDIA_ROOTS,
): Promise<MediaBlob> {
  if (/^https?:\/\//i.test(ref)) {
    const resp = await (ref.match(/\.(png|jpe?g|gif|webp)(\?|$)/i)
      ? fetchRemoteImage(ref)
      : fetchWithSsrfGuard(ref, { headers: { 'User-Agent': 'Mozilla/5.0' } }))
    if (!resp || !resp.ok) throw new Error(`Could not download ${ref}`)
    let bytes: Uint8Array
    try {
      bytes = await readBodyCapped(resp, MAX_MEDIA_BYTES)
    } catch (err) {
      if (err instanceof ResponseTooLargeError) {
        throw new MediaTooLargeError(`${ref} is too large to analyze`)
      }
      throw err
    }
    const rawCt = resp.headers.get('content-type')?.split(';')[0]?.trim().toLowerCase()
    const ct = rawCt && rawCt !== 'application/octet-stream' ? rawCt : undefined
    const name = basename(new URL(ref).pathname) || undefined
    const mime =
      ct && ct !== 'application/octet-stream' ? ct : MIME_BY_EXT[extname(name ?? '').toLowerCase()]
    if (!mime) throw new Error(`Could not tell the media type of ${ref}`)
    return { bytes, mime, ...(name ? { name } : {}) }
  }
  // data URLs (pictures embedded in a document) carry their own bytes and type
  if (ref.startsWith('data:')) {
    const m = DATA_URL_RE.exec(ref)
    if (!m) throw new Error('Unsupported data URL: only base64-encoded media can be analyzed')
    const [, mime = '', b64 = ''] = m
    const bytes = new Uint8Array(Buffer.from(b64.replace(/\s+/g, ''), 'base64'))
    if (bytes.byteLength > MAX_MEDIA_BYTES) {
      throw new MediaTooLargeError('data URL is too large to analyze')
    }
    return { bytes, mime: mime.toLowerCase() }
  }
  if (ref.startsWith('file:')) {
    const local = readGeneratedImage(ref)
    if (!local) throw new Error(`Not an accessible image: ${ref}`)
    return { bytes: new Uint8Array(local.bytes), mime: local.mime }
  }
  const mime = MIME_BY_EXT[extname(ref).toLowerCase()]
  if (!mime) throw new Error(`Unsupported media file: ${ref} (images, video and audio only)`)
  if (!existsSync(ref)) throw new Error(`File not found: ${ref}`)
  // resolve before the read: readFileSync would follow a symlink out of the allowlist
  const local = verifiedLocalPath(ref, roots)
  if (!local) {
    throw new Error(
      `Refusing to read a local media file outside the allowed directories: ${ref} (allowed: ${roots.join(', ')})`,
    )
  }
  const stat = statSync(local)
  if (!stat.isFile()) throw new Error(`Not a regular file: ${ref}`)
  if (stat.size > MAX_MEDIA_BYTES) {
    throw new MediaTooLargeError(`${ref} is too large to analyze`)
  }
  return { bytes: new Uint8Array(readFileSync(local)), mime, name: basename(ref) }
}

export interface MediaToolOptions {
  /** localized replacement for the default signed-out message */
  notLoggedInError?: string
  /** directories a bare local path may be read from; omitted = extension check only */
  mediaRoots?: readonly string[]
}

export interface MediaBudget {
  maxItems?: number
  maxItemBytes?: number
  maxTotalBytes?: number
  concurrency?: number
}

/**
 * Load a tool call's references under a total byte budget and an item cap. Encoded data URLs are
 * measured before they are decoded, the references are loaded with bounded concurrency, and the
 * running total is re-checked as each one lands, so a request can no longer hold N times the
 * per-file cap in memory at once.
 */
export async function loadMediaReferences(
  refs: readonly string[],
  budget: MediaBudget = MEDIA_BUDGET,
  roots: readonly string[] = NO_MEDIA_ROOTS,
): Promise<MediaBlob[]> {
  const maxItems = budget.maxItems ?? MAX_MEDIA_ITEMS
  const maxItemBytes = budget.maxItemBytes ?? MAX_MEDIA_BYTES
  const maxTotalBytes = budget.maxTotalBytes ?? MAX_MEDIA_TOTAL_BYTES
  const concurrency = Math.max(1, budget.concurrency ?? MEDIA_LOAD_CONCURRENCY)
  assertMediaItemCount(refs.length, maxItems)
  const declared = refs.map(dataUrlDecodedSize)
  for (const bytes of declared) {
    if (bytes > maxItemBytes) {
      throw new MediaTooLargeError(`data URL is too large to analyze (limit ${maxItemBytes} bytes)`)
    }
  }
  const declaredTotal = declared.reduce((n, bytes) => n + bytes, 0)
  if (declaredTotal > maxTotalBytes) {
    throw new MediaBudgetExceededError(
      `Media in one request is too large to analyze (${declaredTotal} bytes, limit ${maxTotalBytes}); analyze it in smaller batches`,
    )
  }
  const blobs: MediaBlob[] = new Array(refs.length)
  let next = 0
  let landed = 0
  const worker = async (): Promise<void> => {
    while (next < refs.length) {
      const index = next++
      const blob = await loadMediaReference(refs[index]!, roots)
      landed += blob.bytes.byteLength
      if (landed > maxTotalBytes) {
        throw new MediaBudgetExceededError(
          `Media in one request is too large to analyze (over ${maxTotalBytes} bytes); analyze it in smaller batches`,
        )
      }
      blobs[index] = blob
    }
  }
  await Promise.all(Array.from({ length: Math.min(concurrency, refs.length) }, () => worker()))
  return blobs
}

/** Genspark background-removal model — chained after generation for transparentBackground */
export const GSK_RMBG_MODEL = 'fal-bria-rmbg'

export type GenerateImageToolOp = GskGenerateImageOptions & {
  /** The result must have real PNG alpha (icons/logos/cutouts). Generation models cannot
   * produce transparency from the prompt alone — they paint a fake gray checkerboard into
   * the pixels — so the tool strips the background in a second pass instead. */
  transparentBackground?: boolean
}

export async function generateImageTool(
  settingsPath: string,
  op: GenerateImageToolOp,
  options: MediaToolOptions = {},
): Promise<{ url?: string; error?: string }> {
  const prompt = String(op.prompt ?? '').trim()
  if (!prompt) return { error: 'prompt must not be empty' }
  const mediaRoots = options.mediaRoots ?? NO_MEDIA_ROOTS
  const settings = readAiSettingsFile(settingsPath)
  const byok = activeMediaConfig(settings, 'image')
  try {
    if (!byok) {
      const gate = gskGate(settings, options.notLoggedInError ?? GSK_NOT_LOGGED_IN_ERROR)
      if (gate) return gate
      // gskGenerateImage hands the references straight to the CLI argv, so this route
      // enforces the same item ceiling as the BYOK one instead of passing them on
      assertMediaItemCount(op.referenceImageUrls?.length ?? 0, MEDIA_BUDGET.maxItems)
      const gen = await gskGenerateImage({ ...op, prompt })
      if (!op.transparentBackground || op.model === GSK_RMBG_MODEL) return { url: gen.url }
      try {
        const cut = await gskGenerateImage({
          prompt: 'remove the background completely, keep only the subject',
          model: GSK_RMBG_MODEL,
          referenceImageUrls: [gen.url],
        })
        return { url: cut.url }
      } catch {
        return { url: gen.url } // strip failed: the opaque image is still usable
      }
    }
    // `model` names Genspark-only special models (fal-*); BYOK uses the configured image model
    const references = await loadMediaReferences(
      op.referenceImageUrls ?? [],
      MEDIA_BUDGET,
      mediaRoots,
    )
    const image = await generateImageWithProvider(byok.provider, byok.config, {
      prompt,
      aspectRatio: op.aspectRatio,
      references,
      transparent: op.transparentBackground === true,
    })
    return { url: storeGeneratedImage(image.bytes, image.mime) }
  } catch (err) {
    return { error: errorText(err) }
  }
}

export async function analyzeMediaTool(
  settingsPath: string,
  op: { mediaUrls: string[]; requirements: string },
  options: MediaToolOptions = {},
): Promise<{ text?: string; error?: string }> {
  const mediaUrls = (op.mediaUrls ?? []).map(String).filter(Boolean)
  const requirements = String(op.requirements ?? '').trim()
  if (!mediaUrls.length) return { error: 'mediaUrls must not be empty' }
  if (!requirements) return { error: 'requirements must not be empty' }
  const mediaRoots = options.mediaRoots ?? NO_MEDIA_ROOTS
  const settings = readAiSettingsFile(settingsPath)
  const imageByok = activeMediaConfig(settings, 'analysis')
  const videoByok = activeMediaConfig(settings, 'video')
  try {
    const viaGsk = async () => {
      const gate = gskGate(settings, options.notLoggedInError ?? GSK_NOT_LOGGED_IN_ERROR)
      if (gate) return gate
      // gskAnalyzeMedia spreads the URLs straight into the CLI argv, so this route
      // enforces the same item ceiling as the BYOK one instead of passing them on
      assertMediaItemCount(mediaUrls.length, MEDIA_BUDGET.maxItems)
      return { text: await gskAnalyzeMedia({ mediaUrls, requirements }) }
    }
    if (!imageByok && !videoByok) return await viaGsk()
    // route on the loaded bytes' real MIME, not the URL spelling: images go to the
    // image-analysis provider, anything with video/audio to the video one
    let media: MediaBlob[]
    try {
      media = await loadMediaReferences(mediaUrls, MEDIA_BUDGET, mediaRoots)
    } catch (err) {
      // only the size cap hands the request back to Genspark (the CLI streams large
      // files itself); scheme / path / SSRF rejections stay rejections
      if (err instanceof MediaTooLargeError && (!imageByok || !videoByok)) return await viaGsk()
      return { error: errorText(err) }
    }
    const hasVideo = media.some((m) => !m.mime.startsWith('image/'))
    const byok = hasVideo ? videoByok : imageByok
    if (!byok) return await viaGsk()
    return {
      text: await analyzeMediaWithProvider(byok.provider, byok.config, { media, requirements }),
    }
  } catch (err) {
    return { error: errorText(err) }
  }
}
