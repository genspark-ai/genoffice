/**
 * generate_image / analyze_media for the five editors' main processes: one
 * place that reads ai-settings.json live, routes to the BYOK media provider
 * when one is configured, and otherwise to the Genspark CLI behind the usual
 * login + cloud-tools gate. BYOK providers answer with bytes; those land in
 * the local generated-image store and come back as a file:// URL that the
 * insert pipelines' fetchRemoteImage accepts.
 */

import { existsSync, readFileSync, statSync } from 'node:fs'
import { basename, extname } from 'node:path'
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
import {
  fetchRemoteImage,
  fetchWithSsrfGuard,
  readGeneratedImage,
  storeGeneratedImage,
} from '@genoffice/electron-utils'
import { gskAnalyzeMedia, gskGenerateImage, hasGskAuth, type GskGenerateImageOptions } from './gsk'

export const GSK_NOT_LOGGED_IN_ERROR =
  'Genspark account is not logged in on this machine; ask the user to log in first'
export const GSK_TOOLS_OFF_ERROR =
  'Genspark cloud tools are turned off in Settings (AI Model); enable them or configure an image provider under Settings (AI Media) to use this tool'

/** 200 MB: enough for a long clip through the Gemini Files API, small enough to hold in memory */
const MAX_MEDIA_BYTES = 200 * 1024 * 1024

/** the only load failure that may hand the request back to Genspark; validation failures never do */
export class MediaTooLargeError extends Error {}

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
 * (attachments). Only media extensions are read locally — the model must not be
 * able to ship arbitrary files to a vendor.
 */
export async function loadMediaReference(ref: string): Promise<MediaBlob> {
  if (/^https?:\/\//i.test(ref)) {
    const resp = await (ref.match(/\.(png|jpe?g|gif|webp)(\?|$)/i)
      ? fetchRemoteImage(ref)
      : fetchWithSsrfGuard(ref, { headers: { 'User-Agent': 'Mozilla/5.0' } }))
    if (!resp || !resp.ok) throw new Error(`Could not download ${ref}`)
    const declared = Number(resp.headers.get('content-length') ?? 0)
    if (declared > MAX_MEDIA_BYTES) throw new MediaTooLargeError(`${ref} is too large to analyze`)
    const bytes = new Uint8Array(await resp.arrayBuffer())
    if (bytes.byteLength > MAX_MEDIA_BYTES) {
      throw new MediaTooLargeError(`${ref} is too large to analyze`)
    }
    const ct = resp.headers.get('content-type')?.split(';')[0]?.trim()
    const name = basename(new URL(ref).pathname) || undefined
    const mime =
      ct && ct !== 'application/octet-stream' ? ct : MIME_BY_EXT[extname(name ?? '').toLowerCase()]
    if (!mime) throw new Error(`Could not tell the media type of ${ref}`)
    return { bytes, mime, ...(name ? { name } : {}) }
  }
  if (ref.startsWith('file:')) {
    const local = readGeneratedImage(ref)
    if (!local) throw new Error(`Not an accessible image: ${ref}`)
    return { bytes: new Uint8Array(local.bytes), mime: local.mime }
  }
  const mime = MIME_BY_EXT[extname(ref).toLowerCase()]
  if (!mime) throw new Error(`Unsupported media file: ${ref} (images, video and audio only)`)
  if (!existsSync(ref)) throw new Error(`File not found: ${ref}`)
  if (statSync(ref).size > MAX_MEDIA_BYTES) {
    throw new MediaTooLargeError(`${ref} is too large to analyze`)
  }
  return { bytes: new Uint8Array(readFileSync(ref)), mime, name: basename(ref) }
}

export interface MediaToolOptions {
  /** localized replacement for the default signed-out message */
  notLoggedInError?: string
}

export async function generateImageTool(
  settingsPath: string,
  op: GskGenerateImageOptions,
  options: MediaToolOptions = {},
): Promise<{ url?: string; error?: string }> {
  const prompt = String(op.prompt ?? '').trim()
  if (!prompt) return { error: 'prompt must not be empty' }
  const settings = readAiSettingsFile(settingsPath)
  const byok = activeMediaConfig(settings, 'image')
  try {
    if (!byok) {
      const gate = gskGate(settings, options.notLoggedInError ?? GSK_NOT_LOGGED_IN_ERROR)
      if (gate) return gate
      return { url: (await gskGenerateImage({ ...op, prompt })).url }
    }
    // `model` names Genspark-only special models (fal-*); BYOK uses the configured image model
    const references = await Promise.all((op.referenceImageUrls ?? []).map(loadMediaReference))
    const image = await generateImageWithProvider(byok.provider, byok.config, {
      prompt,
      aspectRatio: op.aspectRatio,
      references,
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
  const settings = readAiSettingsFile(settingsPath)
  const imageByok = activeMediaConfig(settings, 'analysis')
  const videoByok = activeMediaConfig(settings, 'video')
  try {
    const viaGsk = async () => {
      const gate = gskGate(settings, options.notLoggedInError ?? GSK_NOT_LOGGED_IN_ERROR)
      if (gate) return gate
      return { text: await gskAnalyzeMedia({ mediaUrls, requirements }) }
    }
    if (!imageByok && !videoByok) return await viaGsk()
    // route on the loaded bytes' real MIME, not the URL spelling: images go to the
    // image-analysis provider, anything with video/audio to the video one
    let media: MediaBlob[]
    try {
      media = await Promise.all(mediaUrls.map(loadMediaReference))
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
