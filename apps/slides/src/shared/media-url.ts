/**
 * Lazy media URLs for Slides (#763 follow-up).
 *
 * The render tree used to carry every picture as a `data:` URL, so a photo-heavy
 * deck kept its media as base64 strings inside the model (and in each copy of it)
 * *and* as decoded bitmaps. Measured on the 409-slide reporter deck, ~34 MB of
 * base64 plus its copies, and most of the remaining time-to-first-paint after the
 * decode-on-demand change.
 *
 * Instead, main serves the bytes over a custom scheme and the tree carries a short
 * URL. This module is deliberately electron-free so the URL shape, the parsing and
 * the size threshold can be unit-tested; the protocol handler that uses it lives in
 * media-protocol.ts, mirroring what the docs app does for `genoffice-docx-media://`.
 */
import { SLIDES_MEDIA_SCHEME } from '@genoffice/electron-utils'

/**
 * Parts at least this large are served by URL; smaller ones stay inline as data
 * URLs. Below it the base64 copy is not worth a round trip (and keeping tiny icons
 * inline means a deck of icons never touches the protocol at all).
 */
export const LAZY_MEDIA_MIN_PART_BYTES = 64 * 1024

/** Mime families the protocol serves: plain rasters, no per-slide context needed. */
const LAZY_RASTER_MIME = /^image\/(png|jpe?g|webp|gif|bmp|avif)$/

export interface SlidesMediaPart {
  bytes: Uint8Array
  mime: string
}

/** What the protocol handler asks for bytes: the deck's archive, by media ref. */
export interface SlidesMediaSource {
  read(mediaRef: string): SlidesMediaPart | undefined
}

export function slidesMediaBase(key: string): string {
  return `${SLIDES_MEDIA_SCHEME}://${key}/`
}

export function slidesMediaUrl(base: string, mediaRef: string): string {
  return base + encodeURIComponent(mediaRef)
}

/** Parse a served URL back into its deck key and media ref (export/diagnostics). */
export function parseSlidesMediaUrl(url: string): { key: string; mediaRef: string } | null {
  const prefix = `${SLIDES_MEDIA_SCHEME}://`
  if (!url.startsWith(prefix)) return null
  const rest = url.slice(prefix.length)
  const slash = rest.indexOf('/')
  if (slash <= 0) return null
  const key = rest.slice(0, slash)
  const mediaRef = decodeURIComponent(rest.slice(slash + 1))
  if (!key || !mediaRef) return null
  return { key, mediaRef }
}

/**
 * Whether this part should be served by URL rather than inlined as a data URL.
 *
 * Only rasters qualify. SVG and EMF need a per-slide theme retint and stay inline,
 * which is also what lets one media source serve a whole deck (rasters need no
 * slide context).
 */
export function shouldServeLazily(mime: string, byteLength: number): boolean {
  return LAZY_RASTER_MIME.test(mime) && byteLength >= LAZY_MEDIA_MIN_PART_BYTES
}
