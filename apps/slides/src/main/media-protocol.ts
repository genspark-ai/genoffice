/**
 * Serves Slides media over `genoffice-slides-media://<deck>/<part>` (#763 follow-up).
 *
 * Mirrors what the docs app does for `genoffice-docx-media://`: the render tree
 * carries a short URL per picture instead of a base64 copy, and the bytes stay
 * where they already live — the deck's archive in this process.
 *
 * The scheme is registered as privileged in packages/electron-utils (before app
 * ready, next to the renderer and docx schemes); this module only handles requests.
 */
import { protocol } from 'electron'
import { SLIDES_MEDIA_SCHEME, SLIDES_MEDIA_SCHEME_PRIVILEGE } from '@genoffice/electron-utils'
import { parseSlidesMediaUrl, slidesMediaBase, type SlidesMediaSource } from '../shared/media-url'

const sources = new Map<string, SlidesMediaSource>()
let nextDeck = 0

/** Register a deck's media and hand back the base URL its render tree should use. */
export function registerSlidesMediaSource(source: SlidesMediaSource): string {
  // Self-registering on first use: a render tree is only built after app ready, and
  // this way both entry points (the standalone slides app and the shell, which
  // imports this module) are covered without a second wiring site. The privileged
  // scheme itself is declared in packages/electron-utils before app ready.
  registerSlidesMediaProtocol()
  const key = `deck${(nextDeck += 1).toString(36)}`
  sources.set(key, source)
  return slidesMediaBase(key)
}

export function registerSlidesMediaProtocol(): void {
  if (SLIDES_MEDIA_SCHEME_PRIVILEGE.scheme !== SLIDES_MEDIA_SCHEME) {
    throw new Error('slides media scheme privilege registered under another name')
  }
  if (protocol.isProtocolHandled(SLIDES_MEDIA_SCHEME)) return
  protocol.handle(SLIDES_MEDIA_SCHEME, (request) => {
    const parsed = parseSlidesMediaUrl(request.url)
    const part = parsed ? sources.get(parsed.key)?.read(parsed.mediaRef) : undefined
    if (!part) return new Response(null, { status: 404 })
    return new Response(part.bytes, {
      headers: {
        'Content-Type': part.mime,
        'Content-Length': String(part.bytes.byteLength),
        // media bytes are immutable for a session, and the URL is per-deck
        'Cache-Control': 'no-store',
      },
    })
  })
}
