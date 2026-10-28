/**
 * The per-app custom media schemes exist only inside Electron; the web front
 * end rewrites them to same-origin paths. Bundle code is rewritten at serve
 * time (static.ts), but scheme strings also travel through the IPC bridge
 * (an open reply carries the document bytes as a media URL, the renderer
 * passes image sources back on save) — those are rewritten here, on both
 * directions of every frame, keyed by the tab's app.
 */

export const SCHEME_REWRITES: Record<string, [string, string][]> = {
  docs: [['genoffice-docx-media', '/__media']],
  markdown: [['md-asset', '/__mdasset']],
  html: [
    ['html-preview', '/__preview'],
    ['html-asset', '/__htmlasset'],
  ],
}

function walk(value: unknown, visit: (s: string) => string): unknown {
  if (typeof value === 'string') return visit(value)
  if (Array.isArray(value)) return value.map((v) => walk(v, visit))
  if (value !== null && typeof value === 'object') {
    // only plain objects rebuild; typed arrays (decoded binaries) and other
    // classed instances pass through untouched — index-key reconstruction
    // would turn a Uint8Array into {0:1, 1:2, …}
    const proto = Object.getPrototypeOf(value)
    if (proto !== Object.prototype && proto !== null) return value
    const out: Record<string, unknown> = {}
    for (const [key, entry] of Object.entries(value as Record<string, unknown>)) {
      out[key] = walk(entry, visit)
    }
    return out
  }
  return value
}

/** server → browser: scheme literals become the same-origin paths the served
 *  bundle code expects */
export function rewriteOut(app: string, value: unknown): unknown {
  const rules = SCHEME_REWRITES[app]
  if (!rules) return value
  return walk(value, (s) => {
    for (const [scheme, prefix] of rules) s = s.split(scheme).join(prefix)
    return s
  })
}

/** browser → server: same-origin paths become the scheme literals the
 *  main-process handlers expect */
export function rewriteIn(app: string, value: unknown): unknown {
  const rules = SCHEME_REWRITES[app]
  if (!rules) return value
  return walk(value, (s) => {
    for (const [scheme, prefix] of rules) s = s.split(prefix).join(scheme)
    return s
  })
}
