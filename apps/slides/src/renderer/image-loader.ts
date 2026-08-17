/**
 * Incremental image loading: each decoded image is surfaced in small batches
 * instead of waiting for the whole deck (385 pictures used to render nothing
 * until the last one settled). Loaded/in-flight urls are tracked across calls
 * so re-collecting urls after an edit never reloads or discards progress.
 */
import { metafileToDataUrl } from '@genoffice/docx-engine/metafile'

export type ApplyImages = (entries: ReadonlyArray<readonly [string, HTMLImageElement]>) => void

/** EMF/WMF data URLs: browsers cannot decode metafiles — rasterize to PNG first (keyed by the original url). */
const METAFILE_RE = /^data:(image\/x-(?:emf|wmf)|image\/(?:emf|wmf));base64,/

async function rasterizeMetafile(url: string): Promise<string | null> {
  const m = METAFILE_RE.exec(url)
  if (!m) return null
  const b64 = url.slice(url.indexOf(',') + 1)
  const bin = atob(b64)
  const bytes = new Uint8Array(bin.length)
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i)
  const mime = m[1]!.includes('emf') ? 'image/x-emf' : 'image/x-wmf'
  return metafileToDataUrl(bytes, mime)
}

export function createImageLoader(apply: ApplyImages, batchSize = 16, delayMs = 100) {
  const loaded = new Map<string, HTMLImageElement>()
  const loading = new Set<string>()
  const buf = new Map<string, HTMLImageElement>()
  let timer: ReturnType<typeof setTimeout> | null = null
  let disposed = false

  const flush = () => {
    if (timer) {
      clearTimeout(timer)
      timer = null
    }
    if (disposed || buf.size === 0) return
    const entries = [...buf]
    buf.clear()
    apply(entries)
  }

  return {
    load(urls: Iterable<string>) {
      for (const u of urls) {
        if (loaded.has(u) || loading.has(u)) continue
        loading.add(u)
        const img = new Image()
        const done = (ok: boolean) => {
          loading.delete(u)
          if (ok) {
            loaded.set(u, img)
            if (!disposed) buf.set(u, img)
          }
          if (buf.size >= batchSize || loading.size === 0) flush()
          else if (!timer && buf.size > 0) timer = setTimeout(flush, delayMs)
        }
        img.onload = () => done(true)
        img.onerror = () => done(false)
        if (METAFILE_RE.test(u)) {
          void rasterizeMetafile(u)
            .then((png) => {
              if (png) img.src = png
              else done(false)
            })
            .catch(() => done(false))
        } else {
          img.src = u
        }
      }
    },
    // Only guards setState after unmount; in-flight loads keep filling `loaded`
    dispose() {
      disposed = true
      if (timer) clearTimeout(timer)
    },
  }
}
