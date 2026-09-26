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

/**
 * EMF headers size the raster from their device bounds, up to the converter's 8192px cap:
 * a 409-slide deck with 67 metafiles held 2.7 GB of 268 MB canvases. A slide picture never
 * shows above ~2x slide width, so cap the longest side.
 */
export const MAX_METAFILE_RASTER_PX = 2048
/** Base64 budget for metafile raster input (~40MB of bytes). */
export const MAX_METAFILE_BASE64_CHARS = 56 * 1024 * 1024

/**
 * Metafile text draws through canvas fonts, so the Office-private FontFaces (DFonts/cloud/
 * embedded, registered by doc-fonts.ts after the deck settles) must be in place first — an
 * EMF rasterized before that keeps its fallback face forever (Excel OLE previews in
 * Meiryo UI came out in the browser's default sans). `false` = a sync is in flight.
 */
function waitForDocFonts(timeoutMs = 4000): Promise<void> {
  if (typeof window === 'undefined' || window.__genofficeDocFontsSynced !== false)
    return Promise.resolve()
  return new Promise((resolve) => {
    const started = Date.now()
    const tick = () => {
      if (window.__genofficeDocFontsSynced !== false || Date.now() - started >= timeoutMs) resolve()
      else setTimeout(tick, 50)
    }
    setTimeout(tick, 50)
  })
}

async function rasterizeMetafile(url: string): Promise<string | null> {
  const m = METAFILE_RE.exec(url)
  if (!m) return null
  await waitForDocFonts()
  const b64 = url.slice(url.indexOf(',') + 1)
  // atob + Uint8Array double-allocate the payload on the main thread: refuse
  // huge metafiles before decoding (mirrors the 40MB PNG cap in element-clipboard).
  if (b64.length > MAX_METAFILE_BASE64_CHARS) return null
  const bin = atob(b64)
  const bytes = new Uint8Array(bin.length)
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i)
  const mime = m[1]!.includes('emf') ? 'image/x-emf' : 'image/x-wmf'
  return metafileToDataUrl(bytes, mime, { maxSidePx: MAX_METAFILE_RASTER_PX })
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
    /** urls still decoding — 0 means every image the deck asked for has settled */
    pending(): number {
      return loading.size
    },
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
