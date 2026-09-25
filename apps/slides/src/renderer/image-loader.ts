/**
 * Incremental image loading: each decoded image is surfaced in small batches
 * instead of waiting for the whole deck (385 pictures used to render nothing
 * until the last one settled). Loaded/in-flight urls are tracked across calls
 * so re-collecting urls after an edit never reloads or discards progress.
 *
 * #763 follow-up — media is *decode-on-demand* now:
 *
 * - the caller passes only the urls the app needs at that moment (the slides on
 *   screen plus a neighbour for prefetch), so opening a 409-slide deck no longer
 *   decodes all 138 of its images up front. Measured on the reporter's deck that
 *   was ~1.9 GB of the 3.2 GB process memory and ~38% of the time to first paint,
 *   with most of the GPU process's bytes being image textures.
 * - a decoded image is kept at most `maxSide` px on its longest side. The stage is
 *   1280 px wide and a rail thumbnail 126 px, so a 3000 px photo has no reason to
 *   be held at 3000 px. Images already within the cap keep their `<img>`; larger
 *   ones are drawn once into a canvas at the capped size, which is what is kept.
 * - the decoded set is bounded by `budgetBytes`: least recently needed images are
 *   dropped and reported through `onEvict` so the UI forgets them, which is what
 *   keeps memory flat while the user pages through a long deck.
 */
import { metafileToDataUrl } from '@genoffice/docx-engine/metafile'

/** A decoded image as the renderer consumes it (kept as `<img>` so nothing downstream changes). */
export type SlideImage = HTMLImageElement

export type ApplyImages = (entries: ReadonlyArray<readonly [string, SlideImage]>) => void

/** EMF/WMF data URLs: browsers cannot decode metafiles — rasterize to PNG first (keyed by the original url). */
const METAFILE_RE = /^data:(image\/x-(?:emf|wmf)|image\/(?:emf|wmf));base64,/

/** Base64 budget for metafile raster input (~40MB of bytes). */
export const MAX_METAFILE_BASE64_CHARS = 56 * 1024 * 1024

/** 4 bytes per pixel: what a decoded image or canvas costs. */
const BYTES_PER_PIXEL = 4

/** Longest side retained when the caller does not say (2x the 1280 px stage). */
export const DEFAULT_MAX_SIDE_PX = 2560

/** Decoded pixels kept at once (~3 screens of media at the cap). */
export const DEFAULT_BUDGET_BYTES = 384 * 1024 * 1024

/**
 * Metafile text draws through canvas fonts, so the Office-private FontFaces (DFonts/cloud/
 * embedded, registered by doc-fonts.ts after the deck settles) must be in place first — an
 * EMF rasterized before that keeps its fallback face forever (Excel OLE previews in Meiryo UI
 * came out in the browser's default sans). `false` = a sync is in flight.
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
  return metafileToDataUrl(bytes, mime)
}

function loadElement(url: string): Promise<HTMLImageElement | null> {
  return new Promise((resolve) => {
    const img = new Image()
    img.onload = () => resolve(img)
    img.onerror = () => resolve(null)
    img.src = url
  })
}

/** Sources worth capping: lossy, and typically the ones that arrive at camera sizes. */
const CAP_CANDIDATE_RE = /^data:image\/(jpeg|jpg|webp|avif)[;,]/

function elementFromBlob(blob: Blob): Promise<HTMLImageElement | null> {
  const url = URL.createObjectURL(blob)
  return loadElement(url).finally(() => URL.revokeObjectURL(url))
}

/**
 * Decode `url`, keeping it at most `maxSide` px on the longest side.
 *
 * The cap is only applied to lossy sources (JPEG/WebP/AVIF) that exceed it: a
 * 3000 px photo is drawn once into a `maxSide` canvas and handed back as an
 * object-URL image, so what the renderer holds — and what the GPU uploads — is the
 * capped size. PNG/EMF/SVG keep the existing path untouched (they are usually
 * already small, and re-encoding them would cost more than it saves).
 *
 * Falls back to a plain `<img>` whenever the pipeline is unavailable (no
 * `createImageBitmap`, an undecodable blob, a test environment), i.e. exactly the
 * behaviour this module had before the cap existed.
 */
export async function decodeCapped(url: string, maxSide: number): Promise<SlideImage | null> {
  let sourceUrl = url
  if (METAFILE_RE.test(url)) {
    const png = await rasterizeMetafile(url)
    if (!png) return null
    sourceUrl = png
  }
  const capWorthTrying =
    CAP_CANDIDATE_RE.test(sourceUrl) &&
    typeof createImageBitmap === 'function' &&
    typeof fetch === 'function'
  if (!capWorthTrying) return await loadElement(sourceUrl)

  let bitmap: ImageBitmap | null = null
  try {
    const blob = await (await fetch(sourceUrl)).blob()
    bitmap = await createImageBitmap(blob)
    const longest = Math.max(bitmap.width, bitmap.height)
    if (longest <= maxSide) return await loadElement(sourceUrl)
    const scale = maxSide / longest
    const canvas = document.createElement('canvas')
    canvas.width = Math.max(1, Math.round(bitmap.width * scale))
    canvas.height = Math.max(1, Math.round(bitmap.height * scale))
    const ctx = canvas.getContext('2d')
    if (!ctx) return await loadElement(sourceUrl)
    ctx.drawImage(bitmap, 0, 0, canvas.width, canvas.height)
    const capped = await new Promise<Blob | null>((resolve) =>
      canvas.toBlob((b) => resolve(b), 'image/webp', 0.92),
    )
    if (!capped) return await loadElement(sourceUrl)
    return await elementFromBlob(capped)
  } catch {
    return await loadElement(sourceUrl)
  } finally {
    bitmap?.close?.()
  }
}

/** Decoded cost of an image: 4 bytes per pixel. */
export function imageBytes(image: SlideImage): number {
  const w = image.naturalWidth || image.width
  const h = image.naturalHeight || image.height
  return Math.max(1, w) * Math.max(1, h) * BYTES_PER_PIXEL
}

export interface ImageLoaderOptions {
  /** Longest side a retained image keeps, in px. */
  maxSide?: number
  /** Decoded pixels kept before the least recently needed images are dropped. */
  budgetBytes?: number
  /** Called for every dropped image, so the UI can forget it. */
  onEvict?: (url: string) => void
  /** Images surfaced to `apply` per batch. */
  batchSize?: number
  /** Delay before a partial batch is surfaced, in ms. */
  delayMs?: number
}

export function createImageLoader(apply: ApplyImages, options: ImageLoaderOptions = {}) {
  const maxSide = options.maxSide ?? DEFAULT_MAX_SIDE_PX
  const budgetBytes = options.budgetBytes ?? DEFAULT_BUDGET_BYTES
  const batchSize = options.batchSize ?? 16
  const delayMs = options.delayMs ?? 100

  /** Urls the app wants right now; anything else is an eviction candidate. */
  let needed = new Set<string>()
  /** Recency order — the first key is the coldest. */
  const loaded = new Map<string, SlideImage>()
  const loading = new Set<string>()
  const buf = new Map<string, SlideImage>()
  let bytes = 0
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

  /** Drop the coldest images that are no longer needed, until the budget is met. */
  const evict = () => {
    if (bytes <= budgetBytes) return
    for (const url of [...loaded.keys()]) {
      if (bytes <= budgetBytes) break
      if (needed.has(url) || buf.has(url)) continue // on screen, or about to be
      const image = loaded.get(url)
      loaded.delete(url)
      if (image) bytes -= imageBytes(image)
      options.onEvict?.(url)
    }
  }

  const touch = (url: string) => {
    const image = loaded.get(url)
    if (!image) return
    loaded.delete(url)
    loaded.set(url, image)
  }

  const settle = (url: string, image: SlideImage | null) => {
    loading.delete(url)
    if (image && !disposed) {
      loaded.set(url, image)
      bytes += imageBytes(image)
      buf.set(url, image)
    }
    if (buf.size >= batchSize || loading.size === 0) flush()
    else if (!timer && buf.size > 0) timer = setTimeout(flush, delayMs)
    if (loading.size === 0) evict()
  }

  /**
   * The path every source took before the cap existed: an `<img>` created
   * synchronously and decoded by the browser at its own size. Kept as-is for
   * non-lossy sources (PNG/EMF/SVG), which are usually already small.
   */
  const startUncapped = (url: string) => {
    const img = new Image()
    const done = (ok: boolean) => settle(url, ok ? img : null)
    img.onload = () => done(true)
    img.onerror = () => done(false)
    if (METAFILE_RE.test(url)) {
      void rasterizeMetafile(url)
        .then((png) => {
          if (png) img.src = png
          else done(false)
        })
        .catch(() => done(false))
    } else {
      img.src = url
    }
  }

  /** Lossy sources that exceed `maxSide` are decoded through the capping pipeline. */
  const startCapped = (url: string) => {
    void decodeCapped(url, maxSide)
      .then((image) => settle(url, image))
      .catch(() => settle(url, null))
  }

  const capAvailable = typeof createImageBitmap === 'function' && typeof fetch === 'function'

  return {
    /** urls still decoding — 0 means every image the deck asked for has settled */
    pending(): number {
      return loading.size
    },
    /** Decoded pixels retained right now (diagnostics/tests). */
    bytes(): number {
      return bytes
    },
    /**
     * The urls needed right now. Anything else becomes evictable once the budget
     * is exceeded, so paging through a long deck does not retain its whole media set.
     */
    load(urls: Iterable<string>) {
      needed = new Set(urls)
      for (const url of needed) touch(url)
      for (const url of needed) {
        if (loaded.has(url) || loading.has(url)) continue
        loading.add(url)
        if (capAvailable && CAP_CANDIDATE_RE.test(url)) startCapped(url)
        else startUncapped(url)
      }
      if (loading.size === 0) evict()
    },
    /** Forget decoded media (a different deck was opened). */
    clear() {
      loaded.clear()
      buf.clear()
      bytes = 0
    },
    // Only guards setState after unmount; in-flight loads keep filling `loaded`
    dispose() {
      disposed = true
      if (timer) clearTimeout(timer)
    },
  }
}
