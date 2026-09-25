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

/**
 * Longest side for images the app only draws in the rail. The rail thumbnail is
 * 126 px wide, so decoding at the stage cap wasted most of the pixels: measured on
 * the reporter's deck, one 16-row window needs ~100 pictures, which at 2560 px came
 * to 2.7 GB — seven times the byte budget, so nothing was ever evictable.
 */
export const DEFAULT_THUMB_MAX_SIDE_PX = 256

/** Decodes in flight at once; the pipeline runs on the main thread. */
export const DEFAULT_MAX_CONCURRENT = 4

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

/**
 * Sources worth capping: any raster. This started out as lossy-only (JPEG/WebP/AVIF)
 * on the assumption that PNG is small — wrong for the measured deck, whose 138 parts
 * are PNG screenshots: they took the uncapped path and were held at 31-44 MB each, so
 * a 16-row window (~100 pictures) came to 2.7 GB and nothing could be evicted.
 * SVG stays out: it is retinted per slide, so it is never capped.
 */
const CAP_CANDIDATE_RE = /^data:image\/(png|jpeg|jpg|webp|avif|gif|bmp)[;,]/

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
    // PNG keeps its format: screenshots and anything with alpha should not be re-encoded
    // lossily; everything else goes to WebP, which keeps alpha and compresses harder.
    const isPng = /^data:image\/png/i.test(sourceUrl)
    const capped = await new Promise<Blob | null>((resolve) =>
      canvas.toBlob((b) => resolve(b), isPng ? 'image/png' : 'image/webp', 0.92),
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
  /** Longest side a retained image keeps, in px: what the stage can draw. */
  maxSide?: number
  /** Longest side for images the app only shows in the rail (see the default). */
  thumbMaxSide?: number
  /** Decoded pixels kept before the least recently needed images are dropped. */
  budgetBytes?: number
  /** Called for every dropped image, so the UI can forget it. */
  onEvict?: (url: string) => void
  /** Images surfaced to `apply` per batch. */
  batchSize?: number
  /** Delay before a partial batch is surfaced, in ms. */
  delayMs?: number
  /**
   * Decodes in flight at once. A queue of dozens of large images starves React's
   * commits — measured: a row scrolled into view never mounted for 20 s while
   * ~100 decodes were queued on the main thread.
   */
  maxConcurrent?: number
}

export function createImageLoader(apply: ApplyImages, options: ImageLoaderOptions = {}) {
  const maxSide = options.maxSide ?? DEFAULT_MAX_SIDE_PX
  const thumbMaxSide = Math.min(options.thumbMaxSide ?? DEFAULT_THUMB_MAX_SIDE_PX, maxSide)
  const budgetBytes = options.budgetBytes ?? DEFAULT_BUDGET_BYTES
  const batchSize = options.batchSize ?? 16
  const delayMs = options.delayMs ?? 100
  const maxConcurrent = Math.max(1, options.maxConcurrent ?? DEFAULT_MAX_CONCURRENT)

  /** Urls the app wants right now; anything else is an eviction candidate. */
  let needed = new Set<string>()
  /** Urls the app draws on the stage rather than only in the rail. */
  let large = new Set<string>()
  /** Recency order — the first key is the coldest. */
  const loaded = new Map<string, SlideImage>()
  /** Longest side each retained image was decoded at, so a rail-sized one can be upgraded. */
  const decodedSide = new Map<string, number>()
  const loading = new Set<string>()
  const queue: string[] = []
  const queued = new Set<string>()
  const buf = new Map<string, SlideImage>()

  /** Stage-sized for what is on the canvas, rail-sized for thumbnails. */
  const sideFor = (url: string) => (large.has(url) ? maxSide : thumbMaxSide)
  let bytes = 0
  let decoded = 0
  let evicted = 0
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
      decodedSide.delete(url)
      if (image) bytes -= imageBytes(image)
      evicted += 1
      options.onEvict?.(url)
    }
  }

  const touch = (url: string) => {
    const image = loaded.get(url)
    if (!image) return
    loaded.delete(url)
    loaded.set(url, image)
  }

  const settle = (url: string, image: SlideImage | null, side: number) => {
    loading.delete(url)
    if (image && !disposed) {
      loaded.set(url, image)
      decodedSide.set(url, side)
      bytes += imageBytes(image)
      decoded += 1
      buf.set(url, image)
    }
    if (buf.size >= batchSize || loading.size === 0) flush()
    else if (!timer && buf.size > 0) timer = setTimeout(flush, delayMs)
    if (loading.size === 0) evict()
    pump()
  }

  /** Start queued decodes up to the concurrency limit. */
  const pump = () => {
    while (loading.size < maxConcurrent && queue.length > 0) {
      const url = queue.shift() as string
      queued.delete(url)
      if (loaded.has(url) || loading.has(url)) continue
      loading.add(url)
      if (capAvailable && CAP_CANDIDATE_RE.test(url)) startCapped(url, sideFor(url))
      else startUncapped(url)
    }
  }

  /**
   * The path every source took before the cap existed: an `<img>` created
   * synchronously and decoded by the browser at its own size. Kept as-is for
   * non-lossy sources (PNG/EMF/SVG), which are usually already small.
   */
  const startUncapped = (url: string) => {
    const img = new Image()
    // uncapped sources are kept as decoded; MAX_SAFE_INTEGER marks them as never worth upgrading
    const done = (ok: boolean) => settle(url, ok ? img : null, Number.MAX_SAFE_INTEGER)
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
  const startCapped = (url: string, side: number) => {
    void decodeCapped(url, side)
      .then((image) => settle(url, image, side))
      .catch(() => settle(url, null, side))
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
     * Counters for perf work: what this loader holds and has been through. Paging a
     * long deck should keep `retainedBytes` near the budget — if it does and process
     * memory is still high, the bytes are in the browser's own image cache, not here.
     */
    stats() {
      let largeRetained = 0
      let thumbRetained = 0
      for (const side of decodedSide.values()) {
        if (side >= maxSide) largeRetained += 1
        else thumbRetained += 1
      }
      return {
        retainedBytes: bytes,
        retainedImages: loaded.size,
        neededImages: needed.size,
        inFlight: loading.size,
        queued: queue.length,
        largeRetained,
        thumbRetained,
        decoded,
        evicted,
        stageSide: maxSide,
        thumbSide: thumbMaxSide,
      }
    },
    /**
     * The urls needed right now. Anything else becomes evictable once the budget
     * is exceeded, so paging through a long deck does not retain its whole media set.
     */
    load(urls: Iterable<string>, stageUrls: Iterable<string> = []) {
      needed = new Set(urls)
      large = new Set(stageUrls)
      for (const url of needed) {
        touch(url)
        const want = sideFor(url)
        const held = decodedSide.get(url)
        if (held !== undefined && held >= want) continue
        if (held !== undefined) {
          // held at rail size but now drawn on the stage: drop it and decode larger
          const image = loaded.get(url)
          loaded.delete(url)
          decodedSide.delete(url)
          if (image) bytes -= imageBytes(image)
          options.onEvict?.(url)
        }
        if (loading.has(url) || queued.has(url)) continue
        queued.add(url)
        queue.push(url)
      }
      // forget queued work nobody needs any more (the window moved on)
      for (let i = queue.length - 1; i >= 0; i -= 1) {
        const url = queue[i] as string
        if (needed.has(url)) continue
        queue.splice(i, 1)
        queued.delete(url)
      }
      pump()
      if (loading.size === 0 && queue.length === 0) evict()
    },
    /** Forget decoded media (a different deck was opened). */
    clear() {
      loaded.clear()
      decodedSide.clear()
      buf.clear()
      queue.length = 0
      queued.clear()
      bytes = 0
    },
    // Only guards setState after unmount; in-flight loads keep filling `loaded`
    dispose() {
      disposed = true
      if (timer) clearTimeout(timer)
    },
  }
}
