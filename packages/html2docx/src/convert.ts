/**
 * Orchestration: render the HTML in the host's browser, extract the intent
 * tree, screenshot image-like elements, generate the DOCX.
 */
import type { BrowserDriver, ClipRect } from './driver'
import { BROWSER_HELPER_SCRIPTS, EXTRACTOR_CALL } from './extract'
import { generateDocx } from './generate'
import { FONT_RULES } from './generate/fonts'

export type ConvertStage = 'load' | 'extract' | 'screenshot' | 'generate'

export interface ConvertOptions {
  onProgress?: (progress: { stage: ConvertStage; pct: number }) => void
  signal?: AbortSignal
  log?: (message: string) => void
  /** Debug: in-page probes record `__h2dTrace` entries for elements matching this selector. */
  traceSelector?: string
  /** Debug: receives the extracted IR and trace probes before generation. */
  onIr?: (ir: any[], trace: unknown[]) => void
  /** Keep tables at their rendered width instead of filling the content
   *  width (how Word lays out imported HTML chunks). */
  naturalTableWidth?: boolean
}

export interface ConvertResult {
  docx: Uint8Array
  ir: ValidatedIr[]
  /** Text that only survives as pixels (inside screenshots); lets evaluations
   *  tell "intentionally rasterized" from "actually lost". */
  screenshotText: string
  stats: { screenshots: number; rasterizedDocumentText: boolean }
}

/** Single validated top-level IR node produced by the in-page extractor. */
export interface ValidatedIr {
  type: string
  shotId?: string
  [key: string]: unknown
}

/** A4 at 96dpi so layout (line wraps, column gaps) matches print. */
const A4_VIEWPORT = { width: 794, height: 1123, deviceScaleFactor: 2 }
const NAVIGATION_TIMEOUT_MS = 30000

const sleep = (ms: number) => new Promise((resolve) => setTimeout(resolve, ms))

function createAbortError(): DOMException {
  return new DOMException('html2docx conversion aborted', 'AbortError')
}

function throwIfAborted(signal?: AbortSignal): void {
  if (signal?.aborted) throw createAbortError()
}

/**
 * Validate raw page output before the converter touches it. Page JavaScript
 * can be compromised or return an unexpected shape, so reject non-arrays,
 * null items, and nodes without a string type instead of failing later
 * with a confusing TypeError.
 */
export function normalizeIr(raw: unknown): ValidatedIr[] {
  if (!Array.isArray(raw)) {
    const received = raw === null ? 'null' : typeof raw
    throw new Error(
      `html2docx: extractor returned malformed IR (expected an array, got ${received}); ` +
        `refusing to continue with compromised page JavaScript`,
    )
  }
  return raw.map((item, index) => {
    if (typeof item !== 'object' || item === null) {
      throw new Error(
        `html2docx: extractor returned malformed IR at index ${index} ` +
          `(expected an object node); refusing to continue with compromised page JavaScript`,
      )
    }
    const candidate = item as Record<string, unknown>
    if (typeof candidate.type !== 'string' || candidate.type.length === 0) {
      throw new Error(
        `html2docx: extractor returned malformed IR at index ${index} ` +
          `(missing string "type"); refusing to continue with compromised page JavaScript`,
      )
    }
    if (
      candidate.shotId !== undefined &&
      candidate.shotId !== null &&
      (typeof candidate.shotId !== 'string' || candidate.shotId.length === 0)
    ) {
      throw new Error(
        `html2docx: extractor returned malformed IR at index ${index} ` +
          `(invalid "shotId"); refusing to continue with compromised page JavaScript`,
      )
    }
    // Numeric geometry flows from page JS into division (renderer scale =
    // maxPx / node.width) and image dimensions. Reject NaN/Infinity/negative
    // here so a hostile page cannot produce corrupt-geometry docx.
    for (const key of ['width', 'height', 'widthFrac', 'heightPx', 'xPx', 'yPx'] as const) {
      const v = candidate[key]
      if (v !== undefined && v !== null) {
        const n = Number(v)
        const signed = key === 'xPx' || key === 'yPx'
        if (!Number.isFinite(n) || (!signed && n < 0)) {
          throw new Error(
            `html2docx: extractor returned malformed IR at index ${index} ` +
              `(invalid numeric "${key}"); refusing to continue with compromised page JavaScript`,
          )
        }
      }
    }
    return item as ValidatedIr
  })
}

/**
 * Race a driver promise against an AbortSignal so aborting during a slow
 * page wait (image decode with a long timeout) rejects immediately instead
 * of hanging until the page-side timeout or finish event.
 */
export function raceWithAbort<T>(promise: Promise<T>, signal?: AbortSignal): Promise<T> {
  if (!signal) return promise
  if (signal.aborted) return Promise.reject(createAbortError())
  return new Promise<T>((resolve, reject) => {
    const onAbort = () => reject(createAbortError())
    signal.addEventListener('abort', onAbort, { once: true })
    promise.then(
      (value) => {
        signal.removeEventListener('abort', onAbort)
        resolve(value)
      },
      (error) => {
        signal.removeEventListener('abort', onAbort)
        reject(error)
      },
    )
  })
}

export async function convertHtmlToDocx(
  input: { url: string },
  driver: BrowserDriver,
  options: ConvertOptions = {},
): Promise<ConvertResult> {
  const { signal, traceSelector, onIr } = options
  const log = options.log ?? (() => {})
  const progress = (stage: ConvertStage, pct: number) => options.onProgress?.({ stage, pct })
  const { url } = input

  progress('load', 0)
  await driver.setViewport(A4_VIEWPORT)
  if (traceSelector) {
    await driver.addInitScript(`globalThis.__h2dTraceSelector = ${JSON.stringify(traceSelector)};`)
  }
  log(`[html2docx] loading ${url}`)
  await driver.goto(url, { timeoutMs: NAVIGATION_TIMEOUT_MS })
  await driver.evaluate(() => document.fonts.ready.catch(() => {}))
  await sleep(300)
  throwIfAborted(signal)

  // Authoring templates commonly use an 880px design canvas and are then
  // scaled to A4. Extract at that authored width so narrower A4 reflow does
  // not create artificial wraps and inflated table-row heights.
  const authoredWidth = await driver.evaluate<number | null>(() => {
    // Some document templates constrain body directly; responsive webpages
    // more often constrain repeated `.inner`/`.container` wrappers. At the
    // initial A4 viewport those pages may already have crossed into their
    // mobile breakpoint, destroying desktop grids before extraction.
    const candidates = [document.body, ...document.body.querySelectorAll('*')]
    const widths = candidates.flatMap((el) => {
      const rect = el.getBoundingClientRect()
      const maxWidth = parseFloat(getComputedStyle(el).maxWidth)
      return Number.isFinite(maxWidth) &&
        maxWidth > window.innerWidth &&
        maxWidth <= 1600 &&
        rect.width >= window.innerWidth * 0.55 &&
        rect.height >= 20
        ? [maxWidth]
        : []
    })
    return widths.length ? Math.round(Math.max(...widths)) : null
  })
  // Content that overflows the A4 viewport horizontally (a 6-column table
  // whose last column collapses to 1px) renders broken at 794px while the
  // evaluation baseline at 1024px shows it intact. Reload wider whenever
  // the layout visibly does not fit.
  const overflowsA4 =
    !authoredWidth &&
    (await driver.evaluate<boolean>(() => {
      if (document.documentElement.scrollWidth > window.innerWidth + 8) return true
      if (
        [...document.querySelectorAll('table')].some(
          (table) => table.scrollWidth > table.clientWidth + 4,
        )
      ) {
        return true
      }
      // table-layout:auto absorbs overflow by crushing cells instead of
      // scrolling — a cell whose content no longer fits on one word is the
      // telltale ("Wave 1" pills stacking letter-by-letter).
      return [...document.querySelectorAll('td, th')].some(
        (cell) => cell.scrollWidth > cell.clientWidth + 2,
      )
    }))
  if (authoredWidth || overflowsA4) {
    const renderWidth = Math.max(1024, authoredWidth || 0)
    await driver.setViewport({
      width: renderWidth,
      height: Math.round((1123 * renderWidth) / 794),
      deviceScaleFactor: 2,
    })
    await driver.goto(url, { timeoutMs: NAVIGATION_TIMEOUT_MS })
    await driver.evaluate(() => document.fonts.ready.catch(() => {}))
    await sleep(300)
  }
  throwIfAborted(signal)
  progress('load', 60)

  // Offscreen website images may remain deferred even without an explicit
  // loading="lazy" attribute. Visit the full page before screenshot-based
  // extraction, then wait for every reachable image to decode. Race the
  // page-side wait (which has a long per-image timeout) against the caller
  // abort signal so aborting during load rejects immediately.
  await raceWithAbort(
    driver.evaluate(async () => {
      const wait = (ms: number) => new Promise((resolve) => setTimeout(resolve, ms))
      const originalY = window.scrollY
      const htmlScrollBehavior = document.documentElement.style.scrollBehavior
      const bodyScrollBehavior = document.body.style.scrollBehavior
      document.documentElement.style.scrollBehavior = 'auto'
      document.body.style.scrollBehavior = 'auto'
      const maxY = Math.max(0, document.documentElement.scrollHeight - window.innerHeight)
      const step = Math.max(400, Math.round(window.innerHeight * 0.75))
      for (let y = 0; y <= maxY; y += step) {
        window.scrollTo(0, y)
        await wait(60)
      }
      window.scrollTo(0, maxY)
      await wait(100)
      const imageWaits = [...document.images].map((image) => {
        if (image.complete && image.naturalWidth > 0) {
          return image.decode?.().catch(() => {})
        }
        return new Promise<void>((resolve) => {
          const finish = () => resolve()
          image.addEventListener('load', finish, { once: true })
          image.addEventListener('error', finish, { once: true })
          setTimeout(finish, 15000)
        })
      })
      await Promise.all(imageWaits)
      window.scrollTo(0, originalY)
      await wait(100)
      document.documentElement.style.scrollBehavior = htmlScrollBehavior
      document.body.style.scrollBehavior = bodyScrollBehavior
    }),
    signal,
  )
  await sleep(200)
  throwIfAborted(signal)
  progress('load', 100)

  // Word substitutes unavailable webfonts (Cormorant → Georgia, Playfair →
  // Georgia, …) with often much wider glyphs. Expose a measurement helper
  // that re-measures a one-line label under the substituted font so width
  // floors reflect what Word will render, not the authored webfont.
  await driver.evaluate(
    (rules: [string, string, string][]) => {
      const parsed = rules.map(
        ([source, flags, name]) => [new RegExp(source, flags), name] as [RegExp, string],
      )
      const w = window as any
      w.__h2dWordFontLineWidth = (el: HTMLElement) => {
        let family: string
        try {
          family = getComputedStyle(el).fontFamily || ''
        } catch {
          return 0
        }
        let mapped = 'Arial'
        for (const [re, name] of parsed) {
          if (re.test(family)) {
            mapped = name
            break
          }
        }
        const prevFont = el.style.fontFamily
        const prevWhiteSpace = el.style.whiteSpace
        let width = 0
        try {
          el.style.fontFamily = mapped
          el.style.whiteSpace = 'nowrap'
          const range = document.createRange()
          range.selectNodeContents(el)
          width = range.getBoundingClientRect().width
        } catch {
          /* keep 0 */
        }
        el.style.fontFamily = prevFont || ''
        el.style.whiteSpace = prevWhiteSpace || ''
        return width
      }
    },
    FONT_RULES.map(([re, name]) => [re.source, re.flags, name]),
  )
  for (const source of BROWSER_HELPER_SCRIPTS) {
    await driver.addScript(source)
  }
  log('[html2docx] extracting intent tree')
  progress('extract', 0)
  const rawIr = await driver.evaluate<unknown>(EXTRACTOR_CALL)
  // Page JavaScript may be compromised, so validate the shape before use.
  const ir = normalizeIr(rawIr)
  if (onIr) {
    const trace = await driver.evaluate<unknown[]>(() => (globalThis as any).__h2dTrace || [])
    onIr(ir, trace)
  }
  log(`[html2docx] extracted ${ir.length} top-level nodes`)
  throwIfAborted(signal)
  progress('extract', 100)

  // Broken images render Chrome's placeholder chrome (icon + clipped alt
  // text) — never authored design. Hide them so they can't bleed into any
  // element screenshot (extraction already emitted their alt-text fallback).
  await driver.evaluate(() => {
    for (const img of document.querySelectorAll('img')) {
      if (!img.complete || img.naturalWidth === 0) img.style.visibility = 'hidden'
    }
  })

  // screenshot every element marked as image-like
  const images: Record<string, Uint8Array> = {}
  const shotIds: string[] = []
  const shotNodes = new Map<string, any>()
  const pageBgShotIds = new Set<string>(
    ir.flatMap((node) => (node.type === 'pagebg' && node.shotId ? [node.shotId] : [])),
  )
  const collectRunShots = (runs: any[] | undefined) => {
    for (const run of runs || []) {
      if (run.inlineImage && run.shotId) {
        shotIds.push(run.shotId)
        if (run.unwrap || run.clip) shotNodes.set(run.shotId, run)
      }
    }
  }
  ;(function collectShots(nodes: any[]) {
    for (const n of nodes) {
      if (n.type === 'image' || n.type === 'floatimg' || (n.type === 'pagebg' && n.shotId)) {
        shotIds.push(n.shotId)
        shotNodes.set(n.shotId, n)
      }
      collectRunShots(n.runs)
      for (const item of n.items || []) collectRunShots(item.runs)
      for (const key of ['children', 'cells']) {
        if (Array.isArray(n[key])) {
          for (const sub of n[key]) {
            collectRunShots(sub.runs)
            if (sub.children) collectShots(sub.children)
          }
          if (key === 'children') collectShots(n[key])
        }
      }
      if (n.rows) {
        for (const row of n.rows) {
          for (const cell of row.cells) {
            collectRunShots(cell.runs)
            if (cell.children) collectShots(cell.children)
          }
        }
      }
    }
  })(ir)

  const uniqueShotIds = [...new Set(shotIds)]
  let shotIndex = 0
  for (const id of uniqueShotIds) {
    throwIfAborted(signal)
    progress('screenshot', Math.round((shotIndex++ / Math.max(1, uniqueShotIds.length)) * 100))
    const shotNode = shotNodes.get(id)
    const selector = `[data-h2d-id="${id}"]`
    const hasElement = shotNode?.clip ? false : await driver.elementExists(selector)
    if (!hasElement && !shotNode?.clip) continue
    try {
      if (shotNode?.unwrap) {
        // Crushed inline label (letters stacked in a too-narrow column):
        // relax to its natural one-line size for the screenshot.
        await driver.evaluate((shotId: string) => {
          const target = document.querySelector<HTMLElement>(`[data-h2d-id="${shotId}"]`)
          if (!target) return
          target.setAttribute('data-h2d-unwrap-style', target.getAttribute('style') || '')
          target.style.whiteSpace = 'nowrap'
          target.style.width = 'max-content'
          target.style.maxWidth = 'none'
        }, id)
      }
      if (shotNode?.isolate) {
        await driver.evaluate((shotId: string) => {
          const target = document.querySelector<HTMLElement>(`[data-h2d-id="${shotId}"]`)
          if (!target) return
          for (const candidate of document.body.querySelectorAll<HTMLElement>('*')) {
            if (candidate === target || candidate.contains(target) || target.contains(candidate)) {
              continue
            }
            candidate.setAttribute('data-h2d-isolate-visibility', candidate.style.visibility || '')
            candidate.style.visibility = 'hidden'
          }
          // surrounding paint must not leak into the shot: the page and
          // the target's ancestors (white content sheet) otherwise give
          // the capture an opaque bounding box that occludes neighboring
          // floats in Word
          let scrub = target.parentElement
          while (scrub) {
            scrub.setAttribute('data-h2d-isolate-bg', scrub.getAttribute('style') || '')
            scrub.style.background = 'transparent'
            scrub = scrub.parentElement
          }
        }, id)
      }
      if (pageBgShotIds.has(id)) {
        await driver.evaluate((shotId: string) => {
          const backdrop = document.querySelector<HTMLElement>(`[data-h2d-id="${shotId}"]`)!
          for (const child of document.body.children as HTMLCollectionOf<HTMLElement>) {
            child.setAttribute('data-h2d-old-visibility', child.style.visibility || '')
            child.style.visibility = child === backdrop ? 'visible' : 'hidden'
          }
          backdrop.setAttribute('data-h2d-old-style', backdrop.getAttribute('style') || '')
          Object.assign(backdrop.style, {
            position: 'fixed',
            left: '0',
            top: '0',
            zIndex: '2147483647',
            visibility: 'visible',
          })
        }, id)
      }
      // Badges/ribbons absolutely positioned to overhang their card's box
      // (top: -12px "MOST VALUE" pills) are cut in half by an element
      // screenshot — expand the capture to the union of the element and
      // its near overhanging children, and resize the node to match.
      let overhang: ClipRect | null = null
      if (hasElement && shotNode?.type === 'image' && !shotNode.pageComposition && !shotNode.clip) {
        overhang = await driver.evaluate<ClipRect | null>((shotId: string) => {
          const target = document.querySelector<HTMLElement>(`[data-h2d-id="${shotId}"]`)
          if (!target) return null
          const ts = getComputedStyle(target)
          if (ts.overflow === 'hidden' || ts.overflow === 'clip') return null
          const r = target.getBoundingClientRect()
          const box = { left: r.left, top: r.top, right: r.right, bottom: r.bottom }
          for (const d of target.querySelectorAll('*')) {
            if (getComputedStyle(d).position !== 'absolute') continue
            const dr = d.getBoundingClientRect()
            if (!dr.width || !dr.height) continue
            // near overhangs only — far-flung decorations stay dropped
            if (
              dr.left < r.left - 48 ||
              dr.right > r.right + 48 ||
              dr.top < r.top - 48 ||
              dr.bottom > r.bottom + 48
            )
              continue
            box.left = Math.min(box.left, dr.left)
            box.top = Math.min(box.top, dr.top)
            box.right = Math.max(box.right, dr.right)
            box.bottom = Math.max(box.bottom, dr.bottom)
          }
          if (
            box.left >= r.left - 1 &&
            box.top >= r.top - 1 &&
            box.right <= r.right + 1 &&
            box.bottom <= r.bottom + 1
          )
            return null
          // rects are viewport-relative and earlier element screenshots
          // may have scrolled the page; the clip needs page coordinates.
          // In RTL the horizontal overflow hides LEFT of the viewport, so
          // page x = viewport x + overflow + scrollLeft (scrollLeft <= 0).
          const doc = document.documentElement
          const rtlShift =
            getComputedStyle(doc).direction === 'rtl' ||
            getComputedStyle(document.body).direction === 'rtl'
              ? Math.max(0, doc.scrollWidth - doc.clientWidth)
              : 0
          const x = Math.max(0, Math.round(box.left + window.scrollX + rtlShift))
          const y = Math.max(0, Math.round(box.top + window.scrollY))
          return {
            x,
            y,
            width: Math.round(box.right - box.left),
            height: Math.round(box.bottom - box.top),
          }
        }, id)
      }
      if (overhang) {
        images[id] = await driver.screenshot({ clip: overhang })
        shotNode.width = overhang.width
        shotNode.height = overhang.height
      } else if (shotNode?.clip) {
        images[id] = await driver.screenshot({ clip: shotNode.clip })
      } else {
        // Element screenshots hand viewport coordinates to CDP, whose clip
        // origin is the LEFTMOST scrollable content. An RTL document with
        // horizontal overflow hides that overflow left of the viewport, so
        // the two origins differ by the overflow width and the capture
        // lands shifted (TOC numbers / line starts cut off). Take an
        // explicit page-coordinate clip instead.
        const rtlBox = await driver.evaluate<ClipRect | null>((shotId: string) => {
          const doc = document.documentElement
          const rtl =
            getComputedStyle(doc).direction === 'rtl' ||
            getComputedStyle(document.body).direction === 'rtl'
          const shift = rtl ? Math.max(0, doc.scrollWidth - doc.clientWidth) : 0
          if (!shift) return null
          const target = document.querySelector(`[data-h2d-id="${shotId}"]`)
          if (!target) return null
          const r = target.getBoundingClientRect()
          return {
            x: Math.max(0, Math.round(r.left + shift + doc.scrollLeft)),
            y: Math.max(0, Math.round(r.top + window.scrollY)),
            width: Math.round(r.width),
            height: Math.round(r.height),
          }
        }, id)
        const omitBackground = Boolean(shotNode?.isolate)
        const shot = rtlBox
          ? await driver.screenshot({ clip: rtlBox, omitBackground })
          : await driver.screenshotElement(selector, { omitBackground })
        if (shot) images[id] = shot
      }
      if (images[id]) log(`[html2docx] screenshot ${id}: ${images[id].length} bytes`)
    } catch (e) {
      log(`[html2docx] screenshot ${id} failed: ${e instanceof Error ? e.message : String(e)}`)
    } finally {
      // Each restore runs independently so one failure cannot skip the
      // remaining restores or leave the page mutated for the next shot.
      if (shotNode?.unwrap) {
        try {
          await driver.evaluate((shotId: string) => {
            const target = document.querySelector(`[data-h2d-id="${shotId}"]`)
            if (!target) return
            target.setAttribute('style', target.getAttribute('data-h2d-unwrap-style') || '')
            target.removeAttribute('data-h2d-unwrap-style')
          }, id)
        } catch (e) {
          log(
            `[html2docx] restore unwrap ${id} failed: ${e instanceof Error ? e.message : String(e)}`,
          )
        }
      }
      if (shotNode?.isolate) {
        try {
          await driver.evaluate(() => {
            for (const candidate of document.querySelectorAll<HTMLElement>(
              '[data-h2d-isolate-visibility]',
            )) {
              candidate.style.visibility =
                candidate.getAttribute('data-h2d-isolate-visibility') || ''
              candidate.removeAttribute('data-h2d-isolate-visibility')
            }
            for (const scrubbed of document.querySelectorAll('[data-h2d-isolate-bg]')) {
              scrubbed.setAttribute('style', scrubbed.getAttribute('data-h2d-isolate-bg') || '')
              scrubbed.removeAttribute('data-h2d-isolate-bg')
            }
          })
        } catch (e) {
          log(
            `[html2docx] restore isolate ${id} failed: ${e instanceof Error ? e.message : String(e)}`,
          )
        }
      }
      if (pageBgShotIds.has(id)) {
        try {
          await driver.evaluate((shotId: string) => {
            const backdrop = document.querySelector(`[data-h2d-id="${shotId}"]`)
            if (backdrop) {
              backdrop.setAttribute('style', backdrop.getAttribute('data-h2d-old-style') || '')
              backdrop.removeAttribute('data-h2d-old-style')
            }
            for (const child of document.body.children as HTMLCollectionOf<HTMLElement>) {
              child.style.visibility = child.getAttribute('data-h2d-old-visibility') || ''
              child.removeAttribute('data-h2d-old-visibility')
            }
          }, id)
        } catch (e) {
          log(
            `[html2docx] restore pagebg ${id} failed: ${e instanceof Error ? e.message : String(e)}`,
          )
        }
      }
    }
  }
  progress('screenshot', 100)

  // Text living inside rasterized elements: it is present in the DOCX as
  // pixels, not as runs. Exported for evaluation so text-coverage checks
  // can distinguish "intentionally rasterized" from "actually lost".
  const rasterizedDocumentText = [...shotNodes.values()].some((node) => node.capturesDocumentText)
  const screenshotText = await driver.evaluate<string>((includeBodyText: boolean) => {
    const parts: string[] = []
    if (includeBodyText) {
      const bodyText = (document.body.innerText || '').trim()
      if (bodyText) parts.push(bodyText)
    }
    for (const el of document.querySelectorAll<HTMLElement>('[data-h2d-id]')) {
      const text = (
        el.innerText ||
        (el.tagName === 'IMG' ? el.getAttribute('alt') : '') ||
        ''
      ).trim()
      if (text) parts.push(text)
    }
    return parts.join('\n')
  }, rasterizedDocumentText)

  throwIfAborted(signal)
  log('[html2docx] generating docx')
  progress('generate', 0)
  const docx = new Uint8Array(
    await generateDocx(ir, images, { naturalTableWidth: options.naturalTableWidth }),
  )
  log(`[html2docx] generated ${docx.length} bytes`)
  progress('generate', 100)
  return {
    docx,
    ir,
    screenshotText,
    stats: { screenshots: Object.keys(images).length, rasterizedDocumentText },
  }
}
