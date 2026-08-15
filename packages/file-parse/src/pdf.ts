import { createRequire } from 'node:module'
import { dirname, join } from 'node:path'

// pdfjs needs the standard_fonts data directory for non-embedded standard fonts
// (Helvetica etc.); under Node a filesystem path works (same usage as the official
// Node example). The packaged build may fail require.resolve (only out/** is bundled),
// so return undefined and omit it — degrading to "non-embedded standard fonts may be
// slightly incomplete" (most embedded-font PDFs unaffected) instead of crashing parsing.
function standardFontDataUrl(): string | undefined {
  try {
    const require = createRequire(import.meta.url)
    const pdfPath = require.resolve('pdfjs-dist/legacy/build/pdf.mjs')
    return `${join(dirname(pdfPath), '..', '..', 'standard_fonts')}/`
  } catch {
    return undefined
  }
}

/**
 * pdfjs's Node compat layer borrows DOMMatrix from the optional dep @napi-rs/canvas,
 * and pdf.mjs calls `new DOMMatrix()` at module top level — in the packaged build
 * (no node_modules inside asar) the require fails, so the import throws
 * "DOMMatrix is not defined" and PDF attachment parsing breaks entirely (not
 * reproducible in dev since the dep happens to be present). Text extraction only
 * needs a tiny subset of 2D affine matrices, so ship a pure-JS fallback that
 * doesn't depend on packaging layout.
 */
function installDomMatrixPolyfill(): void {
  const g = globalThis as { DOMMatrix?: unknown }
  if (g.DOMMatrix) return
  class DOMMatrixPolyfill {
    a = 1
    b = 0
    c = 0
    d = 1
    e = 0
    f = 0
    constructor(init?: number[] | DOMMatrixPolyfill) {
      if (Array.isArray(init) && init.length >= 6) {
        ;[this.a, this.b, this.c, this.d, this.e, this.f] = init as [
          number,
          number,
          number,
          number,
          number,
          number,
        ]
      } else if (init && typeof init === 'object') {
        const m = init as DOMMatrixPolyfill
        this.a = m.a
        this.b = m.b
        this.c = m.c
        this.d = m.d
        this.e = m.e
        this.f = m.f
      }
    }
    get is2D(): boolean {
      return true
    }
    get isIdentity(): boolean {
      return (
        this.a === 1 && this.b === 0 && this.c === 0 && this.d === 1 && this.e === 0 && this.f === 0
      )
    }
    /** this × other (DOM spec semantics: result applied to a point = this(other(p))) */
    #product(o: DOMMatrixPolyfill): [number, number, number, number, number, number] {
      return [
        this.a * o.a + this.c * o.b,
        this.b * o.a + this.d * o.b,
        this.a * o.c + this.c * o.d,
        this.b * o.c + this.d * o.d,
        this.a * o.e + this.c * o.f + this.e,
        this.b * o.e + this.d * o.f + this.f,
      ]
    }
    #assign(v: [number, number, number, number, number, number]): this {
      ;[this.a, this.b, this.c, this.d, this.e, this.f] = v
      return this
    }
    multiply(o: DOMMatrixPolyfill): DOMMatrixPolyfill {
      return new DOMMatrixPolyfill(this.#product(o))
    }
    multiplySelf(o: DOMMatrixPolyfill): this {
      return this.#assign(this.#product(o))
    }
    preMultiplySelf(o: DOMMatrixPolyfill): this {
      return this.#assign(o.#product(this))
    }
    translate(tx = 0, ty = 0): DOMMatrixPolyfill {
      return this.multiply(new DOMMatrixPolyfill([1, 0, 0, 1, tx, ty]))
    }
    translateSelf(tx = 0, ty = 0): this {
      return this.multiplySelf(new DOMMatrixPolyfill([1, 0, 0, 1, tx, ty]))
    }
    scale(sx = 1, sy?: number): DOMMatrixPolyfill {
      return this.multiply(new DOMMatrixPolyfill([sx, 0, 0, sy ?? sx, 0, 0]))
    }
    scaleSelf(sx = 1, sy?: number): this {
      return this.multiplySelf(new DOMMatrixPolyfill([sx, 0, 0, sy ?? sx, 0, 0]))
    }
    invertSelf(): this {
      const { a, b, c, d, e, f } = this
      const det = a * d - b * c
      if (!det || !Number.isFinite(det)) return this.#assign([NaN, NaN, NaN, NaN, NaN, NaN])
      return this.#assign([
        d / det,
        -b / det,
        -c / det,
        a / det,
        (c * f - d * e) / det,
        (b * e - a * f) / det,
      ])
    }
    inverse(): DOMMatrixPolyfill {
      return new DOMMatrixPolyfill(this).invertSelf()
    }
    transformPoint(p: { x?: number; y?: number } = {}): {
      x: number
      y: number
      z: number
      w: number
    } {
      const x = p.x ?? 0
      const y = p.y ?? 0
      return {
        x: this.a * x + this.c * y + this.e,
        y: this.b * x + this.d * y + this.f,
        z: 0,
        w: 1,
      }
    }
  }
  g.DOMMatrix = DOMMatrixPolyfill
}

/** extract text from a pdf with pdfjs-dist (pure JS, no native deps), one section per page */
export async function pdfToText(bytes: Uint8Array): Promise<string> {
  installDomMatrixPolyfill()
  // Explicitly import the worker module (its top level registers globalThis.pdfjsWorker,
  // which the fake worker prefers) — otherwise pdfjs looks up pdf.worker.mjs by path at
  // runtime, the file isn't next to the bundled chunk, and it fails with "Setting up fake
  // worker failed". A literal specifier lets the bundler include it in the output.
  // @ts-expect-error the worker build artifact has no type declarations; imported only for its top-level side effect (registering globalThis.pdfjsWorker)
  await import('pdfjs-dist/legacy/build/pdf.worker.mjs')
  const { getDocument } = await import('pdfjs-dist/legacy/build/pdf.mjs')
  const fontUrl = standardFontDataUrl()
  // pdfjs-dist 6.x removed PDFDocumentProxy.destroy(); cleanup goes through the loading task
  const loadingTask = getDocument({
    // pdfjs transfers the given buffer, so pass a copy
    data: new Uint8Array(bytes),
    useSystemFonts: true,
    ...(fontUrl ? { standardFontDataUrl: fontUrl } : {}),
    verbosity: 0,
  })
  const doc = await loadingTask.promise
  try {
    const pages: string[] = []
    for (let i = 1; i <= doc.numPages; i++) {
      const page = await doc.getPage(i)
      const content = await page.getTextContent()
      let text = ''
      for (const item of content.items) {
        if ('str' in item) {
          text += item.str
          if (item.hasEOL) text += '\n'
        }
      }
      pages.push(text.trim())
      page.cleanup()
    }
    return pages.join('\n\n')
  } finally {
    await loadingTask.destroy()
  }
}
