/** Hostile-input fuzz for the PDF import surface — the pdf sibling of the
 *  slides pptx-hostile-fuzz.test.ts and sheets xlsx-sidecar-hostile.test.ts.
 *  The regular PDF tests feed well-formed documents; this one feeds corrupted
 *  bytes to pdf.js (the reader the renderer actually opens untrusted files
 *  with) and asserts the same property: bounded, clean failure. A load either
 *  refuses with an Error or succeeds, never hangs, and a healthy document
 *  still opens afterwards, so a parse that poisons module-level state fails
 *  on its own. No fidelity assertions: corrupted bytes have no right output.
 *
 *  Corruption is seeded (same seed -> same bytes). CI runs the default counts;
 *  HOSTILE_PDF_SEEDS sweeps a larger space locally. */
import { describe, expect, it } from 'vitest'
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs'
import { buildFuzzPdf, PDF_CORRUPTIONS, rng } from './helpers/fuzz-pdf'

const SEEDS = Math.max(1, Number(process.env.HOSTILE_PDF_SEEDS ?? 8))
const LOAD_TIMEOUT_MS = 8000
const LOAD_TIMED_OUT = 'pdf.js load timed out'

/** The outcome of a failed load: a parse error is a clean refusal, a load that
 *  never settles is a timeout and has to be visible as one. */
export function outcomeOf(error: unknown): 'refused' | 'timed-out' {
  return error instanceof Error && error.message === LOAD_TIMED_OUT ? 'timed-out' : 'refused'
}

type PDFDocumentLoadingTask = ReturnType<typeof getDocument>

/** pdf.js owns the buffer, so pass a copy. A timed-out load is destroyed so a
 *  worker cannot leak into the next case; every outcome is caught upstream. */
async function attemptLoad(
  bytes: Uint8Array,
  load: (bytes: Uint8Array) => PDFDocumentLoadingTask = (b) =>
    getDocument({ data: b.slice(), verbosity: 0, useSystemFonts: true }),
  timeoutMs: number = LOAD_TIMEOUT_MS,
): Promise<'opened' | 'refused' | 'timed-out'> {
  let task: PDFDocumentLoadingTask | undefined
  let timer: ReturnType<typeof setTimeout> | undefined
  try {
    task = load(bytes)
    const doc = await Promise.race([
      task.promise,
      new Promise<never>((_reject, reject) => {
        timer = setTimeout(() => reject(new Error(LOAD_TIMED_OUT)), timeoutMs)
      }),
    ])
    // Walking the first page is what a thumbnail/render pass does immediately
    // after open; a document that "opens" but cannot be read is a refusal.
    if (doc.numPages > 0 && doc.numPages <= 64) {
      const page = await doc.getPage(1)
      await page.getTextContent()
      await page.cleanup()
    }
    // pdf.js v6 removed PDFDocumentProxy.destroy(); the loading task owns the
    // worker and is destroyed in the finally block below.
    return 'opened'
  } catch (error) {
    // A load that never settles is not a clean refusal. It has to be its own
    // outcome, or a hang passes as "refuses cleanly ... never hangs".
    return outcomeOf(error)
  } finally {
    if (timer !== undefined) clearTimeout(timer)
    if (task !== undefined) void task.destroy().catch(() => undefined)
  }
}

async function healthyOpens(seed: number): Promise<void> {
  expect(await attemptLoad(await buildFuzzPdf(seed)), `healthy seed ${seed}`).toBe('opened')
}

describe('hostile PDF import fuzz', () => {
  it('reports a load that never settles as a timeout, not a clean refusal', async () => {
    // A hang must not pass as "refuses cleanly ... never hangs": only the
    // suite's own budget ends the load, and the case has to show that.
    const outcome = await attemptLoad(
      new Uint8Array([0x25, 0x50, 0x44, 0x46]),
      () => ({ promise: new Promise(() => {}), destroy: async () => undefined }) as never,
      50,
    )
    expect(outcome).toBe('timed-out')
  })

  it('the generator produces documents pdf.js accepts', async () => {
    for (let seed = 1; seed <= 5; seed++) await healthyOpens(seed)
  })

  for (const [mode, corrupt] of PDF_CORRUPTIONS) {
    describe(mode, () => {
      for (let seed = 1; seed <= SEEDS; seed++) {
        it(`seed ${seed}: load refuses cleanly or succeeds, never hangs`, async () => {
          const valid = await buildFuzzPdf(seed)
          const bytes = corrupt(valid, rng(seed * 7919 + 31))
          const outcome = await attemptLoad(bytes)
          expect(['refused', 'opened']).toContain(outcome)
        }, 20000)
      }

      it('a healthy document still opens after the hostile batch', async () => {
        await healthyOpens(101)
      })
    })
  }

  // Fixed resource constructions: a reader of untrusted bytes must answer
  // each of these inside its normal budget rather than allocating forever.
  describe('resource constructions', () => {
    it('an object flood opens or refuses without hanging', async () => {
      const parts: string[] = ['%PDF-1.7\n']
      for (let i = 1; i <= 2000; i++) parts.push(`${i} 0 obj\n<< /Flood ${i} >>\nendobj\n`)
      parts.push('trailer\n<< /Size 2001 /Root 1 0 R >>\n%%EOF\n')
      const bytes = new TextEncoder().encode(parts.join(''))
      expect(['refused', 'opened']).toContain(await attemptLoad(bytes))
    }, 30000)

    it('a deeply nested array opens or refuses without hanging', async () => {
      const valid = await buildFuzzPdf(7)
      const text = Buffer.from(valid).toString('latin1')
      const bomb = `[${'['.repeat(20000)}0${']'.repeat(20000)}]`
      const bytes = Buffer.from(`${text}\n% nested ${bomb}\n`, 'latin1')
      expect(['refused', 'opened']).toContain(await attemptLoad(new Uint8Array(bytes)))
    }, 30000)

    it('a stream declaring a huge length opens or refuses without hanging', async () => {
      const bytes = Buffer.from(
        '%PDF-1.4\n1 0 obj\n<< /Length 2147483647 >>\nstream\nx\nendstream\nendobj\n' +
          'trailer\n<< /Size 2 >>\n%%EOF\n',
        'latin1',
      )
      expect(['refused', 'opened']).toContain(await attemptLoad(new Uint8Array(bytes)))
    }, 30000)
  })
})
