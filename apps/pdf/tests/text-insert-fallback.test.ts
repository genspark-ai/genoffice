import { describe, expect, it, vi } from 'vitest'

// Simulate the Windows condition that broke every Insert Text save: none of the
// FALLBACK_FONT_PATHS files exist (arialuni.ttf only ships with legacy Office).
// Everything else reads normally, so the pdfium wasm and the installed-font index
// still work — exactly the shape of a stock Windows machine.
vi.mock('node:fs', async (importOriginal) => {
  const real = await importOriginal<typeof import('node:fs')>()
  const gone = ['Arial Unicode.ttf', 'arialuni.ttf', 'DejaVuSans.ttf']
  const readFileSync = ((path: unknown, ...rest: unknown[]) => {
    if (typeof path === 'string' && gone.some((g) => path.endsWith(g))) {
      throw Object.assign(new Error(`ENOENT: ${path}`), { code: 'ENOENT' })
    }
    return (real.readFileSync as (...a: unknown[]) => unknown)(path, ...rest)
  }) as typeof real.readFileSync
  return { ...real, readFileSync, default: { ...real, readFileSync } }
})

import { PDFDocument } from 'pdf-lib'
import { applyTextInserts, fallbackFontFor } from '../src/main/text-edit'
import type { TextInsertInput } from '../src/shared/ipc'

async function blankPage(): Promise<Uint8Array> {
  const doc = await PDFDocument.create()
  doc.addPage([595, 842])
  return doc.save({ useObjectStreams: false })
}

const insert = (text: string): TextInsertInput => ({
  pageIndex: 0,
  origin: [50, 700],
  text,
  fontSize: 14,
  color: [0, 0, 0],
})

describe('insert-text fallback without any fallback font file (Windows shape)', () => {
  it('resolves a face through the installed-font index instead of failing the insert', async () => {
    // The old single-file fallback threw 'no fallback font available' here, which
    // skipped EVERY insert ("Inserted text could not be saved on page(s): …")
    expect(fallbackFontFor('Latin insert')).not.toBeNull()
    const result = await applyTextInserts(await blankPage(), [insert('Latin fallback insert')])
    expect(result.skipped).toEqual([])
  })

  it('covers CJK inserts through a CJK-capable installed face', async () => {
    if (fallbackFontFor('插入中文') === null) return // machine has no CJK face at all
    const result = await applyTextInserts(await blankPage(), [insert('插入中文文本')])
    expect(result.skipped).toEqual([])
  })

  it('rescues text outside every curated candidate via the full-index scan', async () => {
    // ⌘ (U+2318) is not in Arial/Times/etc.; on any desktop some installed face maps
    // it — the same face the browser preview used to display it
    const covering = fallbackFontFor('⌘')
    if (covering === null) return // bare CI container without such a face
    const result = await applyTextInserts(await blankPage(), [insert('Press ⌘ to save')])
    expect(result.skipped).toEqual([])
  })

  it('reports the reason when no installed face covers the text', async () => {
    // Unassigned codepoint U+0378: no real font maps it (LastResort-style fonts use
    // cmap format 13, which the coverage reader deliberately does not treat as real
    // coverage), so the insert must skip with the no-font reason instead of embedding
    // .notdef boxes
    const result = await applyTextInserts(await blankPage(), [insert('bad \u0378 char')])
    expect(result.skipped).toHaveLength(1)
    expect(result.skipped[0]!.reason).toContain('no available font')
  })
})
