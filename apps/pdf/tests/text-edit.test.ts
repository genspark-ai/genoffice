import { existsSync } from 'node:fs'
import { PDFDocument, StandardFonts } from 'pdf-lib'
import { describe, expect, it } from 'vitest'
import {
  applyTextEdits,
  applyTextInserts,
  mergeEngineCodepoints,
  textInsertAxes,
  validateTextEdits,
} from '../src/main/text-edit'
import type { TextEditInput, TextInsertInput } from '../src/shared/ipc'

/** Happy-path apply: no edit may be skipped */
async function applyAll(bytes: Uint8Array, edits: TextEditInput[]): Promise<Uint8Array> {
  const result = await applyTextEdits(bytes, edits)
  expect(result.skipped).toEqual([])
  return result.bytes
}

/** Page text of the first page via pdf.js (pdfium writes the file, a second engine reads it) */
async function extractText(bytes: Uint8Array): Promise<string> {
  const { getDocument } = await import('pdfjs-dist/legacy/build/pdf.mjs')
  const doc = await getDocument({ data: bytes.slice(), useSystemFonts: true }).promise
  try {
    const page = await doc.getPage(1)
    const content = await page.getTextContent()
    return content.items.map((i) => ('str' in i ? i.str : '')).join('')
  } finally {
    await doc.loadingTask.destroy()
  }
}

interface Fixture {
  bytes: Uint8Array
  rect: [number, number, number, number]
}

/** One-page PDF with a single Helvetica text run at a known position */
async function makeFixture(text: string): Promise<Fixture> {
  const doc = await PDFDocument.create()
  const page = doc.addPage([595, 842])
  const font = await doc.embedFont(StandardFonts.Helvetica)
  const size = 14
  page.drawText(text, { x: 50, y: 700, size, font })
  const w = font.widthOfTextAtSize(text, size)
  return {
    bytes: await doc.save({ useObjectStreams: false }),
    rect: [45, 694, 50 + w + 5, 700 + size + 4],
  }
}

const edit = (f: Fixture, oldText: string, newText: string): TextEditInput => ({
  pageIndex: 0,
  rect: f.rect,
  oldText,
  newText,
  fontSize: 14,
})

describe('applyTextInserts', () => {
  it('counter-rotates text axes for every page rotation', () => {
    expect(textInsertAxes(0)).toEqual([1, 0, 0, 1])
    expect(textInsertAxes(90)).toEqual([0, 1, -1, 0])
    expect(textInsertAxes(180)).toEqual([-1, 0, 0, -1])
    expect(textInsertAxes(270)).toEqual([0, -1, 1, 0])
    expect(textInsertAxes(-90)).toEqual([0, -1, 1, 0])
  })

  it('adds searchable text without replacing existing page content', async () => {
    const f = await makeFixture('Existing text')
    const insert: TextInsertInput = {
      pageIndex: 0,
      origin: [50, 650],
      text: 'New searchable text',
      fontSize: 14,
      color: [0, 0, 0],
    }
    const result = await applyTextInserts(f.bytes, [insert])
    expect(result.skipped).toEqual([])
    expect(await extractText(result.bytes)).toContain('Existing text')
    expect(await extractText(result.bytes)).toContain('New searchable text')
  })

  it('embeds a subset font for inserted CJK text', async () => {
    if (!existsSync('/System/Library/Fonts/Supplemental/Arial Unicode.ttf')) return
    const f = await makeFixture('Existing text')
    const result = await applyTextInserts(f.bytes, [
      {
        pageIndex: 0,
        origin: [50, 650],
        text: '插入中文文本',
        fontSize: 14,
        color: [0, 0, 0],
      },
    ])
    expect(result.skipped).toEqual([])
    expect(await extractText(result.bytes)).toContain('插入中文文本')
    expect(result.bytes.length).toBeLessThan(200 * 1024)
  })
})

describe('applyTextEdits translate (block move)', () => {
  it('moves a whole run by translating its objects without rewriting them', async () => {
    const f = await makeFixture('Move me somewhere else')
    const probe = edit(f, 'Move me somewhere else', 'Move me somewhere else')
    const [before] = await validateTextEdits(f.bytes, [probe])
    expect(before!.reason).toBeNull()
    const out = await applyAll(f.bytes, [{ ...probe, translate: [30, -40] }])
    expect(await extractText(out)).toBe('Move me somewhere else')
    const shifted: TextEditInput = {
      ...probe,
      rect: [f.rect[0] + 30, f.rect[1] - 40, f.rect[2] + 30, f.rect[3] - 40],
    }
    const [after] = await validateTextEdits(out, [shifted])
    expect(after!.reason).toBeNull()
    for (const [i, delta] of [30, -40, 30, -40].entries()) {
      expect(Math.abs(after!.bounds![i]! - (before!.bounds![i]! + delta))).toBeLessThan(1)
    }
    // Nothing was rebuilt: no font program was embedded by the move
    expect(out.length).toBeLessThan(f.bytes.length + 2048)
  })

  it('refuses to move a fragment of a larger run', async () => {
    const f = await makeFixture('Alpha beta gamma')
    const move: TextEditInput = { ...edit(f, 'beta', 'beta'), translate: [10, 10] }
    const [v] = await validateTextEdits(f.bytes, [move])
    expect(v!.reason).toMatch(/moved as one unit/)
    const result = await applyTextEdits(f.bytes, [move])
    expect(result.skipped).toHaveLength(1)
    expect(result.skipped[0]!.reason).toMatch(/moved as one unit/)
    expect(await extractText(result.bytes)).toBe('Alpha beta gamma')
  })
})

describe('applyTextEdits', () => {
  it('replaces text in place when the font already covers it', async () => {
    const f = await makeFixture('Total revenue was 1500 dollars')
    const out = await applyAll(f.bytes, [
      edit(f, 'Total revenue was 1500 dollars', 'Total revenue was 5100 dollars'),
    ])
    expect(await extractText(out)).toBe('Total revenue was 5100 dollars')
  })

  it('never inherits a zero fill alpha into the rebuilt run', async () => {
    if (!existsSync('/System/Library/Fonts/Supplemental/Arial Unicode.ttf')) return
    // Some producers (Chrome print) leave pdfium's fill-alpha reading at 0 for text
    // that renders opaque; drawing with opacity 0 reproduces the same read
    const doc = await PDFDocument.create()
    const page = doc.addPage([595, 842])
    const font = await doc.embedFont(StandardFonts.Helvetica)
    page.drawText('Alpha zero run', { x: 50, y: 700, size: 14, font, opacity: 0 })
    const w = font.widthOfTextAtSize('Alpha zero run', 14)
    const bytes = await doc.save({ useObjectStreams: false })
    const rect: [number, number, number, number] = [45, 694, 50 + w + 5, 718]
    // the CJK glyph in the new text forces the rebuild path (outside the WinAnsi reuse gate)
    const out = await applyAll(bytes, [
      { pageIndex: 0, rect, oldText: 'Alpha zero run', newText: 'Alpha 应 run', fontSize: 14 },
    ])
    const { chainPdfium, loadPdfium, withDocument } = await import('../src/main/text-edit')
    const m = await loadPdfium()
    const alphas = await chainPdfium(() =>
      withDocument(m, out, async (d) => {
        const p = m._FPDF_LoadPage(d, 0)
        const colPtr = m._malloc(16)
        const found: number[] = []
        const count = m._FPDFPage_CountObjects(p)
        for (let i = 0; i < count; i++) {
          const obj = m._FPDFPage_GetObject(p, i)
          if (m._FPDFPageObj_GetType(obj) !== 1) continue
          if (m._FPDFPageObj_GetFillColor(obj, colPtr, colPtr + 4, colPtr + 8, colPtr + 12)) {
            found.push(m.HEAPU8[colPtr + 12]!)
          }
        }
        m._free(colPtr)
        m._FPDF_ClosePage(p)
        return found
      }),
    )
    // The rebuilt objects must be opaque — an inherited alpha 0 drew invisible text
    expect(alphas.length).toBeGreaterThan(0)
    for (const a of alphas) expect(a).toBe(255)
  })

  it('rebuilds the run with an embedded subset font for chars outside the original', async () => {
    if (!existsSync('/System/Library/Fonts/Supplemental/Arial Unicode.ttf')) return
    const f = await makeFixture('Amount due 500')
    const out = await applyAll(f.bytes, [edit(f, 'Amount due 500', '应付金额 900 元')])
    expect(await extractText(out)).toBe('应付金额 900 元')
    // Subset embedding must not balloon the file (a full CJK font is ~20MB)
    expect(out.length).toBeLessThan(200 * 1024)
  })

  it('rebuilds with a requested edit font when its file exists', async () => {
    const { listEditFonts } = await import('../src/main/text-edit')
    const fontId = listEditFonts()[0]
    if (!fontId) return // machine has none of the curated font files
    const f = await makeFixture('Refont me please')
    // € exercises the cmap coverage check (it sits outside the Latin blocks proper)
    const out = await applyAll(f.bytes, [
      { ...edit(f, 'Refont me please', 'Refonted € text'), newFont: fontId },
    ])
    expect(await extractText(out)).toBe('Refonted € text')
    expect(out.length).toBeLessThan(200 * 1024)
    if (fontId === 'arial' && existsSync('/System/Library/Fonts/Supplemental/Arial.ttf')) {
      // BaseFont names in the output dict prove the chosen face was embedded, not the fallback
      const raw = Buffer.from(out).toString('latin1')
      expect(raw).toContain('ArialMT')
      expect(raw).not.toContain('ArialUnicodeMS')
    }
  })

  it('keeps the original embedded font on rebuild when its subset covers the text', async () => {
    if (!existsSync('/System/Library/Fonts/Supplemental/Arial.ttf')) return
    const f = await makeFixture('Amount due 500')
    const first = await applyAll(f.bytes, [
      { ...edit(f, 'Amount due 500', 'Hello there'), newFont: 'arial' },
    ])
    // newColor forces the rebuild path; every replacement char is already in the subset
    const out = await applyAll(first, [
      { ...edit(f, 'Hello there', 'there Hello'), newColor: [255, 0, 0] },
    ])
    expect(await extractText(out)).toBe('there Hello')
    const raw = Buffer.from(out).toString('latin1')
    expect(raw).toContain('ArialMT')
    expect(raw).not.toContain('ArialUnicodeMS')
  })

  it('resolves the same-named installed font when the subset lacks the new glyphs', async () => {
    if (!existsSync('/System/Library/Fonts/Supplemental/Arial.ttf')) return
    const f = await makeFixture('Amount due 500')
    const first = await applyAll(f.bytes, [
      { ...edit(f, 'Amount due 500', 'Hello there'), newFont: 'arial' },
    ])
    // z/q/k/j… are not in the 'Hello there' subset: the embedded face cannot cover, so the
    // rebuild must find system Arial by its PostScript name instead of the fallback face
    const out = await applyAll(first, [
      { ...edit(f, 'Hello there', 'Zebra quick jump'), newColor: [255, 0, 0] },
    ])
    expect(await extractText(out)).toBe('Zebra quick jump')
    const raw = Buffer.from(out).toString('latin1')
    expect(raw).toContain('ArialMT')
    expect(raw).not.toContain('ArialUnicodeMS')
  })

  it('keeps a CJK document font by resolving the installed face (未来→放心 scenario)', async () => {
    const { hasPingFang } = await import('./pingfang')
    if (!hasPingFang) return
    const { loadPdfium, saveDoc } = await import('../src/main/text-edit')
    const { findSystemFont } = await import('../src/main/font-locate')
    const { subsetTtf } = await import('../src/main/font-subset')
    // Fixture authored via pdfium itself: a PingFang subset holding only the original run,
    // like a real-world exported document (pdf-lib cannot embed custom fonts sans fontkit)
    const face = findSystemFont('PingFangSC-Regular', '')
    expect(face).not.toBeNull()
    const sub = await subsetTtf(face!, '可能在未来十年')
    const m = (await loadPdfium()) as Awaited<ReturnType<typeof loadPdfium>> & {
      _FPDF_CreateNewDocument(): number
      _FPDFPage_New(doc: number, index: number, w: number, h: number): number
    }
    const doc = m._FPDF_CreateNewDocument()
    const page = m._FPDFPage_New(doc, 0, 595, 842)
    const fontPtr = m._malloc(sub.length)
    m.HEAPU8.set(sub, fontPtr)
    // 2 = TRUETYPE, 1 = TYPE1 (CFF): PingFang is OTTO-flavored on recent macOS
    const font = m._FPDFText_LoadFont(
      doc,
      fontPtr,
      sub.length,
      sub.readUInt32BE(0) === 0x4f54544f ? 1 : 2,
      1,
    )
    m._free(fontPtr)
    const obj = m._FPDFPageObj_CreateTextObj(doc, font, 14)
    const text16 = Buffer.from('可能在未来十年\0', 'utf16le')
    const tp = m._malloc(text16.length)
    m.HEAPU8.set(text16, tp)
    expect(m._FPDFText_SetText(obj, tp)).toBeTruthy()
    m._free(tp)
    m._FPDFPageObj_Transform(obj, 1, 0, 0, 1, 50, 700)
    m._FPDFPage_InsertObject(page, obj)
    expect(m._FPDFPage_GenerateContent(page)).toBeTruthy()
    const fixture = saveDoc(m, doc)
    m._FPDF_ClosePage(page)
    m._FPDF_CloseDocument(doc)

    const out = await applyAll(fixture, [
      {
        pageIndex: 0,
        rect: [45, 692, 155, 722],
        oldText: '可能在未来十年',
        newText: '可能在放心十年',
        fontSize: 14,
      },
    ])
    expect(await extractText(out)).toBe('可能在放心十年')
    const raw = Buffer.from(out).toString('latin1')
    expect(raw).toContain('PingFang')
    expect(raw).not.toContain('ArialUnicodeMS')
    // The CFF-flavored face must sit in the FontFile3/OpenType slot: PDFium's TYPE1
    // path files it under /FontFile, which mainstream viewers render as blanks
    expect(raw).toContain('FontFile3')
    expect(raw).toContain('OpenType')
  })

  it('skips an edit whose replacement no available font can draw', async () => {
    if (!existsSync('/System/Library/Fonts/Supplemental/Arial Unicode.ttf')) return
    const f = await makeFixture('Amount due 500')
    // Emoji sit beyond every rebuild face incl. the fallback; embedding them would strand
    // missing glyphs and abort the save at read-back verification — skip the edit instead
    const result = await applyTextEdits(f.bytes, [edit(f, 'Amount due 500', 'Pay 🦄 now')])
    expect(result.skipped).toHaveLength(1)
    expect(result.skipped[0]!.reason).toContain('no available font')
    expect(result.bytes).toBe(f.bytes)
  })

  it('falls back to the default font for unknown or CJK-incompatible font requests', async () => {
    if (!existsSync('/System/Library/Fonts/Supplemental/Arial Unicode.ttf')) return
    const f = await makeFixture('Amount due 500')
    // Unknown id and a CJK replacement with a Latin-only face both take the fallback path
    const out = await applyAll(f.bytes, [
      { ...edit(f, 'Amount due 500', '应付金额 900 元'), newFont: 'arial' },
    ])
    expect(await extractText(out)).toBe('应付金额 900 元')
    const f2 = await makeFixture('Simple line')
    const out2 = await applyAll(f2.bytes, [
      { ...edit(f2, 'Simple line', 'Replaced line'), newFont: 'no-such-font' },
    ])
    expect(await extractText(out2)).toBe('Replaced line')
  })

  it('stacks newline-separated lines one leading apart', async () => {
    if (!existsSync('/System/Library/Fonts/Supplemental/Arial Unicode.ttf')) return
    const f = await makeFixture('Chapter heading')
    const out = await applyAll(f.bytes, [edit(f, 'Chapter heading', 'First line\nSecond line')])
    const { getDocument } = await import('pdfjs-dist/legacy/build/pdf.mjs')
    const doc = await getDocument({ data: out.slice(), useSystemFonts: true }).promise
    try {
      const content = await (await doc.getPage(1)).getTextContent()
      const items = content.items.filter(
        (i): i is (typeof content.items)[0] & { str: string; transform: number[] } =>
          'str' in i && i.str.trim() !== '',
      )
      expect(items.map((i) => i.str)).toEqual(['First line', 'Second line'])
      // Same left edge, second baseline one 1.2×size leading below the first
      expect(items[1]!.transform[4]).toBeCloseTo(items[0]!.transform[4]!, 1)
      expect(items[0]!.transform[5]! - items[1]!.transform[5]!).toBeCloseTo(14 * 1.2, 1)
    } finally {
      await doc.loadingTask.destroy()
    }
  })

  it('applies user font-size and color overrides on rebuild', async () => {
    if (!existsSync('/System/Library/Fonts/Supplemental/Arial Unicode.ttf')) return
    const f = await makeFixture('Recolor me')
    const out = await applyAll(f.bytes, [
      { ...edit(f, 'Recolor me', 'Recolored'), newFontSize: 22, newColor: [211, 47, 47] },
    ])
    const { OPS, getDocument } = await import('pdfjs-dist/legacy/build/pdf.mjs')
    const doc = await getDocument({ data: out.slice(), useSystemFonts: true }).promise
    try {
      const page = await doc.getPage(1)
      const content = await page.getTextContent()
      const item = content.items.find((i) => 'str' in i && i.str === 'Recolored')
      expect(item).toBeDefined()
      if (!item || !('transform' in item)) throw new Error('unreachable')
      expect(item.transform[3]).toBeCloseTo(22, 1)
      const ops = await page.getOperatorList()
      const fills = ops.fnArray
        .map((fn, i) => (fn === OPS.setFillRGBColor ? ops.argsArray[i] : null))
        .filter((a) => a !== null)
        .flatMap((a) => Array.from(a as ArrayLike<unknown>))
      expect(fills).toContain('#d32f2f')
    } finally {
      await doc.loadingTask.destroy()
    }
  })

  it('reinserts a rebuilt run at its original content-stream position', async () => {
    if (!existsSync('/System/Library/Fonts/Supplemental/Arial Unicode.ttf')) return
    // Two objects on one line drawn right-to-left: the leftmost (the rebuild anchor)
    // has the HIGHER object index, so a stale anchor index would misplace the insert
    const doc = await PDFDocument.create()
    const page = doc.addPage([595, 842])
    const font = await doc.embedFont(StandardFonts.Helvetica)
    page.drawText('right part', { x: 150, y: 700, size: 14, font })
    page.drawText('left part', { x: 50, y: 700, size: 14, font })
    page.drawText('Footer text', { x: 50, y: 100, size: 12, font })
    const bytes = await doc.save({ useObjectStreams: false })
    const out = await applyAll(bytes, [
      {
        pageIndex: 0,
        rect: [45, 694, 250, 718],
        oldText: 'right partleft part',
        newText: 'Replaced run',
        fontSize: 14,
      },
    ])
    const { getDocument } = await import('pdfjs-dist/legacy/build/pdf.mjs')
    const pdf = await getDocument({ data: out.slice(), useSystemFonts: true }).promise
    try {
      const content = await (await pdf.getPage(1)).getTextContent()
      const items = content.items
        .map((i) => ('str' in i ? i.str : ''))
        .filter((s) => s.trim() !== '')
      // Content-stream order: the rebuilt run must sit where the removed run began,
      // i.e. before the untouched footer that was drawn after it
      expect(items).toEqual(['Replaced run', 'Footer text'])
    } finally {
      await pdf.loadingTask.destroy()
    }
  })

  it('erases the run when the replacement is empty', async () => {
    const f = await makeFixture('Erase me entirely')
    const out = await applyAll(f.bytes, [edit(f, 'Erase me entirely', '')])
    expect(await extractText(out)).toBe('')
  })

  it('treats whitespace-only replacements as deletion too', async () => {
    const f = await makeFixture('Erase me entirely')
    const out = await applyAll(f.bytes, [edit(f, 'Erase me entirely', '\n \n')])
    expect(await extractText(out)).toBe('')
  })

  it('deletes a fragment while keeping the surrounding run text', async () => {
    const f = await makeFixture('Keep this remove that')
    const out = await applyAll(f.bytes, [edit(f, 'remove that', '')])
    const text = await extractText(out)
    expect(text).toContain('Keep this')
    expect(text).not.toContain('remove that')
  })

  it('skips edits whose text no longer matches and reports them', async () => {
    const f = await makeFixture('Signature line')
    const result = await applyTextEdits(f.bytes, [edit(f, 'Different text', 'X')])
    expect(result.skipped).toHaveLength(1)
    expect(result.skipped[0]!).toMatchObject({ pageIndex: 0, oldText: 'Different text' })
    expect(result.bytes).toBe(f.bytes)
  })

  it('applies valid edits even when another edit on the same page is skipped', async () => {
    const doc = await PDFDocument.create()
    const page = doc.addPage([595, 842])
    const font = await doc.embedFont(StandardFonts.Helvetica)
    page.drawText('Amount 1500', { x: 50, y: 700, size: 14, font })
    const bytes = await doc.save({ useObjectStreams: false })
    const w = font.widthOfTextAtSize('Amount 1500', 14)
    const good: TextEditInput = {
      pageIndex: 0,
      rect: [45, 694, 50 + w + 5, 718],
      oldText: 'Amount 1500',
      newText: 'Amount 5100',
      fontSize: 14,
    }
    const stale: TextEditInput = { ...good, rect: [45, 94, 250, 118], oldText: 'Gone text' }
    const result = await applyTextEdits(bytes, [good, stale])
    expect(result.skipped).toHaveLength(1)
    expect(result.skipped[0]!.oldText).toBe('Gone text')
    expect(await extractText(result.bytes)).toBe('Amount 5100')
  })

  it('resolves a fragment edit inside a larger text object (pdf.js span granularity)', async () => {
    // pdf.js can split one PDF text object into several spans; the edit's rect then
    // covers only part of the object. The fragment is spliced into the object's text.
    const f = await makeFixture('Total revenue was 1500 dollars in Q3')
    const doc = await PDFDocument.create()
    const font = await doc.embedFont(StandardFonts.Helvetica)
    const before = font.widthOfTextAtSize('Total revenue was ', 14)
    const fragW = font.widthOfTextAtSize('1500', 14)
    const out = await applyAll(f.bytes, [
      {
        pageIndex: 0,
        rect: [50 + before, 696, 50 + before + fragW, 716],
        oldText: '1500',
        newText: '1050',
        fontSize: 14,
      },
    ])
    expect(await extractText(out)).toBe('Total revenue was 1050 dollars in Q3')
  })

  it('reads text objects longer than the old fixed 4K buffer', async () => {
    // Regression: FPDFTextObj_GetText's length is in bytes and a short buffer is left
    // untouched — long runs used to come back as garbage and never match
    const long = `S${'o'.repeat(4200)} long`
    const f = await makeFixture(long)
    const out = await applyAll(f.bytes, [edit(f, long, 'Sooo long')])
    expect(await extractText(out)).toBe('Sooo long')
  })

  it('leaves untouched content intact across the rewrite', async () => {
    const doc = await PDFDocument.create()
    const page = doc.addPage([595, 842])
    const font = await doc.embedFont(StandardFonts.Helvetica)
    page.drawText('Keep me exactly', { x: 50, y: 760, size: 12, font })
    page.drawText('Change me now', { x: 50, y: 700, size: 14, font })
    const bytes = await doc.save({ useObjectStreams: false })
    const w = font.widthOfTextAtSize('Change me now', 14)
    const out = await applyAll(bytes, [
      {
        pageIndex: 0,
        rect: [45, 694, 50 + w + 5, 718],
        oldText: 'Change me now',
        newText: 'Changed now me',
        fontSize: 14,
      },
    ])
    const text = await extractText(out)
    expect(text).toContain('Keep me exactly')
    expect(text).toContain('Changed now me')
  })
})

describe('mergeEngineCodepoints', () => {
  it('keeps newText line breaks and seam spaces the engine text lacks', () => {
    // Engine text = two line objects joined, no space or break at the seam
    expect(
      mergeEngineCodepoints(
        'alpha betagamma delta',
        'alpha beta gamma delta',
        'alpha beta\ngamma delta',
      ),
    ).toBe('alpha beta\ngamma delta')
  })

  it('keeps newText structure on a style-only edit (identical text)', () => {
    expect(mergeEngineCodepoints('一二三四五六', '一二三四五六', '一二三\n四五六')).toBe(
      '一二三\n四五六',
    )
  })

  it('writes engine codepoints for unchanged units (extraction variants)', () => {
    // pdf.js extracts the Kangxi radical U+2F00 where the engine has U+4E00
    expect(mergeEngineCodepoints('一二三', '⼀二三', '⼀二三updated')).toBe('一二三updated')
  })

  it('keeps typed text in the changed middle', () => {
    expect(mergeEngineCodepoints('alphabetagamma', 'alpha beta gamma', 'alpha XX gamma')).toBe(
      'alpha XX gamma',
    )
  })

  it('backs off a boundary that would split an engine ligature', () => {
    // Engine has the ligature; the edit changes only the 'i' inside it — the cut
    // must retreat past the whole glyph instead of keeping an extra letter
    expect(mergeEngineCodepoints('ﬁne day', 'fine day', 'fone day')).toBe('fone day')
  })

  it('emits an engine ligature once across its folded units', () => {
    expect(mergeEngineCodepoints('ﬁrst second', 'first second', 'first\nsecond')).toBe(
      'ﬁrst\nsecond',
    )
  })
})

describe('block edits (paragraph rebuild)', () => {
  /** Three-line paragraph fixture at 18pt leading */
  async function makeParagraph(): Promise<Uint8Array> {
    const doc = await PDFDocument.create()
    const page = doc.addPage([595, 842])
    const font = await doc.embedFont(StandardFonts.Helvetica)
    const lines = ['first line of text', 'second line follows', 'third line ends here']
    for (const [i, l] of lines.entries())
      page.drawText(l, { x: 60, y: 700 - i * 18, size: 12, font })
    return doc.save({ useObjectStreams: false })
  }

  /** Per-item text + baseline y of the first page (pdf.js) */
  async function extractLines(bytes: Uint8Array): Promise<{ str: string; y: number }[]> {
    const { getDocument } = await import('pdfjs-dist/legacy/build/pdf.mjs')
    const doc = await getDocument({ data: bytes.slice(), useSystemFonts: true }).promise
    try {
      const page = await doc.getPage(1)
      const content = await page.getTextContent()
      return content.items
        .filter((i): i is { str: string; transform: number[] } => 'str' in i && i.str.length > 0)
        .map((i) => ({ str: i.str, y: Math.round(i.transform[5]!) }))
    } finally {
      await doc.loadingTask.destroy()
    }
  }

  it('keeps reflowed lines through a style-only edit (line structure survives the splice)', async () => {
    const bytes = await makeParagraph()
    const out = await applyAll(bytes, [
      {
        pageIndex: 0,
        rect: [55, 655, 320, 718],
        oldText: 'first line of text second line follows third line ends here',
        newText: 'first line of text\nsecond line follows\nthird line ends here',
        fontSize: 12,
        newColor: [200, 0, 0],
        origin: [60, 700],
        lineLeading: 18,
      },
    ])
    const lines = await extractLines(out)
    expect(lines.map((l) => l.str)).toEqual([
      'first line of text',
      'second line follows',
      'third line ends here',
    ])
    expect(lines.map((l) => l.y)).toEqual([700, 682, 664])
  })

  it('keeps seam spaces when the reflow moves the break positions', async () => {
    const bytes = await makeParagraph()
    const out = await applyAll(bytes, [
      {
        pageIndex: 0,
        rect: [55, 655, 320, 718],
        oldText: 'first line of text second line follows third line ends here',
        // Rewrapped to two lines: the old line seams now sit inside a line and
        // must come out as spaces, not glued words
        newText: 'first line of text second line follows\nthird line ends here',
        fontSize: 12,
        newColor: [0, 128, 0],
        origin: [60, 700],
        lineLeading: 18,
      },
    ])
    // The object-preserving rebuild keeps the original objects side by side on the
    // merged line; compare per-baseline text so extraction granularity doesn't matter
    const lines = await extractLines(out)
    const byY = new Map()
    for (const l of lines) byY.set(l.y, (byY.get(l.y) ?? '') + l.str)
    expect([...byY.values()].map((s) => s.replace(/\s+/g, ' ').trim())).toEqual([
      'first line of text second line follows',
      'third line ends here',
    ])
  })

  it('places lines at their per-line x offsets (centered reflow)', async () => {
    const bytes = await makeParagraph()
    const out = await applyAll(bytes, [
      {
        pageIndex: 0,
        rect: [55, 655, 320, 718],
        oldText: 'first line of text second line follows third line ends here',
        newText: 'centered one\ncentered two',
        fontSize: 12,
        newColor: [0, 0, 200],
        origin: [60, 700],
        lineLeading: 18,
        lineXOffsets: [40, 25],
      },
    ])
    const { getDocument } = await import('pdfjs-dist/legacy/build/pdf.mjs')
    const doc = await getDocument({ data: out.slice(), useSystemFonts: true }).promise
    const page = await doc.getPage(1)
    const content = await page.getTextContent()
    const items = content.items
      .filter((i): i is { str: string; transform: number[] } => 'str' in i && i.str.length > 0)
      .map((i) => ({ str: i.str, x: Math.round(i.transform[4]!) }))
    await doc.loadingTask.destroy()
    expect(items).toEqual([
      { str: 'centered one', x: 100 },
      { str: 'centered two', x: 85 },
    ])
  })
})

describe('selection color runs', () => {
  /** Text objects of page 1 straight from pdfium: text, fill color, matrix origin */
  async function inspectObjects(
    bytes: Uint8Array,
  ): Promise<{ text: string; color: [number, number, number]; x: number; y: number }[]> {
    const { chainPdfium, loadPdfium, withDocument } = await import('../src/main/text-edit')
    const m = await loadPdfium()
    return chainPdfium(() =>
      withDocument(m, bytes, async (doc) => {
        const page = m._FPDF_LoadPage(doc, 0)
        const textPage = m._FPDFText_LoadPage(page)
        const matPtr = m._malloc(24)
        const colPtr = m._malloc(16)
        const out: { text: string; color: [number, number, number]; x: number; y: number }[] = []
        try {
          const count = m._FPDFPage_CountObjects(page)
          for (let i = 0; i < count; i++) {
            const obj = m._FPDFPage_GetObject(page, i)
            if (m._FPDFPageObj_GetType(obj) !== 1) continue
            const len = m._FPDFTextObj_GetText(obj, textPage, 0, 0)
            let text = ''
            if (len > 2) {
              const buf = m._malloc(len)
              m._FPDFTextObj_GetText(obj, textPage, buf, len)
              text = Buffer.from(m.HEAPU8.buffer, buf, len - 2).toString('utf16le')
              m._free(buf)
            }
            m._FPDFPageObj_GetMatrix(obj, matPtr)
            const mat = Array.from(m.HEAPF32.subarray(matPtr >> 2, (matPtr >> 2) + 6))
            m._FPDFPageObj_GetFillColor(obj, colPtr, colPtr + 4, colPtr + 8, colPtr + 12)
            out.push({
              text,
              color: [m.HEAPU8[colPtr]!, m.HEAPU8[colPtr + 4]!, m.HEAPU8[colPtr + 8]!],
              x: mat[4]!,
              y: mat[5]!,
            })
          }
          return out
        } finally {
          m._free(matPtr)
          m._free(colPtr)
          m._FPDFText_ClosePage(textPage)
          m._FPDF_ClosePage(page)
        }
      }),
    )
  }

  it('plannedCharStyles aligns user runs onto engine-spliced text', async () => {
    const { plannedCharStyles } = await import('../src/main/text-edit')
    const red: [number, number, number] = [255, 0, 0]
    // Fragment case: the planned text wraps container text around the replacement;
    // legacy colorRuns input is accepted as color-only style runs
    const out = plannedCharStyles(
      { newText: 'THESE words', colorRuns: [{ start: 0, end: 5, color: red }] },
      'Highlight THESE words',
    )
    expect(out).not.toBeNull()
    // 'THESE' (10-14) styled; its trailing space inherits, the prefix space does not
    expect(out!.map((c) => (c ? 1 : 0)).join('')).toBe('000000000011111100000')
    expect(out![10]!.color).toEqual(red)
    // styleRuns carry face/size/flags through the same alignment
    const styled = plannedCharStyles(
      { newText: 'THESE words', styleRuns: [{ start: 0, end: 5, bold: true, size: 20 }] },
      'Highlight THESE words',
    )
    expect(styled).not.toBeNull()
    expect(styled![10]).toMatchObject({ bold: true, size: 20 })
    expect(styled![0]).toBeNull()
    // No runs → null (uniform rebuild)
    expect(plannedCharStyles({ newText: 'x', colorRuns: [] }, 'x')).toBeNull()
  })

  it('splits a recolored selection into per-color objects at measured offsets', async () => {
    if (!existsSync('/System/Library/Fonts/Supplemental/Arial Unicode.ttf')) return
    const f = await makeFixture('Highlight these words')
    const out = await applyAll(f.bytes, [
      {
        ...edit(f, 'Highlight these words', 'Highlight these words'),
        colorRuns: [{ start: 10, end: 15, color: [211, 47, 47] }],
      },
    ])
    expect(await extractText(out)).toBe('Highlight these words')
    const objs = (await inspectObjects(out)).filter((o) => o.text.trim() !== '')
    expect(objs.map((o) => o.text)).toEqual(['Highlight ', 'these ', 'words'])
    expect(objs.map((o) => o.color)).toEqual([
      [0, 0, 0],
      [211, 47, 47],
      [0, 0, 0],
    ])
    // Segments start where the preceding text ends (Helvetica-class metrics, ±15%)
    const doc = await PDFDocument.create()
    const font = await doc.embedFont(StandardFonts.Helvetica)
    const w1 = font.widthOfTextAtSize('Highlight ', 14)
    const w2 = font.widthOfTextAtSize('Highlight these ', 14)
    expect(objs[0]!.x).toBeCloseTo(50, 1)
    expect(objs[1]!.x).toBeGreaterThan(50 + w1 * 0.85)
    expect(objs[1]!.x).toBeLessThan(50 + w1 * 1.15)
    expect(objs[2]!.x).toBeGreaterThan(50 + w2 * 0.85)
    expect(objs[2]!.x).toBeLessThan(50 + w2 * 1.15)
    // All three segments share the baseline
    expect(objs[1]!.y).toBeCloseTo(objs[0]!.y, 1)
    expect(objs[2]!.y).toBeCloseTo(objs[0]!.y, 1)
  })

  it('applies color runs inside a fragment edit (pdf.js span granularity)', async () => {
    if (!existsSync('/System/Library/Fonts/Supplemental/Arial Unicode.ttf')) return
    const f = await makeFixture('Total revenue was 1500 dollars in Q3')
    const doc = await PDFDocument.create()
    const font = await doc.embedFont(StandardFonts.Helvetica)
    const before = font.widthOfTextAtSize('Total revenue was ', 14)
    const fragW = font.widthOfTextAtSize('1500', 14)
    const out = await applyAll(f.bytes, [
      {
        pageIndex: 0,
        rect: [50 + before, 696, 50 + before + fragW, 716],
        oldText: '1500',
        newText: '1500',
        fontSize: 14,
        colorRuns: [{ start: 0, end: 4, color: [26, 115, 232] }],
      },
    ])
    expect(await extractText(out)).toBe('Total revenue was 1500 dollars in Q3')
    const objs = (await inspectObjects(out)).filter((o) => o.text.trim() !== '')
    expect(objs.map((o) => o.text)).toEqual(['Total revenue was ', '1500 ', 'dollars in Q3'])
    expect(objs[1]!.color).toEqual([26, 115, 232])
    expect(objs[0]!.color).toEqual([0, 0, 0])
    expect(objs[2]!.color).toEqual([0, 0, 0])
  })

  it('carries colors across the lines of a paragraph rebuild', async () => {
    if (!existsSync('/System/Library/Fonts/Supplemental/Arial Unicode.ttf')) return
    const doc = await PDFDocument.create()
    const page = doc.addPage([595, 842])
    const font = await doc.embedFont(StandardFonts.Helvetica)
    const lines = ['first line of text', 'second line follows', 'third line ends here']
    for (const [i, l] of lines.entries())
      page.drawText(l, { x: 60, y: 700 - i * 18, size: 12, font })
    const bytes = await doc.save({ useObjectStreams: false })
    const newText = 'first line of text\nsecond line follows\nthird line ends here'
    const out = await applyAll(bytes, [
      {
        pageIndex: 0,
        rect: [55, 655, 320, 718],
        oldText: 'first line of text second line follows third line ends here',
        newText,
        fontSize: 12,
        origin: [60, 700],
        lineLeading: 18,
        // The whole middle line: a fully-covered line stays one object, just recolored
        colorRuns: [{ start: 19, end: 38, color: [211, 47, 47] }],
      },
    ])
    const objs = (await inspectObjects(out)).filter((o) => o.text.trim() !== '')
    expect(objs.map((o) => o.text)).toEqual(lines)
    expect(objs.map((o) => o.color)).toEqual([
      [0, 0, 0],
      [211, 47, 47],
      [0, 0, 0],
    ])
  })

  /** One-page PDF authored by pdfium itself with the given lines in a CJK-capable
      subset face (pdf-lib cannot embed custom fonts sans fontkit) */
  async function makeCjkFixture(lines: string[]): Promise<Uint8Array | null> {
    const { loadPdfium, saveDoc } = await import('../src/main/text-edit')
    const { findSystemFont } = await import('../src/main/font-locate')
    const { subsetTtf } = await import('../src/main/font-subset')
    const face =
      findSystemFont('PingFangSC-Regular', '') ??
      (existsSync('/System/Library/Fonts/Supplemental/Arial Unicode.ttf')
        ? await import('node:fs').then((fs) =>
            fs.readFileSync('/System/Library/Fonts/Supplemental/Arial Unicode.ttf'),
          )
        : null)
    if (!face) return null
    const sub = await subsetTtf(face, lines.join(''))
    const m = (await loadPdfium()) as Awaited<ReturnType<typeof loadPdfium>> & {
      _FPDF_CreateNewDocument(): number
      _FPDFPage_New(doc: number, index: number, w: number, h: number): number
    }
    const doc = m._FPDF_CreateNewDocument()
    const page = m._FPDFPage_New(doc, 0, 595, 842)
    const fontPtr = m._malloc(sub.length)
    m.HEAPU8.set(sub, fontPtr)
    const font = m._FPDFText_LoadFont(
      doc,
      fontPtr,
      sub.length,
      sub.readUInt32BE(0) === 0x4f54544f ? 1 : 2,
      1,
    )
    m._free(fontPtr)
    for (const [i, text] of lines.entries()) {
      const obj = m._FPDFPageObj_CreateTextObj(doc, font, 14)
      const textPtr = Buffer.from(`${text}\0`, 'utf16le')
      const tp = m._malloc(textPtr.length)
      m.HEAPU8.set(textPtr, tp)
      expect(m._FPDFText_SetText(obj, tp)).toBeTruthy()
      m._free(tp)
      m._FPDFPageObj_Transform(obj, 1, 0, 0, 1, 50, 700 - i * 17)
      m._FPDFPage_InsertObject(page, obj)
    }
    expect(m._FPDFPage_GenerateContent(page)).toBeTruthy()
    const bytes = saveDoc(m, doc)
    m._FPDF_ClosePage(page)
    m._FPDF_CloseDocument(doc)
    return bytes
  }

  it('recolors a CJK selection (advances come from the pdfium font, not the bytes)', async () => {
    const bytes = await makeCjkFixture(['合肥之戰约定'])
    if (!bytes) return
    const out = await applyAll(bytes, [
      {
        pageIndex: 0,
        rect: [45, 692, 145, 722],
        oldText: '合肥之戰约定',
        newText: '合肥之戰约定',
        fontSize: 14,
        colorRuns: [{ start: 0, end: 2, color: [0, 166, 80] }],
      },
    ])
    const objs = (await inspectObjects(out)).filter((o) => o.text.trim() !== '')
    expect(objs.map((o) => o.text)).toEqual(['合肥', '之戰约定'])
    expect(objs.map((o) => o.color)).toEqual([
      [0, 166, 80],
      [0, 0, 0],
    ])
    // The second segment starts two full-width glyphs (2 em) after the first
    expect(objs[1]!.x).toBeCloseTo(objs[0]!.x + 2 * 14, 0)
    expect(objs[1]!.y).toBeCloseTo(objs[0]!.y, 1)
  })

  it('carries a CJK color run across the lines of a block rebuild', async () => {
    const line1 = '壹二三,合肥之戰'
    const line2 = '约定'
    const bytes = await makeCjkFixture([line1, line2])
    if (!bytes) return
    const newText = `${line1}\n${line2}`
    const out = await applyAll(bytes, [
      {
        pageIndex: 0,
        rect: [45, 675, 170, 722],
        oldText: line1 + line2,
        newText,
        fontSize: 14,
        origin: [50, 700],
        lineLeading: 17,
        // the color run crosses the line break (run spans the '\n' at index 8)
        colorRuns: [{ start: 4, end: 11, color: [0, 166, 80] }],
        blockSource: line1 + line2,
      },
    ])
    const objs = (await inspectObjects(out)).filter((o) => o.text.trim() !== '')
    expect(objs.map((o) => o.text)).toEqual(['壹二三,', '合肥之戰', '约定'])
    expect(objs.map((o) => o.color)).toEqual([
      [0, 0, 0],
      [0, 166, 80],
      [0, 166, 80],
    ])
  })
})

describe('selection style runs', () => {
  /** Text objects of page 1 with their font size and base font name */
  async function inspectStyled(
    bytes: Uint8Array,
  ): Promise<{ text: string; size: number; fontName: string; x: number; y: number }[]> {
    const { chainPdfium, loadPdfium, withDocument } = await import('../src/main/text-edit')
    const m = await loadPdfium()
    return chainPdfium(() =>
      withDocument(m, bytes, async (doc) => {
        const page = m._FPDF_LoadPage(doc, 0)
        const textPage = m._FPDFText_LoadPage(page)
        const matPtr = m._malloc(24)
        const sizePtr = m._malloc(4)
        const out: { text: string; size: number; fontName: string; x: number; y: number }[] = []
        try {
          const count = m._FPDFPage_CountObjects(page)
          for (let i = 0; i < count; i++) {
            const obj = m._FPDFPage_GetObject(page, i)
            if (m._FPDFPageObj_GetType(obj) !== 1) continue
            const len = m._FPDFTextObj_GetText(obj, textPage, 0, 0)
            let text = ''
            if (len > 2) {
              const buf = m._malloc(len)
              m._FPDFTextObj_GetText(obj, textPage, buf, len)
              text = Buffer.from(m.HEAPU8.buffer, buf, len - 2).toString('utf16le')
              m._free(buf)
            }
            m._FPDFPageObj_GetMatrix(obj, matPtr)
            const mat = Array.from(m.HEAPF32.subarray(matPtr >> 2, (matPtr >> 2) + 6))
            m._FPDFTextObj_GetFontSize(obj, sizePtr)
            const font = m._FPDFTextObj_GetFont(obj)
            let fontName = ''
            const nameLen = m._FPDFFont_GetBaseFontName(font, 0, 0)
            if (nameLen > 1) {
              const buf = m._malloc(nameLen)
              m._FPDFFont_GetBaseFontName(font, buf, nameLen)
              fontName = Buffer.from(m.HEAPU8.subarray(buf, buf + nameLen - 1)).toString()
              m._free(buf)
            }
            out.push({ text, size: m.HEAPF32[sizePtr >> 2]!, fontName, x: mat[4]!, y: mat[5]! })
          }
          return out
        } finally {
          m._free(matPtr)
          m._free(sizePtr)
          m._FPDFText_ClosePage(textPage)
          m._FPDF_ClosePage(page)
        }
      }),
    )
  }

  it('draws a bolded selection with the style variant of the chosen face', async () => {
    if (!existsSync('/System/Library/Fonts/Supplemental/Arial Bold.ttf')) return
    const f = await makeFixture('Make this bold now')
    const out = await applyAll(f.bytes, [
      {
        ...edit(f, 'Make this bold now', 'Make this bold now'),
        newFont: 'arial',
        styleRuns: [{ start: 5, end: 9, bold: true }],
      },
    ])
    expect(await extractText(out)).toBe('Make this bold now')
    const objs = (await inspectStyled(out)).filter((o) => o.text.trim() !== '')
    expect(objs.map((o) => o.text)).toEqual(['Make ', 'this ', 'bold now'])
    // Same baseline, strictly advancing x
    expect(objs[1]!.x).toBeGreaterThan(objs[0]!.x)
    expect(objs[2]!.x).toBeGreaterThan(objs[1]!.x)
    expect(objs[1]!.y).toBe(objs[0]!.y)
    // The styled segment draws with the bold variant, the rest with the base face
    expect(objs[1]!.fontName).toMatch(/bold/i)
    expect(objs[0]!.fontName).not.toMatch(/bold/i)
    expect(objs[2]!.fontName).toBe(objs[0]!.fontName)
  })

  it('draws a resized selection at its own font size with advances to match', async () => {
    if (!existsSync('/System/Library/Fonts/Supplemental/Arial Unicode.ttf')) return
    const f = await makeFixture('Small BIG small')
    const out = await applyAll(f.bytes, [
      {
        ...edit(f, 'Small BIG small', 'Small BIG small'),
        styleRuns: [{ start: 6, end: 9, size: 22 }],
      },
    ])
    expect(await extractText(out)).toBe('Small BIG small')
    const objs = (await inspectStyled(out)).filter((o) => o.text.trim() !== '')
    expect(objs.map((o) => o.text)).toEqual(['Small ', 'BIG ', 'small'])
    expect(objs.map((o) => Math.round(o.size))).toEqual([14, 22, 14])
    expect(objs[2]!.x).toBeGreaterThan(objs[1]!.x)
  })

  it('combines a selection font, flags and color in one run', async () => {
    if (!existsSync('/System/Library/Fonts/Supplemental/Times New Roman Bold Italic.ttf')) return
    const f = await makeFixture('Style these words fully')
    const out = await applyAll(f.bytes, [
      {
        ...edit(f, 'Style these words fully', 'Style these words fully'),
        styleRuns: [
          { start: 6, end: 17, font: 'times', bold: true, italic: true, color: [211, 47, 47] },
        ],
      },
    ])
    expect(await extractText(out)).toBe('Style these words fully')
    const objs = (await inspectStyled(out)).filter((o) => o.text.trim() !== '')
    expect(objs.map((o) => o.text)).toEqual(['Style ', 'these words ', 'fully'])
    expect(objs[1]!.fontName).toMatch(/times/i)
    expect(objs[1]!.fontName).toMatch(/bold/i)
    expect(objs[1]!.fontName).toMatch(/italic/i)
  })

  it('restyles a stretch without disturbing the surrounding text', async () => {
    if (!existsSync('/System/Library/Fonts/Supplemental/Arial Bold.ttf')) return
    const f = await makeFixture('Keep head styled tail')
    const out = await applyAll(f.bytes, [
      {
        ...edit(f, 'Keep head styled tail', 'Keep head styled tail'),
        styleRuns: [{ start: 10, end: 16, font: 'arial', bold: true }],
      },
    ])
    expect(await extractText(out)).toBe('Keep head styled tail')
    const objs = (await inspectStyled(out)).filter((o) => o.text.trim() !== '')
    // 'styled' is redrawn bold; head and tail survive (as kept or rebuilt objects)
    const styled = objs.find((o) => o.text.trim() === 'styled')
    expect(styled).toBeDefined()
    expect(styled!.fontName).toMatch(/bold/i)
    expect(objs.map((o) => o.text).join('')).toBe('Keep head styled tail')
  })
})

describe('validateTextEdits', () => {
  it('returns a null reason for edits that would apply and a reason for those that would not', async () => {
    const f = await makeFixture('Validated content')
    const results = await validateTextEdits(f.bytes, [
      edit(f, 'Validated content', 'New content'),
      edit(f, 'Never was here', 'X'),
      { ...edit(f, 'Validated content', 'X'), pageIndex: 7 },
    ])
    expect(results[0]!.reason).toBeNull()
    expect(results[1]!.reason).toMatch(/could not be located/)
    expect(results[2]!.reason).toMatch(/page does not exist/)
    // Dry run: the document still applies the same edit afterwards
    const out = await applyAll(f.bytes, [edit(f, 'Validated content', 'New content')])
    expect(await extractText(out)).toBe('New content')
  })

  it('reports the matched run ink bounds so the preview can cover glyph overhang', async () => {
    const f = await makeFixture('Validated content')
    const probe = edit(f, 'Validated content', 'New content')
    const [v] = await validateTextEdits(f.bytes, [probe])
    expect(v!.reason).toBeNull()
    const b = v!.bounds!
    // Ink bounds land in the neighborhood of the edit rect and have positive area
    expect(b[2]).toBeGreaterThan(b[0])
    expect(b[3]).toBeGreaterThan(b[1])
    const [rx1, ry1, rx2, ry2] = probe.rect
    expect(b[0]).toBeGreaterThan(rx1 - 20)
    expect(b[2]).toBeLessThan(rx2 + 20)
    expect(b[1]).toBeGreaterThan(ry1 - 20)
    expect(b[3]).toBeLessThan(ry2 + 20)
    // No bounds for an edit that does not match
    const [miss] = await validateTextEdits(f.bytes, [edit(f, 'Never was here', 'X')])
    expect(miss!.bounds).toBeUndefined()
  })
})

describe('text-first rescue (multi-object lines)', () => {
  /** Lines whose every word is its own PDF text object on a shared baseline — the
      granularity Chrome/Word exports use. Word gaps are positional (no space chars
      inside the objects), so the joined engine text carries no spaces. */
  async function makeWordLines(
    lines: { words: string[]; y: number }[],
  ): Promise<{ bytes: Uint8Array; extents: { left: number; right: number }[] }> {
    const doc = await PDFDocument.create()
    const page = doc.addPage([595, 842])
    const font = await doc.embedFont(StandardFonts.Helvetica)
    const size = 14
    const extents: { left: number; right: number }[] = []
    for (const { words, y } of lines) {
      let x = 50
      let right = 50
      for (const w of words) {
        page.drawText(w, { x, y, size, font })
        right = x + font.widthOfTextAtSize(w, size)
        x = right + font.widthOfTextAtSize(' ', size)
      }
      extents.push({ left: 50, right })
    }
    return { bytes: await doc.save({ useObjectStreams: false }), extents }
  }

  const stripped = (s: string) => s.replace(/\s+/g, '')

  it('rescues a whole-line edit whose rect clips the edge objects', async () => {
    if (!existsSync('/System/Library/Fonts/Supplemental/Arial Unicode.ttf')) return
    const words = ['Use', 'it', 'to', 'cover', 'your', 'entire', 'phone', 'bill']
    const { bytes, extents } = await makeWordLines([{ words, y: 700 }])
    const { left, right } = extents[0]!
    // The rect starts inside 'Use' and ends inside 'bill': the primary path's
    // ≥50%-containment filter drops both edge objects and the line never matched
    const rect: [number, number, number, number] = [left + 18, 696, right - 14, 716]
    const out = await applyAll(bytes, [
      {
        pageIndex: 0,
        rect,
        oldText: words.join(' '),
        newText: 'Use it to pay your whole phone bill',
        fontSize: 14,
      },
    ])
    expect(stripped(await extractText(out))).toBe('Useittopayyourwholephonebill')
  })

  it('rescues a style-only recolor with a clipped rect (screenshot scenario)', async () => {
    if (!existsSync('/System/Library/Fonts/Supplemental/Arial Unicode.ttf')) return
    const { chainPdfium, loadPdfium, withDocument } = await import('../src/main/text-edit')
    const words = ['Use', 'it', 'to', 'cover', 'your', 'phone', 'bill']
    const { bytes, extents } = await makeWordLines([{ words, y: 700 }])
    const { left, right } = extents[0]!
    const oldText = words.join(' ')
    const out = await applyAll(bytes, [
      {
        pageIndex: 0,
        rect: [left + 18, 696, right - 12, 716],
        oldText,
        newText: oldText,
        fontSize: 14,
        newColor: [255, 0, 0],
      },
    ])
    expect(stripped(await extractText(out))).toBe(stripped(oldText))
    // Every text object on the page is repainted red
    const m = await loadPdfium()
    const colors = await chainPdfium(() =>
      withDocument(m, out, async (doc) => {
        const page = m._FPDF_LoadPage(doc, 0)
        const colPtr = m._malloc(16)
        const found: [number, number, number][] = []
        try {
          const count = m._FPDFPage_CountObjects(page)
          for (let i = 0; i < count; i++) {
            const obj = m._FPDFPage_GetObject(page, i)
            if (m._FPDFPageObj_GetType(obj) !== 1) continue
            m._FPDFPageObj_GetFillColor(obj, colPtr, colPtr + 4, colPtr + 8, colPtr + 12)
            found.push([m.HEAPU8[colPtr]!, m.HEAPU8[colPtr + 4]!, m.HEAPU8[colPtr + 8]!])
          }
          return found
        } finally {
          m._free(colPtr)
          m._FPDF_ClosePage(page)
        }
      }),
    )
    expect(colors.length).toBeGreaterThan(0)
    for (const c of colors) expect(c).toEqual([255, 0, 0])
  })

  it('picks the occurrence nearest the rect when the same text repeats', async () => {
    if (!existsSync('/System/Library/Fonts/Supplemental/Arial Unicode.ttf')) return
    const words = ['Repeated', 'footer', 'line']
    const { bytes, extents } = await makeWordLines([
      { words, y: 700 },
      { words, y: 680 },
    ])
    const { left, right } = extents[1]!
    // Clipped rect around the LOWER copy; the padded candidate set still touches the
    // upper copy, so the score must pick the right occurrence
    const out = await applyAll(bytes, [
      {
        pageIndex: 0,
        rect: [left + 20, 676, right - 8, 696],
        oldText: words.join(' '),
        newText: 'Repeated footer EDIT',
        fontSize: 14,
      },
    ])
    const { getDocument } = await import('pdfjs-dist/legacy/build/pdf.mjs')
    const doc = await getDocument({ data: out.slice(), useSystemFonts: true }).promise
    try {
      const content = await (await doc.getPage(1)).getTextContent()
      const byLine = new Map<number, string>()
      for (const i of content.items) {
        if (!('str' in i) || !i.str.trim()) continue
        const y = Math.round(i.transform[5]!)
        byLine.set(y, (byLine.get(y) ?? '') + i.str)
      }
      expect(stripped(byLine.get(700) ?? '')).toBe('Repeatedfooterline')
      expect(stripped(byLine.get(680) ?? '')).toBe('RepeatedfooterEDIT')
    } finally {
      await doc.loadingTask.destroy()
    }
  })

  it('rescues an edit ending mid-object, keeping the object tail verbatim', async () => {
    if (!existsSync('/System/Library/Fonts/Supplemental/Arial Unicode.ttf')) return
    // Three objects; the edit covers the first two and only 'Plus,' of the third —
    // like a text-layer line group stopping before the rest of the engine's run
    const doc = await PDFDocument.create()
    const page = doc.addPage([595, 842])
    const font = await doc.embedFont(StandardFonts.Helvetica)
    const parts = ['Use it to cover', 'your entire phone.', 'Plus, free shipping']
    let x = 50
    const starts: number[] = []
    for (const p of parts) {
      starts.push(x)
      page.drawText(p, { x, y: 700, size: 14, font })
      x += font.widthOfTextAtSize(`${p} `, 14)
    }
    const bytes = await doc.save({ useObjectStreams: false })
    const rectRight = starts[2]! + font.widthOfTextAtSize('Plus,', 14) + 2
    const out = await applyAll(bytes, [
      {
        pageIndex: 0,
        rect: [45, 696, rectRight, 716],
        oldText: 'Use it to cover your entire phone. Plus,',
        newText: 'Use it to cover your entire phone. Anyway,',
        fontSize: 14,
      },
    ])
    expect(stripped(await extractText(out))).toBe('Useittocoveryourentirephone.Anyway,freeshipping')
  })

  it('fails closed for a paragraph rescue that would end mid-object', async () => {
    // A block edit (origin/lineLeading set) may only rescue a WHOLE match: apply
    // strips the layout overrides from fragment matches, which would rebuild the
    // paragraph at the anchor and pull the edge object's tail into the block run
    const doc = await PDFDocument.create()
    const page = doc.addPage([595, 842])
    const font = await doc.embedFont(StandardFonts.Helvetica)
    const parts = ['Use it to cover', 'your entire phone.', 'Plus, free shipping']
    let x = 50
    const starts: number[] = []
    for (const p of parts) {
      starts.push(x)
      page.drawText(p, { x, y: 700, size: 14, font })
      x += font.widthOfTextAtSize(`${p} `, 14)
    }
    const bytes = await doc.save({ useObjectStreams: false })
    const rectRight = starts[2]! + font.widthOfTextAtSize('Plus,', 14) + 2
    const results = await validateTextEdits(bytes, [
      {
        pageIndex: 0,
        rect: [45, 696, rectRight, 716],
        oldText: 'Use it to cover your entire phone. Plus,',
        newText: 'Use it to cover\nyour entire phone. Anyway,',
        fontSize: 14,
        origin: [50, 700],
        lineLeading: 17,
      },
    ])
    expect(results[0]!.reason).toMatch(/could not be located/)
  })

  it('still reports unlocatable text instead of matching elsewhere', async () => {
    const { bytes, extents } = await makeWordLines([{ words: ['Only', 'this', 'line'], y: 700 }])
    const { left, right } = extents[0]!
    const results = await validateTextEdits(bytes, [
      {
        pageIndex: 0,
        rect: [left, 696, right, 716],
        oldText: 'Entirely different words',
        newText: 'X',
        fontSize: 14,
      },
    ])
    expect(results[0]!.reason).toMatch(/could not be located/)
  })
})

describe('whitespace edits (space deletion/insertion between runs)', () => {
  it('wsEditClamp flags space-adjacent units on both sides of the edit', async () => {
    const { wsEditClamp } = await import('../src/main/text-edit')
    // Deleting the space: the 'n' before it and the 'e' after it must be redrawn
    expect(wsEditClamp('phon e bill', 'phone bill')).toEqual({ headNS: 3, tailNS: 4 })
    // Inserting a space clamps the same way
    expect(wsEditClamp('phone bill', 'phon e bill')).toEqual({ headNS: 3, tailNS: 4 })
    // Glyph-only edits and identical texts don't clamp
    expect(wsEditClamp('abc def', 'abX def')).toBeNull()
    expect(wsEditClamp('same text ', 'same text ')).toBeNull()
    // Reflow: a '\n' replacing a space folds to the same unit — no clamp
    expect(wsEditClamp('alpha beta', 'alpha\nbeta')).toBeNull()
    // Pure insertion of a non-space next to a retained space edits no whitespace:
    // the old-side diff window is empty and the retained space must not clamp the
    // seam glyph into the redraw (fatal when it is an undrawable PUA icon glyph)
    expect(wsEditClamp('a b', 'a Xb')).toBeNull()
    expect(wsEditClamp('Note \uf105 Grey', 'Note\nX\uf105 Grey')).toBeNull()
    // Inserting an actual space still clamps (the new-side window carries it)
    expect(wsEditClamp('a Xb', 'a X b')).not.toBeNull()
  })

  it('keeps suffix objects when a leading space is deleted (headNS -1 stays safe)', async () => {
    if (!existsSync('/System/Library/Fonts/Supplemental/Arial Unicode.ttf')) return
    // Divergence at unit 0: pw = 0 makes headNS -1, which disables prefix keeping
    // outright. That is required (the unit right of the seam must not be kept via
    // the prefix, or the deleted space's gap would survive), and it must stay safe:
    // the tail clamp of the same seam still bounds the suffix at ns(suffix) - 1,
    // so keepable objects clear of the seam survive instead of a line-wide redraw.
    expect((await import('../src/main/text-edit')).wsEditClamp(' Use it to', 'Use it to')).toEqual({
      headNS: -1,
      tailNS: 6,
    })
    const doc = await PDFDocument.create()
    const page = doc.addPage([595, 842])
    const font = await doc.embedFont(StandardFonts.Helvetica)
    // First object carries the leading space as a real char; the rest are
    // word-per-object with positional gaps
    const parts = [' Use', 'it', 'to']
    let x = 50
    for (const p of parts) {
      page.drawText(p, { x, y: 700, size: 14, font })
      x += font.widthOfTextAtSize(`${p} `, 14)
    }
    const bytes = await doc.save({ useObjectStreams: false })
    const out = await applyAll(bytes, [
      {
        pageIndex: 0,
        rect: [40, 694, x + 5, 718],
        oldText: ' Use it to',
        newText: 'Use it to',
        fontSize: 14,
      },
    ])
    expect((await extractText(out)).replace(/\s+/g, '')).toBe('Useitto')
    // 'to' survives as its own original object via the suffix keep. ('it' is
    // unkeepable regardless of the clamp: pdfium's text page appends a generated
    // gap-space to its extracted text, failing the keep plan's chars↔codepoints
    // guard — the same edit with a glyph change instead of the space keeps
    // exactly the same set, so the clamp itself flattens nothing.)
    const { chainPdfium, loadPdfium, withDocument } = await import('../src/main/text-edit')
    const m = await loadPdfium()
    const texts = await chainPdfium(() =>
      withDocument(m, out, async (d) => {
        const page2 = m._FPDF_LoadPage(d, 0)
        const textPage = m._FPDFText_LoadPage(page2)
        const found: string[] = []
        try {
          const count = m._FPDFPage_CountObjects(page2)
          for (let i = 0; i < count; i++) {
            const obj = m._FPDFPage_GetObject(page2, i)
            if (m._FPDFPageObj_GetType(obj) !== 1) continue
            const len = m._FPDFTextObj_GetText(obj, textPage, 0, 0)
            if (len <= 2) continue
            const buf = m._malloc(len)
            m._FPDFTextObj_GetText(obj, textPage, buf, len)
            found.push(Buffer.from(m.HEAPU8.buffer, buf, len - 2).toString('utf16le'))
            m._free(buf)
          }
          return found
        } finally {
          m._FPDFText_ClosePage(textPage)
          m._FPDF_ClosePage(page2)
        }
      }),
    )
    expect(texts.map((t) => t.trim())).toContain('to')
    // No leading space anywhere: the ghost char went with the redrawn neighbor
    expect(texts.some((t) => t.startsWith(' '))).toBe(false)
  })

  it('removes a real space char split across two runs (ghost space)', async () => {
    if (!existsSync('/System/Library/Fonts/Supplemental/Arial Unicode.ttf')) return
    // Run 1 carries the space as a real char, run 2 sits past an extra visual gap —
    // the shape earlier edits leave behind (the tello invoice regression)
    const doc = await PDFDocument.create()
    const page = doc.addPage([595, 842])
    const font = await doc.embedFont(StandardFonts.Helvetica)
    page.drawText('Use phon ', { x: 50, y: 700, size: 14, font })
    const w1 = font.widthOfTextAtSize('Use phon ', 14)
    page.drawText('e bill', { x: 50 + w1 + 6, y: 700, size: 14, font })
    const bytes = await doc.save({ useObjectStreams: false })
    const out = await applyAll(bytes, [
      {
        pageIndex: 0,
        rect: [45, 694, 50 + w1 + 6 + font.widthOfTextAtSize('e bill', 14) + 5, 718],
        oldText: 'Use phon e bill',
        newText: 'Use phone bill',
        fontSize: 14,
      },
    ])
    // The deleted space must not survive as a char anywhere in the output
    expect(await extractText(out)).toBe('Use phone bill')
  })

  it('closes a positional gap when the synthesized space is deleted', async () => {
    if (!existsSync('/System/Library/Fonts/Supplemental/Arial Unicode.ttf')) return
    // Word-per-object line with positional gaps (no space chars): the renderer
    // synthesizes the spaces, so deleting one changes no space-free fold unit —
    // the keep plan used to keep every object at its original spacing (no-op save)
    const doc = await PDFDocument.create()
    const page = doc.addPage([595, 842])
    const font = await doc.embedFont(StandardFonts.Helvetica)
    const words = ['Use', 'it', 'to', 'cover', 'phon', 'e', 'bill']
    const xs: number[] = []
    let x = 50
    for (const w of words) {
      xs.push(x)
      page.drawText(w, { x, y: 700, size: 14, font })
      x += font.widthOfTextAtSize(`${w} `, 14)
    }
    const bytes = await doc.save({ useObjectStreams: false })
    const out = await applyAll(bytes, [
      {
        pageIndex: 0,
        rect: [45, 694, x + 5, 718],
        oldText: 'Use it to cover phon e bill',
        newText: 'Use it to cover phone bill',
        fontSize: 14,
      },
    ])
    expect((await extractText(out)).replace(/\s+/g, '')).toBe('Useittocoverphonebill')
    const { getDocument } = await import('pdfjs-dist/legacy/build/pdf.mjs')
    const pdf = await getDocument({ data: out.slice(), useSystemFonts: true }).promise
    try {
      const content = await (await pdf.getPage(1)).getTextContent()
      const items = content.items
        .filter((i): i is { str: string; transform: number[] } => 'str' in i && !!i.str.trim())
        .map((i) => ({ str: i.str, x: i.transform[4]! }))
      // 'phone' must be contiguous (one item, no gap where the deleted space sat)
      // and the untouched prefix stays anchored at the original line start. The
      // unchanged words are kept as original objects now, so pdf.js may split
      // items at the object seams — assert on the join, not one item's string.
      expect(items.map((i) => i.str).join(' ')).toContain('phone')
      const prefix = items.find((i) => i.str.includes('cover'))!
      expect(prefix.x).toBeCloseTo(xs[0]!, 1)
      // 'bill' moves left by about the deleted space width (rebuild-font advances
      // for the redrawn middle allow some tolerance, but the gap must be gone)
      const bill = items.find((i) => i.str.startsWith('bill'))!
      const spaceW = font.widthOfTextAtSize(' ', 14)
      expect(bill.x).toBeLessThan(xs[6]! - spaceW * 0.5)
      expect(bill.x).toBeGreaterThan(xs[6]! - spaceW * 2.5)
    } finally {
      await pdf.loadingTask.destroy()
    }
  })

  it('opens a gap when a space is typed between runs', async () => {
    if (!existsSync('/System/Library/Fonts/Supplemental/Arial Unicode.ttf')) return
    const doc = await PDFDocument.create()
    const page = doc.addPage([595, 842])
    const font = await doc.embedFont(StandardFonts.Helvetica)
    const words = ['Fix', 'ph', 'one', 'bill']
    const xs: number[] = []
    let x = 50
    for (const [i, w] of words.entries()) {
      xs.push(x)
      page.drawText(w, { x, y: 700, size: 14, font })
      // 'ph' and 'one' sit tight (one visual word); other words get a space gap
      x += font.widthOfTextAtSize(w, 14) + (i === 1 ? 0 : font.widthOfTextAtSize(' ', 14))
    }
    const bytes = await doc.save({ useObjectStreams: false })
    const out = await applyAll(bytes, [
      {
        pageIndex: 0,
        rect: [45, 694, x + 5, 718],
        oldText: 'Fix phone bill',
        newText: 'Fix ph one bill',
        fontSize: 14,
      },
    ])
    expect((await extractText(out)).replace(/\s+/g, ' ').trim()).toBe('Fix ph one bill')
    const { getDocument } = await import('pdfjs-dist/legacy/build/pdf.mjs')
    const pdf = await getDocument({ data: out.slice(), useSystemFonts: true }).promise
    try {
      const content = await (await pdf.getPage(1)).getTextContent()
      const items = content.items
        .filter((i): i is { str: string; transform: number[] } => 'str' in i && !!i.str.trim())
        .map((i) => ({ str: i.str, x: i.transform[4]! }))
      // 'bill' moves right by roughly the inserted space width
      const bill = items.find((i) => i.str.startsWith('bill'))!
      const spaceW = font.widthOfTextAtSize(' ', 14)
      expect(bill.x).toBeGreaterThan(xs[3]! + spaceW * 0.3)
      expect(bill.x).toBeLessThan(xs[3]! + spaceW * 2.5)
    } finally {
      await pdf.loadingTask.destroy()
    }
  })
})
