import { existsSync } from 'node:fs'
import { PDFDocument, StandardFonts } from 'pdf-lib'
import { describe, expect, it } from 'vitest'
import { applyTextEdits, mergeEngineCodepoints, validateTextEdits } from '../src/main/text-edit'
import type { TextEditInput } from '../src/shared/ipc'

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

describe('applyTextEdits', () => {
  it('replaces text in place when the font already covers it', async () => {
    const f = await makeFixture('Total revenue was 1500 dollars')
    const out = await applyAll(f.bytes, [
      edit(f, 'Total revenue was 1500 dollars', 'Total revenue was 5100 dollars'),
    ])
    expect(await extractText(out)).toBe('Total revenue was 5100 dollars')
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

  it('skips blank replacements instead of erasing the run', async () => {
    const f = await makeFixture('Do not erase me')
    const result = await applyTextEdits(f.bytes, [edit(f, 'Do not erase me', '\n \n')])
    expect(result.skipped).toHaveLength(1)
    expect(result.skipped[0]!.reason).toMatch(/empty replacement/)
    // Nothing applied → the input bytes come back untouched
    expect(result.bytes).toBe(f.bytes)
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
    const lines = await extractLines(out)
    expect(lines.map((l) => l.str)).toEqual([
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

  it('plannedCharColors aligns user runs onto engine-spliced text', async () => {
    const { plannedCharColors } = await import('../src/main/text-edit')
    const red: [number, number, number] = [255, 0, 0]
    // Fragment case: the planned text wraps container text around the replacement
    const out = plannedCharColors(
      { newText: 'THESE words', colorRuns: [{ start: 0, end: 5, color: red }] },
      'Highlight THESE words',
    )
    expect(out).not.toBeNull()
    // 'THESE' (10-14) colored; its trailing space inherits, the prefix space does not
    expect(out!.map((c) => (c ? 1 : 0)).join('')).toBe('000000000011111100000')
    // No runs → null (uniform rebuild)
    expect(plannedCharColors({ newText: 'x', colorRuns: [] }, 'x')).toBeNull()
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
