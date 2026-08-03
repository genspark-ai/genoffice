import { describe, it, expect } from 'vitest'
import { readFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { dirname, join } from 'node:path'
import JSZip from 'jszip'
import {
  openPptx,
  savePptx,
  patchSlideXml,
  patchTextElementXml,
  setElementFont,
  addTable,
  editTableCellText,
} from '../src/index'
import type { TextElement, TableElement } from '../src/types'

const here = dirname(fileURLToPath(import.meta.url))
const fx = (name: string) => readFileSync(join(here, 'fixtures', name))

function firstTextEl(els: any[]): TextElement {
  return els.find(
    (e) => (e.type === 'text' || e.type === 'shape') && e.text?.paragraphs?.length,
  ) as TextElement
}

describe('Phase 3.3 element-patch save', () => {
  it('no dirty → slide XML byte-identical (fidelity preserved)', async () => {
    const original = fx('01_standard_business.pptx')
    const opened = await openPptx(original)
    const out = await savePptx(opened)
    const origZip = await JSZip.loadAsync(original)
    const outZip = await JSZip.loadAsync(out)
    for (const slide of opened.deck.slides) {
      const a = await origZip.file(slide.path)!.async('uint8array')
      const b = await outZip.file(slide.path)!.async('uint8array')
      expect(Buffer.from(b).equals(Buffer.from(a))).toBe(true)
    }
  })

  it('patchSlideXml with no dirty == originalXml', async () => {
    const { deck } = await openPptx(fx('01_standard_business.pptx'))
    for (const slide of deck.slides) {
      expect(patchSlideXml(slide)).toBe(slide.originalXml)
    }
  })

  it('editing text content: new text lands, other bytes intact', async () => {
    const { deck } = await openPptx(fx('01_standard_business.pptx'))
    const slide = deck.slides.find((s) => firstTextEl(s.elements))!
    const el = firstTextEl(slide.elements)
    const _oldText = el
      .text!.paragraphs.flatMap((p) => p.runs)
      .map((r) => r.text)
      .join('')
    // Change the first run's text
    el.text!.paragraphs[0]!.runs[0]!.text = 'PATCHED_TITLE_TEXT'
    el.dirty = true

    const rebuilt = patchSlideXml(slide)
    expect(rebuilt).toContain('PATCHED_TITLE_TEXT')
    // bodyPrefix (spTree open + nvGrpSpPr) kept as-is
    expect(rebuilt.startsWith(slide.bodyPrefix)).toBe(true)
    expect(rebuilt.endsWith(slide.bodySuffix)).toBe(true)
    // Original bytes of other unchanged elements still present
    const otherEl = slide.elements.find((e) => e !== el && e.anchor.originalXml.length > 20)
    if (otherEl) expect(rebuilt).toContain(otherEl.anchor.originalXml)
  })

  it('escapes XML special chars in edited text', () => {
    const el = {
      type: 'text',
      text: { paragraphs: [{ runs: [{ text: '<b>&"x"</b>' }] }] },
    } as unknown as TextElement
    const original =
      '<p:sp><p:txBody><a:p><a:r><a:rPr/><a:t>old</a:t></a:r></a:p></p:txBody></p:sp>'
    const out = patchTextElementXml(el, original)
    expect(out).toContain('&lt;b&gt;&amp;"x"&lt;/b&gt;')
    expect(out).not.toContain('<b>&"x"</b>')
  })

  it('text patch keeps bodyPr vert bytes (vertical-text edit round-trip lossless)', () => {
    const bodyPr = '<a:bodyPr vert="eaVert" wrap="none"/>'
    const original = `<p:sp><p:txBody>${bodyPr}<a:p><a:r><a:rPr/><a:t>旧</a:t></a:r></a:p></p:txBody></p:sp>`
    // In-place patch path (paragraph count unchanged)
    const el1 = {
      type: 'text',
      text: { paragraphs: [{ runs: [{ text: '新しい縦書き' }] }] },
    } as unknown as TextElement
    const out1 = patchTextElementXml(el1, original)
    expect(out1).toContain(bodyPr)
    expect(out1).toContain('新しい縦書き')
    // Structural change → rebuild-txBody path (original bodyPr bytes still kept)
    const el2 = {
      type: 'text',
      text: { paragraphs: [{ runs: [{ text: '一段目' }] }, { runs: [{ text: '二段目' }] }] },
    } as unknown as TextElement
    const out2 = patchTextElementXml(el2, original)
    expect(out2).toContain(bodyPr)
    expect(out2).toContain('二段目')
  })

  it('formatting patch: bold + size + color on existing rPr', () => {
    const el = {
      type: 'text',
      text: {
        paragraphs: [{ runs: [{ text: 'hi', bold: true, fontSize: 24, color: '#FF0000' }] }],
      },
    } as unknown as TextElement
    const original =
      '<p:sp><p:txBody><a:p><a:r><a:rPr lang="en-US" sz="1800"/><a:t>hi</a:t></a:r></a:p></p:txBody></p:sp>'
    const out = patchTextElementXml(el, original)
    expect(out).toContain('sz="2400"') // 24pt
    expect(out).toMatch(/b="1"/)
    expect(out).toContain('srgbClr val="FF0000"')
    expect(out).toContain('<a:t>hi</a:t>')
    // Original lang attribute kept
    expect(out).toContain('lang="en-US"')
  })

  it('injects rPr when run had none', () => {
    const el = {
      type: 'text',
      text: { paragraphs: [{ runs: [{ text: 'x', bold: true, color: '#00FF00' }] }] },
    } as unknown as TextElement
    const original = '<p:sp><p:txBody><a:p><a:r><a:t>x</a:t></a:r></a:p></p:txBody></p:sp>'
    const out = patchTextElementXml(el, original)
    expect(out).toContain('<a:rPr')
    expect(out).toMatch(/b="1"/)
    expect(out).toContain('srgbClr val="00FF00"')
  })

  it('structural rebuild path when run count changes', () => {
    const el = {
      type: 'text',
      text: {
        paragraphs: [{ runs: [{ text: 'A' }, { text: 'B' }] }], // 2 runs vs 1 original
      },
    } as unknown as TextElement
    const original =
      '<p:sp><p:spPr/><p:txBody><a:bodyPr wrap="square"/><a:lstStyle/><a:p><a:r><a:rPr/><a:t>orig</a:t></a:r></a:p></p:txBody></p:sp>'
    const out = patchTextElementXml(el, original)
    // bodyPr kept as-is
    expect(out).toContain('<a:bodyPr wrap="square"/>')
    // Both new run texts present
    expect(out).toContain('<a:t>A</a:t>')
    expect(out).toContain('<a:t>B</a:t>')
    // Old text has been replaced
    expect(out).not.toContain('orig')
  })

  it('full save roundtrip after edit: file opens again with new text', async () => {
    const { deck } = await openPptx(fx('05_unicode_cjk_emoji.pptx'))
    const slide = deck.slides.find((s) => firstTextEl(s.elements))!
    const el = firstTextEl(slide.elements)
    el.text!.paragraphs[0]!.runs[0]!.text = '编辑后的标题ABC'
    el.dirty = true
    const opened = { deck, archive: (await openPptx(fx('05_unicode_cjk_emoji.pptx'))).archive }
    // Save with the same deck (archive only supplies the other entries' bytes)
    const out = await savePptx({ deck, archive: opened.archive })
    // Reopen the output and confirm the new text parses back
    const reopened = await openPptx(out)
    const text = reopened.deck.slides
      .flatMap((s) => s.elements)
      .filter((e: any) => e.type === 'text' || e.type === 'shape')
      .flatMap((e: any) => e.text?.paragraphs ?? [])
      .flatMap((p: any) => p.runs)
      .map((r: any) => r.text)
      .join('')
    expect(text).toContain('编辑后的标题ABC')
  })
})

describe('setElementFont (change font family/size for a whole selected element)', () => {
  it('text/shape: all runs get family+size, dirty set, XML regenerates', async () => {
    const { deck } = await openPptx(fx('01_standard_business.pptx'))
    const slide = deck.slides.find((s) => firstTextEl(s.elements))!
    const el = firstTextEl(slide.elements)
    expect(setElementFont(slide, el.id, { fontFamily: 'Noto Sans SC', fontSizePt: 20 })).toBe(true)
    for (const p of el.text!.paragraphs)
      for (const r of p.runs) {
        expect(r.fontFamily).toBe('Noto Sans SC')
        expect(r.fontSize).toBe(20)
      }
    expect(el.dirty).toBe(true)
    const rebuilt = patchSlideXml(slide)
    expect(rebuilt).toContain('typeface="Noto Sans SC"')
    expect(rebuilt).toContain('sz="2000"')
  })

  it('only family: size untouched; unknown/no-text element returns false', async () => {
    const { deck } = await openPptx(fx('01_standard_business.pptx'))
    const slide = deck.slides.find((s) => firstTextEl(s.elements))!
    const el = firstTextEl(slide.elements)
    const sizeBefore = el.text!.paragraphs[0]!.runs[0]!.fontSize
    expect(setElementFont(slide, el.id, { fontFamily: 'Arial' })).toBe(true)
    expect(el.text!.paragraphs[0]!.runs[0]!.fontSize).toBe(sizeBefore)
    expect(setElementFont(slide, 'no-such-id', { fontFamily: 'Arial' })).toBe(false)
    const pic = slide.elements.find((e) => e.type === 'picture')
    if (pic) expect(setElementFont(slide, pic.id, { fontFamily: 'Arial' })).toBe(false)
  })

  it('table: sizes cells with text and leaves a persisted marker on empty cells', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const r = addTable(opened, 0, {
      rows: 2,
      cols: 2,
      offset: { x: 0, y: 0, cx: 1905000, cy: 952500 },
    })!
    editTableCellText(r.slide, r.elementId, 0, 0, [{ runs: [{ text: 'Hello' }] }])
    expect(setElementFont(r.slide, r.elementId, { fontSizePt: 14 })).toBe(true)

    const tbl = r.slide.elements.find((e) => e.id === r.elementId) as TableElement
    expect(tbl.rows[0]![0]!.text!.paragraphs[0]!.runs[0]!.fontSize).toBe(14)
    // Empty cell got a textless marker run carrying the size
    const marker = tbl.rows[0]![1]!.text!.paragraphs[0]!.runs[0]!
    expect(marker.text).toBe('')
    expect(marker.fontSize).toBe(14)
    // sz persisted for every cell (4 cells; the texty cell's run + 3 markers)
    expect(tbl.anchor.originalXml.match(/sz="1400"/g)?.length).toBe(4)

    // Marker survives save/reopen
    const reopened = await openPptx(await savePptx(opened))
    const tbl2 = reopened.deck.slides[0]!.elements.find((e) => e.type === 'table') as TableElement
    expect(tbl2.rows[0]![1]!.text!.paragraphs[0]!.runs[0]!.fontSize).toBe(14)
  })

  it('table with only empty cells still reports a change', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const r = addTable(opened, 0, {
      rows: 2,
      cols: 2,
      offset: { x: 0, y: 0, cx: 1905000, cy: 952500 },
    })!
    expect(setElementFont(r.slide, r.elementId, { fontSizePt: 14 })).toBe(true)
  })
})
