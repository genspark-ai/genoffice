import { describe, it, expect } from 'vitest'
import { readFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { dirname, join } from 'node:path'
import {
  addPicture,
  addElement,
  addTable,
  copyElementData,
  deleteSlide,
  editTableCellText,
  editTableStructure,
  setTableColWidth,
  elementSpid,
  getSlideAnimations,
  getSlideTransition,
  openPptx,
  pasteElements,
  reorderElement,
  savePptx,
  setSlideAnimations,
  setSlideTransition,
} from '../src/index'
import type { PictureElement, TableElement, TextElement } from '../src/types'

const here = dirname(fileURLToPath(import.meta.url))
const fx = (name: string) => readFileSync(join(here, 'fixtures', name))

const OFF = { x: 914400, y: 914400, cx: 1828800, cy: 914400 }
const SHIFT = { dx: 152400, dy: 152400 }

// 1x1 red PNG
const PNG_1PX = Uint8Array.from(
  Buffer.from(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==',
    'base64',
  ),
)

describe('copy/paste elements', () => {
  it('pastes a textbox onto the same slide with shifted offset', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const slide = opened.deck.slides[0]!
    const el = addElement(slide, {
      kind: 'textbox',
      offset: { ...OFF },
      paragraphs: [{ runs: [{ text: 'copy me', bold: true }] }],
    })
    const before = slide.elements.length
    const clip = copyElementData(opened, slide, el)

    const r = pasteElements(opened, 0, [clip], SHIFT)
    expect(r).not.toBeNull()
    expect(r!.slide.elements.length).toBe(before + 1)
    expect(r!.elementIds.length).toBe(1)

    const pasted = r!.slide.elements.at(-1) as TextElement
    expect(pasted.type).toBe('text')
    expect(pasted.text!.paragraphs[0]!.runs[0]!.text).toBe('copy me')
    expect(pasted.text!.paragraphs[0]!.runs[0]!.bold).toBe(true)
    expect(pasted.transform.offset.x).toBe(OFF.x + SHIFT.dx)
    expect(pasted.transform.offset.y).toBe(OFF.y + SHIFT.dy)

    // cNvPr id does not clash with the original element
    const ids = r!.slide.elements.flatMap((e) =>
      [...e.anchor.originalXml.matchAll(/<p:cNvPr\s[^>]*\bid="(\d+)"/g)].map((m) => m[1]),
    )
    expect(new Set(ids).size).toBe(ids.length)

    // Still present after save → reopen
    const reopened = await openPptx(await savePptx(opened))
    expect(reopened.deck.slides[0]!.elements.length).toBe(before + 1)
  })

  it('pastes a picture onto another slide (rels surgery, media shared)', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    expect(opened.deck.slides.length).toBeGreaterThan(1)
    const slide0 = opened.deck.slides[0]!
    const pic = addPicture(opened, slide0, { bytes: PNG_1PX, ext: 'png', offset: { ...OFF } })!
    const clip = copyElementData(opened, slide0, pic)
    expect(clip.rels.length).toBe(1)
    expect(clip.rels[0]!.target).toBe(pic.mediaRef)

    const before1 = opened.deck.slides[1]!.elements.length
    const r = pasteElements(opened, 1, [clip], SHIFT)
    expect(r).not.toBeNull()
    const pasted = r!.slide.elements.at(-1) as PictureElement
    expect(pasted.type).toBe('picture')
    expect(pasted.mediaRef).toBe(pic.mediaRef)

    const reopened = await openPptx(await savePptx(opened))
    const el2 = reopened.deck.slides[1]!.elements.at(-1) as PictureElement
    expect(reopened.deck.slides[1]!.elements.length).toBe(before1 + 1)
    expect(el2.type).toBe('picture')
    expect(el2.mediaRef).toBe(pic.mediaRef)
    // media part not copied
    const mediaParts = [...reopened.archive.entries.keys()].filter((p) =>
      p.startsWith('ppt/media/'),
    )
    expect(mediaParts.filter((p) => p === pic.mediaRef).length).toBe(1)
  })

  it('pending text edits are baked into the copied xml', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const slide = opened.deck.slides[0]!
    const el = addElement(slide, {
      kind: 'textbox',
      offset: { ...OFF },
      paragraphs: [{ runs: [{ text: 'original text' }] }],
    })
    el.text!.paragraphs = [{ runs: [{ text: 'edited text' }] }]
    el.dirty = true

    const clip = copyElementData(opened, slide, el)
    expect(clip.xml).toContain('edited text')
  })
})

describe('addTable', () => {
  it('inserts a styled table that parses and round-trips', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const before = opened.deck.slides[0]!.elements.length
    const r = addTable(opened, 0, {
      rows: 3,
      cols: 4,
      offset: { x: 914400, y: 914400, cx: 4572000, cy: 1828800 },
    })
    expect(r).not.toBeNull()
    expect(r!.slide.elements.length).toBe(before + 1)

    const tbl = r!.slide.elements.at(-1) as TableElement
    expect(tbl.type).toBe('table')
    expect(tbl.rows.length).toBe(3)
    expect(tbl.colWidths.length).toBe(4)
    expect(tbl.rows[0]!.length).toBe(4)
    // Header row gets the fill from the built-in Medium Style 2
    expect(tbl.rows[0]![0]!.fill).toBeDefined()

    const reopened = await openPptx(await savePptx(opened))
    const tbl2 = reopened.deck.slides[0]!.elements.at(-1) as TableElement
    expect(tbl2.type).toBe('table')
    expect(tbl2.rows.length).toBe(3)
    expect(tbl2.colWidths.length).toBe(4)
  })
})

describe('editTableCellText', () => {
  it('writes cell paragraphs, keeps style, survives save/reopen', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const r = addTable(opened, 0, {
      rows: 2,
      cols: 2,
      offset: { x: 914400, y: 914400, cx: 3657600, cy: 1828800 },
    })!
    const slide = r.slide
    const tblId = r.elementId

    expect(
      editTableCellText(slide, tblId, 0, 1, [{ runs: [{ text: 'Header B', bold: true }] }]),
    ).toBe(true)
    expect(editTableCellText(slide, tblId, 1, 0, [{ runs: [{ text: 'data' }] }])).toBe(true)
    // Out of range rejected
    expect(editTableCellText(slide, tblId, 5, 0, [{ runs: [{ text: 'x' }] }])).toBe(false)

    const tbl = slide.elements.find((e) => e.id === tblId) as TableElement
    expect(tbl.rows[0]![1]!.text!.paragraphs[0]!.runs[0]!.text).toBe('Header B')

    const reopened = await openPptx(await savePptx(opened))
    const tbl2 = reopened.deck.slides[0]!.elements.at(-1) as TableElement
    expect(tbl2.rows[0]![1]!.text!.paragraphs[0]!.runs[0]!.text).toBe('Header B')
    expect(tbl2.rows[0]![1]!.text!.paragraphs[0]!.runs[0]!.bold).toBe(true)
    expect(tbl2.rows[1]![0]!.text!.paragraphs[0]!.runs[0]!.text).toBe('data')
    // Table style reference not broken
    expect(tbl2.rows[0]![0]!.fill).toBeDefined()
  })
})

describe('table structure ops', () => {
  const TBL_OFF = { x: 914400, y: 914400, cx: 3657600, cy: 1828800 }

  it('inserts and deletes rows/cols, frame ext follows, round-trips', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const added = addTable(opened, 0, { rows: 2, cols: 2, offset: { ...TBL_OFF } })!
    const slide = added.slide
    let elementId = added.elementId
    editTableCellText(slide, elementId, 0, 0, [{ runs: [{ text: 'A1' }] }])

    // Insert row below → 3 rows, frame gets taller
    let r = editTableStructure(opened, 0, elementId, { kind: 'insert-row', index: 0 })!
    expect(r).not.toBeNull()
    let tbl = r.slide.elements.find((e) => e.id === r!.elementId) as TableElement
    expect(tbl.rows.length).toBe(3)
    expect(tbl.transform.offset.cy).toBeGreaterThan(TBL_OFF.cy)
    // New row text is empty; original text remains
    expect(tbl.rows[0]![0]!.text!.paragraphs[0]!.runs[0]!.text).toBe('A1')
    elementId = r.elementId

    // Insert column on the left → 3 columns, frame gets wider
    r = editTableStructure(opened, 0, elementId, { kind: 'insert-col', index: 1, before: true })!
    tbl = r.slide.elements.find((e) => e.id === r!.elementId) as TableElement
    expect(tbl.colWidths.length).toBe(3)
    expect(tbl.transform.offset.cx).toBeGreaterThan(TBL_OFF.cx)
    elementId = r.elementId

    // Delete row/column back to 2×2
    r = editTableStructure(opened, 0, elementId, { kind: 'delete-row', index: 2 })!
    elementId = r.elementId
    r = editTableStructure(opened, 0, elementId, { kind: 'delete-col', index: 1 })!
    tbl = r.slide.elements.find((e) => e.id === r!.elementId) as TableElement
    expect(tbl.rows.length).toBe(2)
    expect(tbl.colWidths.length).toBe(2)

    const reopened = await openPptx(await savePptx(opened))
    const tbl2 = reopened.deck.slides[0]!.elements.at(-1) as TableElement
    expect(tbl2.rows.length).toBe(2)
    expect(tbl2.colWidths.length).toBe(2)
    expect(tbl2.rows[0]![0]!.text!.paragraphs[0]!.runs[0]!.text).toBe('A1')
  })

  it('refuses last row/col and merged tables', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const { slide, elementId } = addTable(opened, 0, { rows: 1, cols: 1, offset: { ...TBL_OFF } })!
    expect(editTableStructure(opened, 0, elementId, { kind: 'delete-row', index: 0 })).toBeNull()
    expect(editTableStructure(opened, 0, elementId, { kind: 'delete-col', index: 0 })).toBeNull()

    const el = slide.elements.find((e) => e.id === elementId)!
    el.anchor.originalXml = el.anchor.originalXml.replace('<a:tc>', '<a:tc gridSpan="1">')
    expect(editTableStructure(opened, 0, elementId, { kind: 'insert-row', index: 0 })).toBeNull()
  })

  it('setTableColWidth updates gridCol, model and frame width in place', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const { slide, elementId } = addTable(opened, 0, { rows: 2, cols: 2, offset: { ...TBL_OFF } })!
    const tbl = slide.elements.find((e) => e.id === elementId) as TableElement
    const w0 = tbl.colWidths[0]!

    expect(setTableColWidth(slide, elementId, 0, w0 * 2)).toBe(true)
    expect(tbl.colWidths[0]).toBe(w0 * 2)
    expect(tbl.transform.offset.cx).toBe(w0 * 2 + tbl.colWidths[1]!)
    expect(setTableColWidth(slide, elementId, 5, 100)).toBe(false)

    const reopened = await openPptx(await savePptx(opened))
    const tbl2 = reopened.deck.slides[0]!.elements.at(-1) as TableElement
    expect(tbl2.colWidths[0]).toBe(w0 * 2)
    expect(tbl2.transform.offset.cx).toBe(w0 * 2 + tbl2.colWidths[1]!)
  })

  it('rtl table: growing a column shifts the origin left (visual-right edge anchored)', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const { slide, elementId } = addTable(opened, 0, { rows: 2, cols: 2, offset: { ...TBL_OFF } })!
    const tbl = slide.elements.find((e) => e.id === elementId) as TableElement
    tbl.anchor.originalXml = tbl.anchor.originalXml.replace('<a:tblPr', '<a:tblPr rtl="1"')
    tbl.rtl = true
    const w0 = tbl.colWidths[0]!
    const x0 = tbl.transform.offset.x

    expect(setTableColWidth(slide, elementId, 0, w0 + 91440)).toBe(true)
    expect(tbl.transform.offset.x).toBe(x0 - 91440)
    expect(tbl.transform.offset.cx).toBe(w0 + 91440 + tbl.colWidths[1]!)
    expect(tbl.anchor.originalXml).toContain(`x="${x0 - 91440}"`)
  })
})

describe('deleteSlide', () => {
  it('removes the slide from deck and package, round-trips', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const total = opened.deck.slides.length
    expect(total).toBeGreaterThan(1)
    const victim = opened.deck.slides[1]!.path

    expect(deleteSlide(opened, 1)).toBe(true)
    expect(opened.deck.slides.length).toBe(total - 1)
    expect(opened.archive.entries.has(victim)).toBe(false)

    const reopened = await openPptx(await savePptx(opened))
    expect(reopened.deck.slides.length).toBe(total - 1)
    expect(reopened.deck.slides.some((s) => s.path === victim)).toBe(false)
    // No longer referenced by presentation.xml or [Content_Types]
    expect(reopened.archive.readText('[Content_Types].xml')).not.toContain(`/${victim}`)
  })

  it('refuses to delete the last remaining slide', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    while (opened.deck.slides.length > 1) expect(deleteSlide(opened, 0)).toBe(true)
    expect(deleteSlide(opened, 0)).toBe(false)
    expect(opened.deck.slides.length).toBe(1)
  })
})

describe('reorderElement', () => {
  it('moves element to front/back and persists order', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const slide = opened.deck.slides[0]!
    const el = addElement(slide, { kind: 'rect', offset: { ...OFF }, fillColor: '#112233' })
    expect(slide.elements.at(-1)!.id).toBe(el.id)

    expect(reorderElement(slide, el.id, 'back')).toBe(true)
    expect(slide.elements[0]!.id).toBe(el.id)
    expect(reorderElement(slide, el.id, 'back')).toBe(false) // already at the bottom
    expect(reorderElement(slide, el.id, 'forward')).toBe(true)
    expect(slide.elements[1]!.id).toBe(el.id)
    expect(reorderElement(slide, el.id, 'front')).toBe(true)
    expect(slide.elements.at(-1)!.id).toBe(el.id)

    reorderElement(slide, el.id, 'back')
    const reopened = await openPptx(await savePptx(opened))
    const first: any = reopened.deck.slides[0]!.elements[0]
    expect(first.fill).toEqual({ type: 'solid', color: '#112233' })
  })
})

describe('slide transitions', () => {
  it('sets, reads back, replaces and clears a transition across save/reopen', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const slide = opened.deck.slides[0]!
    expect(getSlideTransition(slide)).toBe('none')

    setSlideTransition(slide, 'fade')
    expect(getSlideTransition(slide)).toBe('fade')
    expect(slide.bodySuffix).toContain('<p:transition><p:fade/></p:transition>')

    // Replace, don't stack
    setSlideTransition(slide, 'push')
    expect(slide.bodySuffix.match(/<p:transition/g)!.length).toBe(1)
    expect(getSlideTransition(slide)).toBe('push')

    const reopened = await openPptx(await savePptx(opened))
    expect(getSlideTransition(reopened.deck.slides[0]!)).toBe('push')

    setSlideTransition(reopened.deck.slides[0]!, 'none')
    expect(getSlideTransition(reopened.deck.slides[0]!)).toBe('none')
    const again = await openPptx(await savePptx(reopened))
    expect(again.deck.slides[0]!.bodySuffix).not.toContain('<p:transition')
  })

  it('writes morph as AlternateContent(p159:morph) with a fade fallback', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const slide = opened.deck.slides[0]!
    setSlideTransition(slide, 'morph')
    expect(getSlideTransition(slide)).toBe('morph')
    // Newer PowerPoint uses p159:morph; older versions use the Fallback fade
    expect(slide.bodySuffix).toContain('<mc:AlternateContent')
    expect(slide.bodySuffix).toContain('<p159:morph option="byObject"/>')
    expect(slide.bodySuffix).toContain(
      '<mc:Fallback><p:transition spd="slow"><p:fade/></p:transition></mc:Fallback>',
    )
    expect(slide.bodySuffix).toContain('Requires="p159"')

    // Lossless across save/reopen
    const reopened = await openPptx(await savePptx(opened))
    expect(getSlideTransition(reopened.deck.slides[0]!)).toBe('morph')
  })

  it('replaces morph with a plain transition (no AlternateContent left) and clears it', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const slide = opened.deck.slides[0]!
    setSlideTransition(slide, 'morph')
    setSlideTransition(slide, 'wipe')
    expect(slide.bodySuffix).not.toContain('<mc:AlternateContent')
    expect(slide.bodySuffix.match(/<p:transition/g)!.length).toBe(1)
    expect(getSlideTransition(slide)).toBe('wipe')

    setSlideTransition(slide, 'morph')
    setSlideTransition(slide, 'none')
    expect(slide.bodySuffix).not.toContain('<p:transition')
    expect(slide.bodySuffix).not.toContain('<mc:AlternateContent')
  })

  it('coexists with shape animations (morph transition before timing)', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const slide = opened.deck.slides[0]!
    const spid = elementSpid(slide.elements[0]!)!
    setSlideTransition(slide, 'morph')
    setSlideAnimations(slide, [
      { spid, effect: 'fade', trigger: 'onClick', durationMs: 500, delayMs: 0 },
    ])
    const reopened = await openPptx(await savePptx(opened))
    const suffix = reopened.deck.slides[0]!.bodySuffix
    expect(suffix.indexOf('p159:morph')).toBeGreaterThanOrEqual(0)
    expect(suffix.indexOf('p159:morph')).toBeLessThan(suffix.indexOf('<p:timing'))
    expect(getSlideTransition(reopened.deck.slides[0]!)).toBe('morph')
    expect(getSlideAnimations(reopened.deck.slides[0]!)).toHaveLength(1)
  })
})

describe('additional transition effects', () => {
  it('cover/pull/dissolve/zoom write-read round-trip', async () => {
    const { patchSlideTransitionXml, readSlideTransitionXml } = await import('../src/generate')
    for (const kind of ['cover', 'pull', 'dissolve', 'zoom'] as const) {
      const body = patchSlideTransitionXml(
        '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>',
        kind,
      )
      expect(body).toContain(`<p:${kind}`)
      expect(readSlideTransitionXml(body)).toBe(kind)
    }
  })

  it('replacing a transition leaves no stale nodes', async () => {
    const { patchSlideTransitionXml, readSlideTransitionXml } = await import('../src/generate')
    let body = patchSlideTransitionXml(
      '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>',
      'cover',
    )
    body = patchSlideTransitionXml(body, 'zoom')
    expect(body).not.toContain('<p:cover')
    expect(readSlideTransitionXml(body)).toBe('zoom')
  })
})
