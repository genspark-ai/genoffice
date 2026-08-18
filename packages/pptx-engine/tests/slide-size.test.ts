/** Slide size switching: sldSz write-back + per-axis proportional rescale + round-trip idempotence. */
import { describe, it, expect } from 'vitest'
import { openPptx, savePptx, setSlideSize, createBlankPptx, addElement } from '../src/index'

describe('setSlideSize', () => {
  it('16:9 → 4:3: sldSz updated, elements rescaled per axis (full-bleed stays full-bleed)', async () => {
    const opened = await openPptx(await createBlankPptx()) // blank is 16:9
    const slide = opened.deck.slides[0]!
    addElement(slide, {
      kind: 'rect',
      offset: { x: 0, y: 0, cx: 12192000, cy: 6858000 },
      fillColor: '#123456',
    })
    expect(setSlideSize(opened, 9144000, 6858000)).toBe(true)
    expect(opened.deck.size).toEqual({ cx: 9144000, cy: 6858000 })

    const el = slide.elements[0]!
    expect(el.transform.offset).toMatchObject({ x: 0, y: 0, cx: 9144000, cy: 6858000 })
    expect(el.dirtyTransform).toBe(true)

    const reopened = await openPptx(await savePptx(opened))
    expect(reopened.deck.size.cx).toBe(9144000)
  })

  it('4:3 → 16:9: content rescales with the canvas (wider, same height)', async () => {
    const opened = await openPptx(await createBlankPptx())
    setSlideSize(opened, 9144000, 6858000)
    const slide = opened.deck.slides[0]!
    addElement(slide, {
      kind: 'rect',
      offset: { x: 900, y: 200, cx: 3000, cy: 400 },
      fillColor: '#123456',
    })
    expect(setSlideSize(opened, 12192000, 6858000)).toBe(true)
    const el = slide.elements[0]!
    const sx = 12192000 / 9144000
    expect(el.transform.offset).toMatchObject({
      x: Math.round(900 * sx),
      y: 200,
      cx: Math.round(3000 * sx),
      cy: 400,
    })
  })

  it('A → B → A restores the original layout (idempotent round-trip)', async () => {
    const opened = await openPptx(await createBlankPptx())
    const slide = opened.deck.slides[0]!
    addElement(slide, {
      kind: 'rect',
      offset: { x: 1234567, y: 765432, cx: 4567890, cy: 2345678 },
      fillColor: '#123456',
    })
    const before = { ...slide.elements[0]!.transform.offset }
    // 16:9 → 3:4 portrait → back (the issue's repro shape)
    setSlideSize(opened, 6858000, 9144000)
    setSlideSize(opened, 12192000, 6858000)
    const after = opened.deck.slides[0]!.elements[0]!.transform.offset
    expect(Math.abs(after.x - before.x)).toBeLessThanOrEqual(2)
    expect(Math.abs(after.y - before.y)).toBeLessThanOrEqual(2)
    expect(Math.abs(after.cx - before.cx)).toBeLessThanOrEqual(2)
    expect(Math.abs(after.cy - before.cy)).toBeLessThanOrEqual(2)
    expect(opened.deck.size).toEqual({ cx: 12192000, cy: 6858000 })
  })

  it('same size rejected (no empty history step created)', async () => {
    const opened = await openPptx(await createBlankPptx())
    expect(setSlideSize(opened, 12192000, 6858000)).toBe(false)
  })
})

describe('setSlideSize scales masters in sync', () => {
  it('when shrinking, layout/master elements scale proportionally and center, part bytes written back', async () => {
    const { listMasterParts, parseMasterPart } = await import('../src/master-edit')
    const { readFileSync } = await import('node:fs')
    const { fileURLToPath } = await import('node:url')
    const { dirname, join } = await import('node:path')
    const here = dirname(fileURLToPath(import.meta.url))
    const opened = await openPptx(readFileSync(join(here, 'fixtures', '01_standard_business.pptx')))
    const parts = listMasterParts(opened.archive)
    expect(parts.length).toBeGreaterThan(0)
    const before = parseMasterPart(opened.archive, parts[0]!.partPath)!
    const el0 = before.elements.find((e) => e.transform.offset.cx > 0)
    expect(el0).toBeTruthy()
    const oldCx = el0!.transform.offset.cx

    const oldSize = { ...opened.deck.size }
    setSlideSize(opened, 9144000, 6858000)
    const after = parseMasterPart(opened.archive, parts[0]!.partPath)!
    // Compute the ratio from the file's actual sldSz (generator may not write the standard 12192000)
    const sx = 9144000 / oldSize.cx
    // Element order is stable after reparse; compare scale ratios index by index
    before.elements.forEach((b, i) => {
      // Passthrough without p:xfrm cannot be patched; skip it (the engine skips it too)
      if (!b.transform.offset.cx || b.type === 'passthrough') return
      const a = after.elements[i]!
      expect(
        Math.abs(a.transform.offset.cx - Math.round(b.transform.offset.cx * sx)),
      ).toBeLessThanOrEqual(1)
    })
    expect(oldCx).toBeGreaterThan(0)
  })

  it('when enlarging, master elements rescale in sync with the slides', async () => {
    const { listMasterParts, parseMasterPart } = await import('../src/master-edit')
    const { readFileSync } = await import('node:fs')
    const { fileURLToPath } = await import('node:url')
    const { dirname, join } = await import('node:path')
    const here = dirname(fileURLToPath(import.meta.url))
    const opened = await openPptx(readFileSync(join(here, 'fixtures', '01_standard_business.pptx')))
    const parts = listMasterParts(opened.archive)
    expect(parts.length).toBeGreaterThan(0)
    const before = parseMasterPart(opened.archive, parts[0]!.partPath)!

    const old = { ...opened.deck.size }
    expect(setSlideSize(opened, old.cx * 2, old.cy)).toBe(true)
    const after = parseMasterPart(opened.archive, parts[0]!.partPath)!
    before.elements.forEach((b, i) => {
      if (!b.transform.offset.cx || b.type === 'passthrough') return
      const a = after.elements[i]!
      expect(
        Math.abs(a.transform.offset.cx - Math.round(b.transform.offset.cx * 2)),
      ).toBeLessThanOrEqual(1)
      expect(a.transform.offset.cy).toBe(b.transform.offset.cy)
    })
  })
})

describe('setElementImageFill', () => {
  it('shape fill replaced by blipFill, media/rels land in the package', async () => {
    const { setElementImageFill } = await import('../src/index')
    const opened = await openPptx(await createBlankPptx())
    const slide = opened.deck.slides[0]!
    const r = addElement(slide, {
      kind: 'rect',
      offset: { x: 0, y: 0, cx: 1000, cy: 1000 },
      fillColor: '#FF0000',
    })
    const png = Buffer.from(
      'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==',
      'base64',
    )
    const mediaPath = setElementImageFill(opened, slide, r.id, { bytes: png, ext: 'png' })
    expect(mediaPath).toMatch(/^ppt\/media\/image\d+\.png$/)
    const el = slide.elements.find((e) => e.id === r.id)!
    expect(el.anchor.originalXml).toContain('<a:blipFill rotWithShape="1"><a:blip r:embed="rId')
    expect(el.anchor.originalXml).not.toContain('FF0000')
    expect((el as import('../src/types').TextElement).fill).toMatchObject({
      type: 'image',
      mediaRef: mediaPath,
    })
    // Save and reopen: media part exists, blipFill parses back to a picture fill
    const reopened = await openPptx(await savePptx(opened))
    expect(reopened.archive.entries.has(mediaPath!)).toBe(true)
  })

  it('tile mode + landed-media reuse: second shape adds no media part, both tile', async () => {
    const { setElementImageFill } = await import('../src/index')
    const opened = await openPptx(await createBlankPptx())
    const slide = opened.deck.slides[0]!
    const a = addElement(slide, {
      kind: 'rect',
      offset: { x: 0, y: 0, cx: 1000, cy: 1000 },
      fillColor: '#FF0000',
    })
    const b = addElement(slide, {
      kind: 'rect',
      offset: { x: 2000, y: 0, cx: 1000, cy: 1000 },
      fillColor: '#00FF00',
    })
    const png = Buffer.from(
      'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==',
      'base64',
    )
    const mediaCount = () =>
      [...opened.archive.entries.keys()].filter((p) => p.startsWith('ppt/media/')).length
    const before = mediaCount()
    const landed = setElementImageFill(
      opened,
      slide,
      a.id,
      { bytes: png, ext: 'png' },
      {
        tile: true,
      },
    )
    expect(landed).toMatch(/^ppt\/media\/image\d+\.png$/)
    expect(mediaCount()).toBe(before + 1)
    // second target reuses the landed part: only a fill swap, no new media
    const reused = setElementImageFill(opened, slide, b.id, { mediaPath: landed! }, { tile: true })
    expect(reused).toBe(landed)
    expect(mediaCount()).toBe(before + 1)
    for (const id of [a.id, b.id]) {
      const el = slide.elements.find((e) => e.id === id)!
      expect(el.anchor.originalXml).toContain('<a:tile')
      // tile metrics mirror what parse restores, so pre-save render scales tiles correctly
      expect((el as import('../src/types').TextElement).fill).toMatchObject({
        type: 'image',
        mode: 'tile',
        tile: { tx: 0, ty: 0, sx: 1, sy: 1, algn: 'tl' },
      })
    }
    // a failed lookup must not land an orphaned media part
    expect(setElementImageFill(opened, slide, 'no-such-id', { bytes: png, ext: 'png' })).toBeNull()
    expect(mediaCount()).toBe(before + 1)
  })
})
