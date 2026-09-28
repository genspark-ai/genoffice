/** Phase 1: Insertable straight lines / arrow lines as p:cxnSp connectors. */
import { describe, it, expect } from 'vitest'
import { openPptx, savePptx, createBlankPptx } from '../src/index'
import { addElement, isLineKind } from '../src/insert'
import type { TextElement } from '../src/types'

describe('line insertion (p:cxnSp)', () => {
  it('recognizes line kinds', () => {
    expect(isLineKind('line')).toBe(true)
    expect(isLineKind('lineArrow')).toBe(true)
    expect(isLineKind('lineArrowDouble')).toBe(true)
    expect(isLineKind('rect')).toBe(false)
    expect(isLineKind('textbox')).toBe(false)
  })

  it('plain line: cxnSp with prst=line, stroke only, zero height allowed', async () => {
    const opened = await openPptx(await createBlankPptx())
    const slide = opened.deck.slides[0]!
    const el = addElement(slide, {
      kind: 'line',
      offset: { x: 100, y: 200, cx: 3000000, cy: 0 },
      stroke: { color: '#ff0000', widthEmu: 12700 },
    })
    expect(el.type).toBe('shape')
    expect(el.presetGeometry).toBe('line')
    expect(el.fill).toEqual({ type: 'none' })
    expect(el.stroke?.headEnd).toBeUndefined()
    expect(el.stroke?.tailEnd).toBeUndefined()
    const xml = el.anchor.originalXml
    expect(xml).toContain('<p:cxnSp>')
    expect(xml).toContain('<a:prstGeom prst="line">')
    expect(xml).toContain('<a:srgbClr val="FF0000"/>')
    expect(xml).not.toContain('txBody')
    expect(xml).not.toContain('headEnd')
    expect(xml).not.toContain('tailEnd')
  })

  it('arrow lines: straightConnector1 with tailEnd / headEnd+tailEnd', async () => {
    const opened = await openPptx(await createBlankPptx())
    const slide = opened.deck.slides[0]!
    const single = addElement(slide, { kind: 'lineArrow', offset: { x: 0, y: 0, cx: 100, cy: 50 } })
    expect(single.presetGeometry).toBe('straightConnector1')
    expect(single.stroke?.tailEnd?.type).toBe('triangle')
    expect(single.stroke?.headEnd).toBeUndefined()
    expect(single.anchor.originalXml).toContain('<a:tailEnd type="triangle"')

    const dbl = addElement(slide, {
      kind: 'lineArrowDouble',
      offset: { x: 0, y: 0, cx: 100, cy: 50 },
    })
    expect(dbl.stroke?.headEnd?.type).toBe('triangle')
    expect(dbl.stroke?.tailEnd?.type).toBe('triangle')
    expect(dbl.anchor.originalXml).toContain('<a:headEnd type="triangle"')
  })

  it('bent/curved connectors insert with their preset geometry (phase 2)', async () => {
    const opened = await openPptx(await createBlankPptx())
    const slide = opened.deck.slides[0]!
    expect(isLineKind('lineBent')).toBe(true)
    expect(isLineKind('lineCurved')).toBe(true)
    const bent = addElement(slide, { kind: 'lineBent', offset: { x: 0, y: 0, cx: 100, cy: 80 } })
    expect(bent.presetGeometry).toBe('bentConnector3')
    expect(bent.anchor.originalXml).toContain('<a:prstGeom prst="bentConnector3">')
    expect(bent.anchor.originalXml).toContain('name="Elbow Connector')
    const curved = addElement(slide, {
      kind: 'lineCurved',
      offset: { x: 0, y: 0, cx: 100, cy: 80 },
    })
    expect(curved.presetGeometry).toBe('curvedConnector3')
    expect(curved.anchor.originalXml).toContain('<a:prstGeom prst="curvedConnector3">')
    expect(curved.anchor.originalXml).toContain('name="Curved Connector')
  })

  it('round-trips through save and reparses as a connector', async () => {
    const opened = await openPptx(await createBlankPptx())
    const slide = opened.deck.slides[0]!
    addElement(slide, {
      kind: 'lineArrow',
      offset: { x: 914400, y: 914400, cx: 1828800, cy: 0 },
      stroke: { color: '#0070c0', widthEmu: 19050 },
    })
    const reopened = await openPptx(Buffer.from(await savePptx(opened)))
    const parsed = reopened.deck.slides[0]!.elements.find(
      (e) => (e as TextElement).presetGeometry === 'straightConnector1',
    ) as TextElement | undefined
    expect(parsed).toBeDefined()
    expect(parsed!.type).toBe('shape')
    expect(parsed!.fill).toEqual({ type: 'none' })
    expect(parsed!.stroke?.tailEnd?.type).toBe('triangle')
    expect(parsed!.stroke?.width).toBe(19050)
    expect(parsed!.transform.offset).toEqual({ x: 914400, y: 914400, cx: 1828800, cy: 0 })
  })

  it('flip flags reach a preset shape model as well as its xml', async () => {
    const opened = await openPptx(await createBlankPptx())
    const slide = opened.deck.slides[0]!
    const el = addElement(slide, {
      kind: 'rightArrow',
      offset: { x: 100, y: 200, cx: 3000000, cy: 900000 },
      flipH: true,
    })
    expect(el.transform.flipH).toBe(true)
    expect(el.transform.flipV).toBe(false)
    expect(el.anchor.originalXml).toContain('flipH="1"')
  })

  it('flip flags mirror the connector and survive save → reopen', async () => {
    const opened = await openPptx(await createBlankPptx())
    const slide = opened.deck.slides[0]!
    const el = addElement(slide, {
      kind: 'line',
      offset: { x: 100, y: 200, cx: 3000000, cy: 900000 },
      stroke: { color: '#0070c0', widthEmu: 12700 },
      flipV: true,
    })
    expect(el.transform.flipV).toBe(true)
    expect(el.transform.flipH).toBe(false)
    expect(el.anchor.originalXml).toContain('flipV="1"')
    const reopened = await openPptx(Buffer.from(await savePptx(opened)))
    const parsed = reopened.deck.slides[0]!.elements.find(
      (e) => (e as TextElement).presetGeometry === 'line',
    ) as TextElement | undefined
    expect(parsed?.transform.flipV).toBe(true)
  })
})
