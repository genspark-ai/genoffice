/**
 * Mirroring an element writes a:xfrm flipH/flipV, so arrows can point
 * the other way. Rotation cannot express a single-axis mirror, and the render layer
 * already honors the flags — this pins the write-back half.
 */
import { describe, it, expect } from 'vitest'
import { readFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { dirname, join } from 'node:path'
import { addElement, openPptx, savePptx } from '../src/index'
import { patchElementXfrm } from '../src/generate'
import type { TextElement } from '../src/types'

const here = dirname(fileURLToPath(import.meta.url))
const fx = (name: string) => readFileSync(join(here, 'fixtures', name))

const OFF = { x: 914400, y: 914400, cx: 1828800, cy: 914400 }

describe('flip write-back', () => {
  it('flipH/flipV land on a:xfrm and survive save → reopen', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const slide = opened.deck.slides[0]!
    const el = addElement(slide, { kind: 'rightArrow', offset: { ...OFF } })
    el.transform.flipH = true
    el.dirtyTransform = true

    const reopened = await openPptx(await savePptx(opened))
    const el2 = reopened.deck.slides[0]!.elements.at(-1) as TextElement
    expect(el2.transform.flipH).toBe(true)
    expect(el2.transform.flipV).toBe(false)
  })

  it('a single flip is not the same as a 180° rotation', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const slide = opened.deck.slides[0]!
    const el = addElement(slide, { kind: 'line', offset: { ...OFF } })
    el.transform.flipV = true
    el.dirtyTransform = true

    const reopened = await openPptx(await savePptx(opened))
    const el2 = reopened.deck.slides[0]!.elements.at(-1) as TextElement
    expect(el2.transform.flipV).toBe(true)
    expect(el2.transform.rot).toBe(0)
  })

  it('clearing a flip removes the attribute instead of writing flipH="0"', () => {
    const xml =
      '<p:sp><p:spPr><a:xfrm flipH="1" flipV="1"><a:off x="1" y="2"/><a:ext cx="3" cy="4"/></a:xfrm>' +
      '<a:prstGeom prst="line"><a:avLst/></a:prstGeom></p:spPr></p:sp>'
    const el = {
      type: 'shape' as const,
      transform: { offset: { x: 1, y: 2, cx: 3, cy: 4 }, rot: 0, flipH: false, flipV: true },
    }
    const out = patchElementXfrm(el as never, xml)
    expect(out).toContain('<a:xfrm flipV="1">')
    expect(out).not.toContain('flipH')
  })
})
