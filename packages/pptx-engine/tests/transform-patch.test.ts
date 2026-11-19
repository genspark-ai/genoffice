import { describe, it, expect } from 'vitest'
import { readFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { dirname, join } from 'node:path'
import { openPptx, savePptx, patchElementXfrm } from '../src/index'
import type { SlideElement } from '../src/types'

const here = dirname(fileURLToPath(import.meta.url))
const fx = (name: string) => readFileSync(join(here, 'fixtures', name))

const OFF = { x: 914400, y: 457200, cx: 3657600, cy: 1828800 }

describe('dirtyTransform xfrm patch', () => {
  it('save → reopen round-trips new geometry', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const slide = opened.deck.slides[0]!
    const el = slide.elements.find((e) => e.type === 'text' || e.type === 'shape')!
    el.transform = { ...el.transform, offset: { ...OFF }, rot: 15 * 60000 }
    el.dirtyTransform = true

    const reopened = await openPptx(await savePptx(opened))
    const el2 = reopened.deck.slides[0]!.elements.find(
      (e) => e.anchor.spIndex === el.anchor.spIndex,
    )!
    expect(el2.transform.offset).toEqual(OFF)
    expect(el2.transform.rot).toBe(15 * 60000)
  })

  it('text dirty + transform dirty compose on the same element', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const slide = opened.deck.slides.find((s) =>
      s.elements.some((e: any) => (e.type === 'text' || e.type === 'shape') && e.text?.paragraphs?.length),
    )!
    const el: any = slide.elements.find(
      (e: any) => (e.type === 'text' || e.type === 'shape') && e.text?.paragraphs?.length,
    )!
    el.text.paragraphs = [{ runs: [{ text: 'combined patch', bold: true }] }]
    el.dirty = true
    el.transform = { ...el.transform, offset: { ...OFF }, rot: 0 }
    el.dirtyTransform = true

    const reopened = await openPptx(await savePptx(opened))
    const slideIdx = opened.deck.slides.indexOf(slide)
    const el2: any = reopened.deck.slides[slideIdx]!.elements.find(
      (e) => e.anchor.spIndex === el.anchor.spIndex,
    )!
    expect(el2.transform.offset).toEqual(OFF)
    const text = el2.text.paragraphs.flatMap((p: any) => p.runs).map((r: any) => r.text).join('')
    expect(text).toBe('combined patch')
  })

  it('injects <a:xfrm> when the sp inherits geometry (no xfrm in xml)', () => {
    const el = {
      type: 'text',
      transform: { offset: { ...OFF }, rot: 0, flipH: false, flipV: true },
    } as unknown as SlideElement
    const xml = '<p:sp><p:nvSpPr/><p:spPr><a:prstGeom prst="rect"/></p:spPr></p:sp>'
    const out = patchElementXfrm(el, xml)
    expect(out).toContain('<a:xfrm flipV="1"><a:off x="914400" y="457200"/>')
    expect(out.indexOf('<a:xfrm')).toBeLessThan(out.indexOf('<a:prstGeom'))
  })

  it('keeps group chOff/chExt while replacing off/ext', () => {
    const el = {
      type: 'group',
      transform: { offset: { ...OFF }, rot: 0, flipH: false, flipV: false },
    } as unknown as SlideElement
    const xml =
      '<p:grpSp><p:grpSpPr><a:xfrm rot="600000"><a:off x="1" y="2"/><a:ext cx="3" cy="4"/>' +
      '<a:chOff x="0" y="0"/><a:chExt cx="3" cy="4"/></a:xfrm></p:grpSpPr></p:grpSp>'
    const out = patchElementXfrm(el, xml)
    expect(out).toContain('<a:off x="914400" y="457200"/>')
    expect(out).toContain('<a:chOff x="0" y="0"/>')
    expect(out).toContain('<a:chExt cx="3" cy="4"/>')
    expect(out).not.toContain('rot="600000"')
  })
})
