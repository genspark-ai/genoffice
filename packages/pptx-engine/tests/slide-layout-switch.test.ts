/** Switching an existing slide's layout + resetting the layout. */
import { describe, it, expect } from 'vitest'
import { readFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { dirname, join } from 'node:path'
import { listSlideLayouts } from '../src/layout'
import { openPptx, savePptx, setSlideLayout, resetSlideLayout } from '../src/index'
import { relsPathFor } from '../src/zip'
import type { TextElement } from '../src/types'

const here = dirname(fileURLToPath(import.meta.url))
const fx = (name: string) => readFileSync(join(here, 'fixtures', name))

describe('setSlideLayout / resetSlideLayout', () => {
  it('switch layout: rels point to the new layout, effective after save and reopen', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const layouts = listSlideLayouts(opened.archive)
    expect(layouts.length).toBeGreaterThan(1)
    const slide0 = opened.deck.slides[0]!
    const curRel = opened.archive.readText(relsPathFor(slide0.path))!
    const target = layouts.find((l) => !curRel.includes(l.path.slice('ppt/'.length)))!
    const fresh = setSlideLayout(opened, 0, target.path)
    expect(fresh).not.toBeNull()
    const newRel = opened.archive.readText(relsPathFor(fresh!.path))!
    expect(newRel).toContain(target.path.slice('ppt/'.length))

    const reopened = await openPptx(await savePptx(opened))
    const rel2 = reopened.archive.readText(relsPathFor(reopened.deck.slides[0]!.path))!
    expect(rel2).toContain(target.path.slice('ppt/'.length))
  })

  it('reset layout: explicit placeholder xfrm removed, geometry falls back to layout inheritance', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const slide = opened.deck.slides[0]!
    const ph = slide.elements.find(
      (e) => e.placeholder && (e.type === 'text' || e.type === 'shape'),
    ) as TextElement
    expect(ph).toBeTruthy()
    // First move the placeholder away (writes an explicit xfrm)
    ph.transform.offset = { x: 111, y: 222, cx: 3333333, cy: 444444 }
    ph.dirtyTransform = true
    const fresh = resetSlideLayout(opened, 0)
    expect(fresh).not.toBeNull()
    const ph2 = fresh!.elements.find((e) => e.placeholder === ph.placeholder)!
    // After reset, geometry comes from layout inheritance (no longer the moved value)
    expect(ph2.transform.offset.x).not.toBe(111)
    expect(ph2.transform.offset.cx).toBeGreaterThan(0)
    expect(ph2.anchor.originalXml).not.toContain('<a:xfrm')
  })
})
