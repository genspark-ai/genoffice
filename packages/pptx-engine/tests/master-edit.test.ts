/** Master/layout editing: enumeration/parsing, byte-surgery write-back, inheritance taking effect (6.16). */
import { describe, it, expect } from 'vitest'
import { readFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { dirname, join } from 'node:path'
import {
  openPptx,
  savePptx,
  patchSlideXml,
  listMasterParts,
  parseMasterPart,
  materializeSlide,
} from '../src/index'
import type { TextElement } from '../src/types'

const here = dirname(fileURLToPath(import.meta.url))
const fx = (name: string) => readFileSync(join(here, 'fixtures', name))

describe('listMasterParts / parseMasterPart', () => {
  it('enumeration: master first, layouts after; both part kinds parse into elements', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const parts = listMasterParts(opened.archive)
    expect(parts[0]!.kind).toBe('master')
    expect(parts.filter((p) => p.kind === 'layout').length).toBeGreaterThan(1)

    const master = parseMasterPart(opened.archive, parts[0]!.partPath)!
    expect(master.elements.length).toBeGreaterThan(0)
    expect(master.elements.some((e) => e.placeholder === 'title')).toBe(true)

    const layout = parseMasterPart(opened.archive, parts[1]!.partPath)!
    expect(layout.elements.length).toBeGreaterThan(0)
    expect(layout.elements.every((e) => e.transform.offset.cx >= 0)).toBe(true)
  })
})

describe('slide master edit write-back (byte surgery)', () => {
  it('edit master title placeholder text: dirty element regenerated, other elements pass through byte-for-byte, persists after savePptx', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const parts = listMasterParts(opened.archive)
    const masterPath = parts[0]!.partPath
    const layoutPath = parts[1]!.partPath
    const layoutBefore = opened.archive.readText(layoutPath)!
    const master = parseMasterPart(opened.archive, masterPath)!

    const title = master.elements.find((e) => e.placeholder === 'title') as TextElement
    title.text!.paragraphs[0]!.runs[0]!.text = 'MASTER-EDITED'
    title.dirty = true

    const patched = patchSlideXml(master)
    expect(patched).toContain('MASTER-EDITED')
    // Untouched elements preserved byte-for-byte
    for (const el of master.elements) {
      if (el === title) continue
      expect(patched).toContain(el.anchor.originalXml)
    }
    opened.archive.entries.set(masterPath, Buffer.from(patched, 'utf8'))

    const reopened = await openPptx(await savePptx(opened))
    expect(reopened.archive.readText(masterPath)).toContain('MASTER-EDITED')
    // Unedited layout parts are byte-identical
    expect(reopened.archive.readText(layoutPath)).toBe(layoutBefore)
  })

  it('edit layout placeholder geometry: slide placeholders without explicit xfrm inherit the new geometry, still there after save and reopen', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const slide0 = opened.deck.slides[0]!
    const slideTitle = slide0.elements.find((e) => e.placeholder === 'ctrTitle')!
    expect(slideTitle.anchor.originalXml).not.toContain('<a:xfrm') // precondition: geometry is inherited

    const layoutPath = slide0.layoutPath!
    const layout = parseMasterPart(opened.archive, layoutPath)!
    const layoutTitle = layout.elements.find((e) => e.placeholder === 'ctrTitle')!
    const newX = 1234567
    layoutTitle.transform = {
      ...layoutTitle.transform,
      offset: { ...layoutTitle.transform.offset, x: newX },
    }
    layoutTitle.dirtyTransform = true
    opened.archive.entries.set(layoutPath, Buffer.from(patchSlideXml(layout), 'utf8'))

    // Inheritance takes effect immediately after reparse
    const fresh = materializeSlide(opened, 0)!
    const freshTitle = fresh.elements.find((e) => e.placeholder === 'ctrTitle')!
    expect(freshTitle.transform.offset.x).toBe(newX)

    // Still in effect after save and reopen
    const reopened = await openPptx(await savePptx(opened))
    const t2 = reopened.deck.slides[0]!.elements.find((e) => e.placeholder === 'ctrTitle')!
    expect(t2.transform.offset.x).toBe(newX)
  })

  it('edit master placeholder text color (run-level patch): part parses with the new color after save and reopen', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const masterPath = listMasterParts(opened.archive)[0]!.partPath
    const master = parseMasterPart(opened.archive, masterPath)!
    const title = master.elements.find((e) => e.placeholder === 'title') as TextElement
    for (const p of title.text!.paragraphs)
      for (const r of p.runs) {
        r.color = '#FF0000'
        // user explicitly recolored: no longer follows the theme or inheritance
        delete r.colorFollowsTheme
        delete r.colorInherited
      }
    title.dirty = true
    opened.archive.entries.set(masterPath, Buffer.from(patchSlideXml(master), 'utf8'))

    const reopened = await openPptx(await savePptx(opened))
    const m2 = parseMasterPart(reopened.archive, masterPath)!
    const t2 = m2.elements.find((e) => e.placeholder === 'title') as TextElement
    expect(t2.text!.paragraphs[0]!.runs[0]!.color?.toUpperCase()).toBe('#FF0000')
  })
})
