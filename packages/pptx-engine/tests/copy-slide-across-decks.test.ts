import { describe, it, expect } from 'vitest'
import { readFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { dirname, join } from 'node:path'
import {
  openPptx,
  savePptx,
  copySlide,
  pasteSlide,
  listLayouts,
  chooseLayout,
  materializeSlideBundle,
} from '../src/index'
import type { SlideBundle } from '../src/index'

const here = dirname(fileURLToPath(import.meta.url))
const fx = (name: string) => readFileSync(join(here, 'fixtures', name))

const slideText = (slide: any): string =>
  slide.elements
    .filter((e: any) => e.type === 'text' || e.type === 'shape')
    .flatMap((e: any) => e.text?.paragraphs ?? [])
    .flatMap((p: any) => p.runs)
    .map((r: any) => r.text)
    .join('')

describe('copySlide / pasteSlide across decks', () => {
  it('carries a slide into another deck and survives save → reopen', async () => {
    const source = await openPptx(fx('05_unicode_cjk_emoji.pptx'))
    const target = await openPptx(fx('01_standard_business.pptx'))
    const sourceText = slideText(source.deck.slides[0])
    const targetCount = target.deck.slides.length

    const bundle = copySlide(source, 0)
    expect(bundle).not.toBeNull()
    const pasted = pasteSlide(target, 0, bundle!)
    expect(pasted).not.toBeNull()
    expect(target.deck.slides.length).toBe(targetCount + 1)
    expect(target.deck.slides[1]).toBe(pasted)

    const reopened = await openPptx(await savePptx(target))
    expect(reopened.deck.slides.length).toBe(targetCount + 1)
    expect(slideText(reopened.deck.slides[1])).toBe(sourceText)
  })

  it('points the pasted slide at a layout that exists in the destination', async () => {
    const source = await openPptx(fx('05_unicode_cjk_emoji.pptx'))
    const target = await openPptx(fx('01_standard_business.pptx'))
    const bundle = copySlide(source, 0)!
    const pasted = pasteSlide(target, 0, bundle)!

    const layoutRel = [...target.archive.readRels(pasted.path).values()].find((rel) =>
      rel.type.endsWith('/slideLayout'),
    )
    expect(layoutRel).toBeDefined()
    const layoutPath = `ppt/${layoutRel!.target.replace('../', '')}`
    expect(listLayouts(target.archive).map((l) => l.path)).toContain(layoutPath)
    // the source's own layout part must not have been dragged along
    expect(target.archive.has('ppt/slideLayouts/slideLayout99.xml')).toBe(false)
  })

  it('copies referenced media under a fresh name without clobbering the target', async () => {
    const source = await openPptx(fx('05_unicode_cjk_emoji.pptx'))
    const target = await openPptx(fx('01_standard_business.pptx'))
    const mediaBefore = [...target.archive.entries.keys()].filter((p) => p.startsWith('ppt/media/'))
    const sourceMedia = [...source.archive.entries.keys()].filter((p) => p.startsWith('ppt/media/'))

    const bundle = copySlide(source, 0)!
    const pasted = pasteSlide(target, 0, bundle)!
    const rels = [...target.archive.readRels(pasted.path).values()]
    const imageRels = rels.filter((rel) => rel.type.endsWith('/image'))

    for (const rel of imageRels) {
      const path = `ppt/${rel.target.replace('../', '')}`
      expect(target.archive.has(path)).toBe(true)
    }
    // nothing that already existed in the target was overwritten
    for (const path of mediaBefore) expect(target.archive.has(path)).toBe(true)
    if (sourceMedia.length > 0 && imageRels.length > 0) {
      const after = [...target.archive.entries.keys()].filter((p) => p.startsWith('ppt/media/'))
      expect(after.length).toBeGreaterThan(mediaBefore.length)
    }

    // every content type the new parts need is declared
    const ct = target.archive.readText('[Content_Types].xml')!
    for (const rel of imageRels) {
      const ext = rel.target.slice(rel.target.lastIndexOf('.') + 1).toLowerCase()
      expect(ct.toLowerCase()).toContain(`extension="${ext}"`)
    }
  })

  it('keeps rIds stable so the pasted slide XML needs no rewriting', async () => {
    const source = await openPptx(fx('05_unicode_cjk_emoji.pptx'))
    const target = await openPptx(fx('01_standard_business.pptx'))
    const sourceRelIds = [...source.archive.readRels(source.deck.slides[0]!.path).values()]
      .filter((rel) => !rel.type.endsWith('/notesSlide'))
      .map((rel) => rel.id)
      .sort()

    const pasted = pasteSlide(target, 0, copySlide(source, 0)!)!
    const pastedRelIds = [...target.archive.readRels(pasted.path).values()]
      .map((rel) => rel.id)
      .sort()
    expect(pastedRelIds).toEqual(sourceRelIds)
  })

  it('gives each bundled media part its own destination when names collide with the target', async () => {
    const target = await openPptx(fx('01_standard_business.pptx'))
    // the target now occupies both names the bundle wants
    target.archive.entries.set('ppt/media/image2.png', new Uint8Array([9, 9, 9]))

    const IMAGE_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'
    const LAYOUT_REL =
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout'
    const bundle: SlideBundle = {
      slideXml: '<p:sld/>',
      rels: [
        { id: 'rId1', type: LAYOUT_REL, target: 'ppt/slideLayouts/slideLayout1.xml', layout: true },
        { id: 'rId2', type: IMAGE_REL, target: 'ppt/media/image1.png' },
        { id: 'rId3', type: IMAGE_REL, target: 'ppt/media/image2.png' },
      ],
      parts: {
        'ppt/media/image1.png': Buffer.from('AAAA-image-A').toString('base64'),
        'ppt/media/image2.png': Buffer.from('BBBB-image-B').toString('base64'),
      },
      contentTypes: { png: 'image/png' },
    }

    const layoutPath = chooseLayout(target.archive, bundle)!
    const relsXml = materializeSlideBundle(
      target.archive,
      bundle,
      'ppt/slides/slide99.xml',
      layoutPath,
    )

    const targets = [...relsXml.matchAll(/Target="(\.\.\/media\/[^"]+)"/g)].map((m) => m[1]!)
    expect(new Set(targets).size).toBe(2)
    const bytesAt = (t: string) =>
      Buffer.from(target.archive.readBytes('ppt/' + t.replace('../', ''))!).toString()
    expect(bytesAt(targets[0]!)).toBe('AAAA-image-A')
    expect(bytesAt(targets[1]!)).toBe('BBBB-image-B')
  })

  it('pastes at the front when afterIndex is -1', async () => {
    const source = await openPptx(fx('05_unicode_cjk_emoji.pptx'))
    const target = await openPptx(fx('01_standard_business.pptx'))
    const sourceText = slideText(source.deck.slides[0])

    const pasted = pasteSlide(target, -1, copySlide(source, 0)!)
    expect(target.deck.slides[0]).toBe(pasted)
    const reopened = await openPptx(await savePptx(target))
    expect(slideText(reopened.deck.slides[0])).toBe(sourceText)
  })

  it('keep-source-formatting imports the source layout chain and registers its master', async () => {
    // forge a source deck whose theme differs from the target's
    const source = await openPptx(fx('05_unicode_cjk_emoji.pptx'))
    const themeXml = source.archive.readText('ppt/theme/theme1.xml')!
    source.archive.entries.set(
      'ppt/theme/theme1.xml',
      Buffer.from(themeXml.replace('4F81BD', '123456'), 'utf8'),
    )
    const target = await openPptx(fx('01_standard_business.pptx'))
    const masters = () =>
      [...target.archive.entries.keys()].filter((p) =>
        /^ppt\/slideMasters\/slideMaster\d+\.xml$/.test(p),
      ).length
    const mastersBefore = masters()

    const bundle = copySlide(source, 0)!
    expect(bundle.chain).toBeDefined()
    const pasted = pasteSlide(target, 0, bundle, { keepSourceFormatting: true })
    expect(pasted).not.toBeNull()

    // a new master arrived and is registered in presentation.xml(.rels)
    expect(masters()).toBe(mastersBefore + 1)
    const pres = target.archive.readText('ppt/presentation.xml')!
    expect([...pres.matchAll(/<p:sldMasterId\s/g)].length).toBe(mastersBefore + 1)

    // the pasted slide points at the imported layout (byte-identical to the source's)
    const layoutRel = [...target.archive.readRels(pasted!.path).values()].find((rel) =>
      rel.type.endsWith('/slideLayout'),
    )!
    const layoutPath = `ppt/${layoutRel.target.replace('../', '')}`
    expect(target.archive.readText(layoutPath)).toBe(
      source.archive.readText(bundle.chain!.layoutPath),
    )

    // survives save → reopen
    const reopened = await openPptx(await savePptx(target))
    expect(reopened.deck.slides.length).toBe(target.deck.slides.length)
  })

  it('keep-source-formatting reuses an identical existing master instead of importing', async () => {
    // the two fixtures share one template: the chain must be reused, not duplicated
    const source = await openPptx(fx('05_unicode_cjk_emoji.pptx'))
    const target = await openPptx(fx('01_standard_business.pptx'))
    const masterParts = () =>
      [...target.archive.entries.keys()].filter((p) => p.startsWith('ppt/slideMasters/')).length
    const mastersBefore = masterParts()

    const bundle = copySlide(source, 0)!
    expect(pasteSlide(target, 0, bundle, { keepSourceFormatting: true })).not.toBeNull()
    expect(masterParts()).toBe(mastersBefore)
  })

  it('pasting the same foreign chain twice imports its master only once', async () => {
    const source = await openPptx(fx('05_unicode_cjk_emoji.pptx'))
    const themeXml = source.archive.readText('ppt/theme/theme1.xml')!
    source.archive.entries.set(
      'ppt/theme/theme1.xml',
      Buffer.from(themeXml.replace('4F81BD', '123456'), 'utf8'),
    )
    const target = await openPptx(fx('01_standard_business.pptx'))
    const masters = () =>
      [...target.archive.entries.keys()].filter((p) =>
        /^ppt\/slideMasters\/slideMaster\d+\.xml$/.test(p),
      ).length
    const mastersBefore = masters()

    const bundle = copySlide(source, 0)!
    pasteSlide(target, 0, bundle, { keepSourceFormatting: true })
    pasteSlide(target, 1, bundle, { keepSourceFormatting: true })
    expect(masters()).toBe(mastersBefore + 1)
  })

  it('keep-source-formatting falls back to the destination theme without a chain', async () => {
    const source = await openPptx(fx('05_unicode_cjk_emoji.pptx'))
    const target = await openPptx(fx('01_standard_business.pptx'))
    const { chain: _chain, ...rest } = copySlide(source, 0)!
    const pasted = pasteSlide(target, 0, rest as SlideBundle, { keepSourceFormatting: true })
    expect(pasted).not.toBeNull()
    const layoutRel = [...target.archive.readRels(pasted!.path).values()].find((rel) =>
      rel.type.endsWith('/slideLayout'),
    )!
    const layoutPath = `ppt/${layoutRel.target.replace('../', '')}`
    expect(listLayouts(target.archive).map((l) => l.path)).toContain(layoutPath)
  })

  it('pasting twice produces two independent slides', async () => {
    const source = await openPptx(fx('05_unicode_cjk_emoji.pptx'))
    const target = await openPptx(fx('01_standard_business.pptx'))
    const bundle = copySlide(source, 0)!
    const first = pasteSlide(target, 0, bundle)!
    const second = pasteSlide(target, 1, bundle)!
    expect(first.path).not.toBe(second.path)

    const reopened = await openPptx(await savePptx(target))
    expect(reopened.deck.slides.length).toBe(target.deck.slides.length)
    expect(slideText(reopened.deck.slides[1])).toBe(slideText(reopened.deck.slides[2]))
  })
})
