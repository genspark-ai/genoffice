/**
 * Linked pictures of a header as Edit > Links has them: listed from both shapes a header
 * picture is kept in, and broken, relinked and updated in the part the save writes.
 */
import { describe, expect, it } from 'vitest'
import JSZip from 'jszip'
import { parseDocx, readSections, saveDocx } from '@genoffice/docx-engine'
import { buildDocx, TINY_PNG_BASE64 } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc, pmDocToSavePlan } from '../src/renderer/editor/convert'
import { hfPictureLink, linkSourceOf, linkTypeOf } from '../src/renderer/editor/link-fields'
import {
  hfPictureLinkRows,
  hfSaveOptions,
  resolveHf,
  withHfEdit,
  withHfPictureLink,
  type HfSectionState,
} from '../src/renderer/hf-sections'

const IMAGE_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'
const OTHER_PNG =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg=='
const graphic = (id: number, link: string) =>
  `<wp:extent cx="914400" cy="457200"/><wp:docPr id="${id}" name="Picture ${id}"/>` +
  '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
  `<pic:pic><pic:blipFill><a:blip r:embed="rId1" r:link="${link}"/></pic:blipFill>` +
  '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="457200"/></a:xfrm></pic:spPr></pic:pic>' +
  '</a:graphicData></a:graphic>'
/** a floating logo alone in its paragraph, then a picture in a line of text */
const HEADER =
  '<w:p><w:r><w:drawing><wp:anchor behindDoc="0" simplePos="0" locked="0" layoutInCell="1" allowOverlap="1" relativeHeight="2">' +
  '<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="page"><wp:posOffset>457200</wp:posOffset></wp:positionH>' +
  '<wp:positionV relativeFrom="page"><wp:posOffset>228600</wp:posOffset></wp:positionV>' +
  `${graphic(11, 'rId2').replace('<wp:docPr', '<wp:wrapNone/><wp:docPr')}</wp:anchor></w:drawing></w:r></w:p>` +
  `<w:p><w:r><w:t xml:space="preserve">Company </w:t></w:r><w:r><w:drawing><wp:inline>${graphic(12, 'rId3')}</wp:inline></w:drawing></w:r></w:p>`
const RELS =
  `<Relationship Id="rId1" Type="${IMAGE_REL}" Target="media/image1.png"/>` +
  `<Relationship Id="rId2" Type="${IMAGE_REL}" Target="file:///C:/art/logo.png" TargetMode="External"/>` +
  `<Relationship Id="rId3" Type="${IMAGE_REL}" Target="file:///C:/art/mark.png" TargetMode="External"/>`

async function opened() {
  const bytes = await buildDocx({
    bodyXml: '<w:p><w:r><w:t>body</w:t></w:r></w:p>',
    withImage: true,
    headerXml: HEADER,
    headerRelsXml: RELS,
  })
  const parsed = await parseDocx(bytes)
  const state: HfSectionState = {
    sections: readSections(parsed),
    edits: {},
    links: {},
    parsed: { hfParts: parsed.hfParts },
  }
  return { bytes, parsed, state }
}

type Opened = Awaited<ReturnType<typeof opened>>

async function saved({ parsed }: Opened, state: HfSectionState) {
  const plan = pmDocToSavePlan(blocksToPmDoc(parsed.blocks) as never, parsed.blocks)
  const bytes = await saveDocx(parsed, plan.saveBlocks, hfSaveOptions(state))
  const zip = await JSZip.loadAsync(bytes)
  const xml = await zip.file('word/header1.xml')!.async('string')
  const rels = await zip.file('word/_rels/header1.xml.rels')!.async('string')
  const rel = (id: string | undefined) =>
    id ? (new RegExp(`<Relationship\\b[^>]*Id="${id}"[^>]*>`).exec(rels)?.[0] ?? '') : ''
  const blips = [...xml.matchAll(/<a:blip\b[^>]*>/g)].map((m) => ({
    embed: rel(/r:embed="([^"]+)"/.exec(m[0])?.[1]),
    link: /Target="([^"]*)"/.exec(rel(/r:link="([^"]+)"/.exec(m[0])?.[1]))?.[1] ?? null,
  }))
  const mediaOf = (blip: { embed: string }) =>
    zip.file(`word/${/Target="([^"]+)"/.exec(blip.embed)![1]}`)!.async('base64')
  return { bytes, xml, blips, mediaOf }
}

const apply = (state: HfSectionState, next: Pick<HfSectionState, 'edits' | 'links'> | null) => {
  expect(next).not.toBeNull()
  return { ...state, ...next! }
}

describe('linked pictures of a header in Edit > Links', () => {
  it('lists the floating logo and the picture in the line of text, and an untouched save keeps every byte', async () => {
    const doc = await opened()
    const rows = hfPictureLinkRows(doc.state)
    expect(rows.map((r) => [r.kind, r.variant, r.drawing, r.target])).toEqual([
      ['header', 'default', 0, 'file:///C:/art/logo.png'],
      ['header', 'default', 1, 'file:///C:/art/mark.png'],
    ])
    const link = hfPictureLink(rows[0])
    expect(linkTypeOf(link)).toBe('picture')
    expect(linkSourceOf(link)).toBe('file:///C:/art/logo.png')
    expect((await saved(doc, doc.state)).bytes).toEqual(doc.bytes)
  })

  it('Break Link takes the picture off the list and its link out of the part, text bytes kept', async () => {
    const doc = await opened()
    const [logo, mark] = hfPictureLinkRows(doc.state)
    let state = apply(doc.state, withHfPictureLink(doc.state, logo, { target: null }))
    expect(hfPictureLinkRows(state).map((r) => r.target)).toEqual(['file:///C:/art/mark.png'])
    let out = await saved(doc, state)
    expect(out.blips.map((b) => b.link)).toEqual([null, 'file:///C:/art/mark.png'])
    expect(out.xml).toContain('<w:t xml:space="preserve">Company </w:t>')
    state = apply(state, withHfPictureLink(state, mark, { target: null }))
    expect(hfPictureLinkRows(state)).toEqual([])
    out = await saved(doc, state)
    expect(out.blips.map((b) => b.link)).toEqual([null, null])
    expect(out.xml).not.toContain('aidocs-')
  })

  it('Change Source and Update Now: the new file is named, its bytes are embedded and shown', async () => {
    const doc = await opened()
    const [logo, mark] = hfPictureLinkRows(doc.state)
    const bytes = { base64: OTHER_PNG, mime: 'image/png' as const }
    let state = apply(
      doc.state,
      withHfPictureLink(doc.state, logo, { target: 'file:///D:/new/logo.png', bytes }),
    )
    state = apply(state, withHfPictureLink(state, mark, { target: 'file:///D:/new/mark.png' }))
    expect(hfPictureLinkRows(state).map((r) => r.target)).toEqual([
      'file:///D:/new/logo.png',
      'file:///D:/new/mark.png',
    ])
    // the strip draws the logo from the bytes read; the picture in the line keeps its own
    const shown = resolveHf(state, 0, 'header', 'default')
    expect(shown.images?.map((image) => image.dataUrl)).toEqual([
      `data:image/png;base64,${OTHER_PNG}`,
    ])
    const run = shown.value!.paras!.flatMap((p) => p.runs).find((r) => r.image)!
    expect(run.image!.linkTarget).toBe('file:///D:/new/mark.png')
    expect(run.image!.dataUrl).toBe(`data:image/png;base64,${TINY_PNG_BASE64}`)

    const out = await saved(doc, state)
    expect(out.blips.map((b) => b.link)).toEqual([
      'file:///D:/new/logo.png',
      'file:///D:/new/mark.png',
    ])
    expect(await out.mediaOf(out.blips[0])).toBe(OTHER_PNG)
    expect(await out.mediaOf(out.blips[1])).toBe(TINY_PNG_BASE64)
    expect(out.xml).not.toContain('aidocs-')
    const again = await parseDocx(out.bytes)
    expect(Object.values(again.hfParts ?? {})[0].linkedPictures?.map((p) => p.target)).toEqual([
      'file:///D:/new/logo.png',
      'file:///D:/new/mark.png',
    ])
  })

  it('a header typed into afterwards is written anew with the changed links on its pictures', async () => {
    const doc = await opened()
    const [logo, mark] = hfPictureLinkRows(doc.state)
    const bytes = { base64: OTHER_PNG, mime: 'image/png' as const }
    let state = apply(doc.state, withHfPictureLink(doc.state, logo, { target: null }))
    state = apply(
      state,
      withHfPictureLink(state, mark, { target: 'file:///D:/new/mark.png', bytes }),
    )
    const value = resolveHf(state, 0, 'header', 'default').value!
    const typed = value.paras!.map((p) =>
      p.runs.some((r) => r.text === 'Company ')
        ? { ...p, runs: p.runs.map((r) => (r.text === 'Company ' ? { ...r, text: 'Firm ' } : r)) }
        : p,
    )
    // a commit that hands over the paragraphs alone: the bytes behind the run's placeholder stay
    state = apply(
      state,
      withHfEdit(state, 0, 'header', 'default', { text: value.text, paras: typed }),
    )
    const out = await saved(doc, state)
    expect(out.xml).toContain('Firm ')
    expect(out.blips.map((b) => b.link)).toEqual([null, 'file:///D:/new/mark.png'])
    expect(await out.mediaOf(out.blips[0])).toBe(TINY_PNG_BASE64)
    expect(await out.mediaOf(out.blips[1])).toBe(OTHER_PNG)
    expect(out.xml).not.toContain('aidocs-')
    // the picture deleted from the line is no longer a link of the document
    const without = typed.map((p) => ({ ...p, runs: p.runs.filter((r) => !r.image) }))
    state = apply(state, withHfEdit(state, 0, 'header', 'default', { ...value, paras: without }))
    expect(hfPictureLinkRows(state)).toEqual([])
  })
})
