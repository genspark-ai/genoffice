/** Linked pictures in a line of text, a table cell and a text box: listed, relinked, updated, broken. */
import { afterEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { parseDocx, saveDocx } from '@genoffice/docx-engine'
import JSZip from 'jszip'
import { buildDocx, TINY_PNG_BASE64 } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { blocksToPmDoc, pmDocOptions, pmDocToSavePlan } from '../src/renderer/editor/convert'
import { insertPictureAtSelection } from '../src/renderer/editor/insert-picture'
import {
  breakLinks,
  collectLinks,
  hasLinks,
  setLinkSource,
  updateLink,
} from '../src/renderer/editor/link-fields'

const IMAGE_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'
const linkRel = (id: string, target: string) =>
  `<Relationship Id="${id}" Type="${IMAGE_REL}" Target="${target}" TargetMode="External"/>`

/** an inline picture as Word writes one that is embedded and linked: r:embed first */
const picture = (linkId: string) =>
  '<w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="914400"/><wp:docPr id="7" name="Picture 7"/>' +
  '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
  `<pic:pic><pic:blipFill><a:blip r:embed="rId10" r:link="${linkId}"/></pic:blipFill>` +
  '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm></pic:spPr></pic:pic>' +
  '</a:graphicData></a:graphic></wp:inline></w:drawing></w:r>'
const text = (s: string) => `<w:r><w:t xml:space="preserve">${s}</w:t></w:r>`
const cell = (inner: string) =>
  `<w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>${inner}</w:tc>`
const TABLE =
  '<w:tbl><w:tblPr><w:tblW w:w="8000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>' +
  `<w:tr>${cell(`<w:p>${text('Logo ')}${picture('rId31')}</w:p>`)}${cell('<w:p/>')}</w:tr></w:tbl>`
const TEXTBOX =
  '<w:p><w:r><mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">' +
  '<mc:Choice Requires="wps"><w:drawing>' +
  '<wp:anchor behindDoc="0" simplePos="0" locked="0" layoutInCell="1" allowOverlap="1">' +
  '<wp:simplePos x="0" y="0"/><wp:extent cx="3000000" cy="1700000"/>' +
  '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
  '<wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
  '<wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="3000000" cy="1700000"/></a:xfrm></wps:spPr>' +
  '<wps:bodyPr/>' +
  `<wps:txbx><w:txbxContent><w:p>${text('In a box ')}${picture('rId32')}</w:p></w:txbxContent></wps:txbx></wps:wsp>` +
  '</a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice></mc:AlternateContent></w:r></w:p>'

const live: Editor[] = []
afterEach(() => {
  for (const editor of live.splice(0)) editor.destroy()
  delete (globalThis as { window?: { desktop?: unknown } }).window?.desktop
})

async function opened(bodyXml: string, extraRels = '') {
  const bytes = await buildDocx({ bodyXml, withImage: true, extraRels })
  const parsed = await parseDocx(bytes)
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks, undefined, pmDocOptions(parsed)),
  })
  live.push(editor)
  return { parsed, editor }
}

async function saved(editor: Editor, parsed: Awaited<ReturnType<typeof parseDocx>>) {
  const plan = pmDocToSavePlan(editor.getJSON() as never, parsed.blocks)
  const bytes = await saveDocx(parsed, plan.saveBlocks, { newPictures: plan.newPictures })
  const zip = await JSZip.loadAsync(bytes)
  return {
    xml: await zip.file('word/document.xml')!.async('string'),
    rels: await zip.file('word/_rels/document.xml.rels')!.async('string'),
    bytes,
  }
}

/** the external targets the blips of the document are linked to, in document order */
function linkedTargets({ xml, rels }: { xml: string; rels: string }): string[] {
  return [...xml.matchAll(/<a:blip\b[^>]*\br:link="([^"]+)"/g)].map((m) => {
    const rel = new RegExp(`<Relationship\\b[^>]*Id="${m[1]}"[^>]*>`).exec(rels)?.[0] ?? ''
    expect(rel).toContain('TargetMode="External"')
    return /Target="([^"]*)"/.exec(rel)?.[1] ?? ''
  })
}

const BODY = `<w:p>${text('Seal: ')}${picture('rId30')}${text(' end')}</w:p>${TABLE}${TEXTBOX}`
const RELS =
  linkRel('rId30', 'file:///pics/seal.png') +
  linkRel('rId31', 'file:///pics/logo.png') +
  linkRel('rId32', 'file:///pics/boxed.png')

describe('linked pictures outside the picture blocks', () => {
  it('are listed where they stand, showing their own bytes, and an untouched document saves its links as they came', async () => {
    const { parsed, editor } = await opened(BODY, RELS)
    const links = collectLinks(editor.state.doc, parsed.blocks)
    expect(links.map((l) => (l.type === 'picture' ? [l.target, l.at?.in] : null))).toEqual([
      ['file:///pics/seal.png', 'line'],
      ['file:///pics/logo.png', 'line'],
      ['file:///pics/boxed.png', 'textbox'],
    ])
    expect(hasLinks(editor.state.doc, parsed.blocks)).toBe(true)
    // embedded and linked: what is drawn is the embedded picture, never the file the link names
    const shown: string[] = []
    editor.state.doc.descendants((node) => {
      if (node.type.name === 'docInlineImage') shown.push(String(node.attrs.dataUrl).slice(0, 15))
    })
    expect(shown).toEqual(['data:image/png;', 'data:image/png;'])
    const out = await saved(editor, parsed)
    expect(linkedTargets(out)).toEqual([
      'file:///pics/seal.png',
      'file:///pics/logo.png',
      'file:///pics/boxed.png',
    ])
    expect(out.xml).toContain('r:link="rId30"')
  })

  it('Break Link leaves the picture without its link; the others keep theirs', async () => {
    const { parsed, editor } = await opened(BODY, RELS)
    const links = () => collectLinks(editor.state.doc, parsed.blocks)
    // the one in the cell and the one in the text box
    expect(breakLinks(editor, [links()[1], links()[2]], parsed.blocks)).toBe(2)
    expect(links().map((l) => l.type === 'picture' && l.target)).toEqual(['file:///pics/seal.png'])
    const out = await saved(editor, parsed)
    expect(linkedTargets(out)).toEqual(['file:///pics/seal.png'])
    expect(out.xml.match(/r:embed="rId10"/g)).toHaveLength(3)
    const again = await parseDocx(out.bytes)
    const reopened = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(again.blocks, undefined, pmDocOptions(again)),
    })
    live.push(reopened)
    expect(collectLinks(reopened.state.doc, again.blocks)).toHaveLength(1)
    editor.commands.undo()
    expect(links()).toHaveLength(3)
  })

  it('Change Source names the new file in a relationship of its own', async () => {
    const { parsed, editor } = await opened(BODY, RELS)
    const links = () => collectLinks(editor.state.doc, parsed.blocks)
    for (const i of [1, 2])
      expect(
        setLinkSource(editor, links()[i], parsed.blocks, {
          path: `/new/p${i}.png`,
          url: `file:///new/p${i}.png`,
        }),
      ).toBe(true)
    expect(links().map((l) => l.type === 'picture' && l.target)).toEqual([
      'file:///pics/seal.png',
      'file:///new/p1.png',
      'file:///new/p2.png',
    ])
    const out = await saved(editor, parsed)
    expect(linkedTargets(out)).toEqual([
      'file:///pics/seal.png',
      'file:///new/p1.png',
      'file:///new/p2.png',
    ])
    expect(out.xml).not.toContain('aidocs-link:')
  })

  it('Update Now takes the bytes of the file and keeps the link', async () => {
    const { parsed, editor } = await opened(BODY, RELS)
    // another picture than the one embedded: two pixels wide
    const fresh =
      'iVBORw0KGgoAAAANSUhEUgAAAAIAAAABCAYAAAD0In+KAAAAEElEQVR42mNkYGD4z8DAwAAABQABAfmB5wAAAABJRU5ErkJggg=='
    ;(globalThis as { window: { desktop?: unknown } }).window.desktop = {
      readLinkSource: async () => ({ ok: true, kind: 'picture', mime: 'image/png', base64: fresh }),
    }
    // jsdom loads no pictures: the size of the new one is given
    class LoadedImage {
      naturalWidth = 2
      naturalHeight = 1
      width = 2
      height = 1
      onload: (() => void) | null = null
      set src(_: string) {
        queueMicrotask(() => this.onload?.())
      }
    }
    const realImage = globalThis.Image
    globalThis.Image = LoadedImage as unknown as typeof Image
    const links = () => collectLinks(editor.state.doc, parsed.blocks)
    const results = []
    for (const i of [2, 1])
      results.push(await updateLink(editor, links()[i], { list: () => [], add: () => {} }))
    globalThis.Image = realImage
    expect(results.map((r) => r.status)).toEqual(['updated', 'updated'])
    expect(links()).toHaveLength(3)
    const out = await saved(editor, parsed)
    expect(linkedTargets(out)).toEqual([
      'file:///pics/seal.png',
      'file:///pics/logo.png',
      'file:///pics/boxed.png',
    ])
    // the two updated pictures point at new media, the untouched one at the part it had
    const embeds = [...out.xml.matchAll(/<a:blip\b[^>]*\br:embed="([^"]+)"/g)].map((m) => m[1])
    expect(embeds[0]).toBe('rId10')
    expect(embeds[1]).not.toBe('rId10')
    expect(embeds[2]).not.toBe('rId10')
    expect(out.xml).not.toContain('aidocs-')
  })

  it('a picture inserted with Link to File into a table cell keeps the link through a save', async () => {
    const { parsed, editor } = await opened(
      `<w:p>${text('Body')}</w:p>${TABLE}`,
      linkRel('rId31', 'file:///pics/logo.png'),
    )
    let empty = -1
    editor.state.doc.descendants((node, pos) => {
      if (node.type.name === 'docParagraph' && node.content.size === 0 && empty < 0) empty = pos
    })
    editor.commands.setTextSelection(empty + 1)
    insertPictureAtSelection(editor, {
      dataUrl: `data:image/png;base64,${TINY_PNG_BASE64}`,
      base64: TINY_PNG_BASE64,
      mime: 'image/png',
      widthPx: 40,
      heightPx: 40,
      label: 'new.png',
      linkTarget: 'file:///pics/new.png',
    })
    const targets = () =>
      collectLinks(editor.state.doc, parsed.blocks).map((l) => l.type === 'picture' && l.target)
    expect(targets()).toEqual(['file:///pics/logo.png', 'file:///pics/new.png'])
    const out = await saved(editor, parsed)
    expect(linkedTargets(out)).toEqual(['file:///pics/logo.png', 'file:///pics/new.png'])
    expect(out.xml).not.toContain('aidocs-')
  })
})
