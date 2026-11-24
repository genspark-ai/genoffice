import { afterEach, describe, expect, it, vi } from 'vitest'
import JSZip from 'jszip'
import { Editor } from '@tiptap/core'
import { NodeSelection } from '@tiptap/pm/state'
import { parseDocx, saveDocx, type HeaderFooter } from '@genoffice/docx-engine'
import {
  buildDocx,
  IMAGE_PARAGRAPH_XML,
  TINY_PNG_BASE64,
} from '../../../packages/docx-engine/tests/helpers/build-docx'
import { setActiveSubEditor } from '../src/renderer/editor/active-editor'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { mountHfEditor } from '../src/renderer/editor/hf-editor'
import {
  canChangePicture,
  changePictureEditor,
  replaceSelectedPictureBytes,
  selectSubEditorPicture,
} from '../src/renderer/editor/picture-bytes'

const DRAWING_RUN = IMAGE_PARAGRAPH_XML.slice('<w:p>'.length, -'</w:p>'.length)
const BODY = `<w:p><w:r><w:t>before</w:t></w:r>${DRAWING_RUN}<w:r><w:t>after</w:t></w:r></w:p>`
const OTHER_PNG =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg=='

afterEach(() => {
  vi.unstubAllGlobals()
  setActiveSubEditor(null)
})

// jsdom loads no images: every new picture reads as 40 x 20
class LoadedImage {
  naturalWidth = 40
  naturalHeight = 20
  onload: (() => void) | null = null
  set src(_url: string) {
    queueMicrotask(() => this.onload?.())
  }
}

describe('Change Picture on a picture in a line of text', () => {
  it('puts the new bytes in the same drawing: the line, the width and the rest of the file stay', async () => {
    const bytes = await buildDocx({ bodyXml: BODY, withImage: true })
    const parsed = await parseDocx(bytes)
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks) as never,
    })
    const save = async () => {
      const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
      const out = await saveDocx(parsed, plan.saveBlocks, {
        ...(plan.newPictures ? { newPictures: plan.newPictures } : {}),
      })
      const zip = await JSZip.loadAsync(out)
      return { out, zip, xml: await zip.file('word/document.xml')!.async('string') }
    }
    expect((await save()).out).toEqual(bytes)

    let at = -1
    editor.state.doc.descendants((node, pos) => {
      if (node.type.name === 'docInlineImage') at = pos
    })
    expect(at).toBeGreaterThan(0)
    expect(canChangePicture(editor)).toBe(false)
    editor.view.dispatch(editor.state.tr.setSelection(NodeSelection.create(editor.state.doc, at)))
    expect(canChangePicture(editor)).toBe(true)
    const before = editor.state.doc.nodeAt(at)!.attrs

    vi.stubGlobal('Image', LoadedImage)
    await replaceSelectedPictureBytes(editor, `data:image/png;base64,${OTHER_PNG}`)
    const after = editor.state.doc.nodeAt(at)!.attrs
    expect(after.dataUrl).toBe(`data:image/png;base64,${OTHER_PNG}`)
    expect(after.widthPx).toBe(before.widthPx)
    expect(after.heightPx).toBe(Math.round(Number(before.widthPx) / 2))
    expect(editor.state.doc.firstChild!.textContent).toBe('beforeafter')

    const { zip, xml } = await save()
    expect(xml).toMatch(/before<\/w:t><\/w:r><w:r>(?:<w:rPr>.*?<\/w:rPr>)?<w:drawing><wp:inline>/)
    expect(xml).not.toContain('aidocs-new-pic-')
    const rId = /<a:blip r:embed="([^"]+)"/.exec(xml)![1]
    expect(rId).not.toBe('rId10')
    const rels = await zip.file('word/_rels/document.xml.rels')!.async('string')
    const target = new RegExp(`Id="${rId}"[^>]*Target="([^"]+)"`).exec(rels)![1]
    expect(await zip.file(`word/${target}`)!.async('base64')).toBe(OTHER_PNG)
    expect(OTHER_PNG).not.toBe(TINY_PNG_BASE64)

    const again = await parseDocx((await save()).out)
    const run = again.blocks[0].runs!.find((r) => r.image)!
    expect(run.image!.dataUrl).toBe(`data:image/png;base64,${OTHER_PNG}`)
    editor.destroy()
  })

  it('acts on the picture selected in the header being edited, and the header saves the new bytes', async () => {
    const bytes = await buildDocx({
      bodyXml: BODY,
      withImage: true,
      headerXml: `<w:p><w:r><w:t xml:space="preserve">Logo </w:t></w:r>${DRAWING_RUN}</w:p>`,
      headerRelsXml:
        '<Relationship Id="rId10" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>',
    })
    const parsed = await parseDocx(bytes)
    const main = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks) as never,
    })
    const value: HeaderFooter = { text: parsed.headerText ?? '', paras: parsed.headerParas! }
    let committed: HeaderFooter | null = null
    const host = document.createElement('div')
    document.body.appendChild(host)
    const handle = mountHfEditor(host, {
      value,
      pageNo: '1',
      pageTotal: '1',
      spellcheck: false,
      onCommit: (next) => (committed = next),
      onExit: () => {},
    })
    const sub = handle.editor
    setActiveSubEditor(sub)
    // nothing selected in the header: Change Picture has nothing to act on
    expect(changePictureEditor(main)).toBeNull()
    const img = sub.view.dom.querySelector('img.doc-inline-img')
    expect(img).not.toBeNull()
    // a right-click on the picture selects it in the header, not in the document
    expect(selectSubEditorPicture(main, img)).toBe(true)
    expect(selectSubEditorPicture(main, main.view.dom.querySelector('img'))).toBe(false)
    expect(sub.state.selection).toBeInstanceOf(NodeSelection)
    expect(changePictureEditor(main)).toBe(sub)

    vi.stubGlobal('Image', LoadedImage)
    await replaceSelectedPictureBytes(sub, `data:image/png;base64,${OTHER_PNG}`)
    handle.exit()
    expect(committed).not.toBeNull()
    const plan = pmDocToSavePlan(main.getJSON() as PmNode, parsed.blocks)
    const zip = await JSZip.loadAsync(
      await saveDocx(parsed, plan.saveBlocks, { header: committed! }),
    )
    const header = await zip.file('word/header1.xml')!.async('string')
    expect(header).toContain('Logo ')
    expect(header.match(/<w:drawing>/g)).toHaveLength(1)
    expect(header).not.toContain('aidocs-new-pic-')
    const rId = /<a:blip r:embed="([^"]+)"/.exec(header)![1]
    const rels = await zip.file('word/_rels/header1.xml.rels')!.async('string')
    const target = new RegExp(`Id="${rId}"[^>]*Target="([^"]+)"`).exec(rels)![1]
    expect(await zip.file(`word/${target}`)!.async('base64')).toBe(OTHER_PNG)
    // the picture of the document was not the one changed
    const body = await zip.file('word/document.xml')!.async('string')
    expect(body).toContain('r:embed="rId10"')
    main.destroy()
    host.remove()
  })

  it('acts on a picture in a text box picked with a right-click while the box is the selected object', async () => {
    const box =
      '<w:p><w:r><mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">' +
      '<mc:Choice Requires="wps"><w:drawing><wp:anchor behindDoc="0" simplePos="0" locked="0" layoutInCell="1" allowOverlap="1">' +
      '<wp:simplePos x="0" y="0"/><wp:extent cx="3000000" cy="1700000"/>' +
      '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
      '<wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
      '<wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="3000000" cy="1700000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr><wps:bodyPr/>' +
      `<wps:txbx><w:txbxContent><w:p><w:r><w:t xml:space="preserve">Boxed </w:t></w:r>${DRAWING_RUN}</w:p></w:txbxContent></wps:txbx>` +
      '</wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice></mc:AlternateContent></w:r></w:p>'
    const bytes = await buildDocx({
      bodyXml: `<w:p><w:r><w:t>first</w:t></w:r></w:p>${box}<w:p><w:r><w:t>last</w:t></w:r></w:p>`,
      withImage: true,
    })
    const parsed = await parseDocx(bytes)
    const host = document.createElement('div')
    document.body.appendChild(host)
    const main = new Editor({
      element: host,
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks) as never,
    })
    const img = main.view.dom.querySelector('.doc-textbox img.doc-inline-img')
    expect(img).not.toBeNull()
    // the box is not being typed in: the picture is found through the editor that holds it
    expect(selectSubEditorPicture(main, img)).toBe(true)
    expect(changePictureEditor(main)).toBeNull()
    let at = -1
    main.state.doc.forEach((node, pos) => {
      if (node.type.name === 'docProtected') at = pos
    })
    main.view.dispatch(main.state.tr.setSelection(NodeSelection.create(main.state.doc, at)))
    const sub = changePictureEditor(main)
    expect(sub).not.toBeNull()
    expect(sub).not.toBe(main)
    // another selection in the document: the box's picture is no longer the one meant
    main.commands.setTextSelection(1)
    expect(changePictureEditor(main)).toBeNull()
    main.view.dispatch(main.state.tr.setSelection(NodeSelection.create(main.state.doc, at)))

    vi.stubGlobal('Image', LoadedImage)
    await replaceSelectedPictureBytes(
      changePictureEditor(main)!,
      `data:image/png;base64,${OTHER_PNG}`,
    )
    const plan = pmDocToSavePlan(main.getJSON() as PmNode, parsed.blocks)
    const zip = await JSZip.loadAsync(
      await saveDocx(parsed, plan.saveBlocks, {
        ...(plan.newPictures ? { newPictures: plan.newPictures } : {}),
      }),
    )
    const xml = await zip.file('word/document.xml')!.async('string')
    const boxXml = /<w:txbxContent>.*?<\/w:txbxContent>/s.exec(xml)![0]
    expect(boxXml).toContain('Boxed ')
    expect(boxXml.match(/<w:drawing>/g)).toHaveLength(1)
    expect(xml).not.toContain('aidocs-new-pic-')
    const rId = /<a:blip r:embed="([^"]+)"/.exec(boxXml)![1]
    expect(rId).not.toBe('rId10')
    const rels = await zip.file('word/_rels/document.xml.rels')!.async('string')
    const target = new RegExp(`Id="${rId}"[^>]*Target="([^"]+)"`).exec(rels)![1]
    expect(await zip.file(`word/${target}`)!.async('base64')).toBe(OTHER_PNG)
    main.destroy()
    host.remove()
  })
})
