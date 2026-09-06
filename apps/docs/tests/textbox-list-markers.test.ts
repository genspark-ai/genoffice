import { describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { parseDocx } from '@genoffice/docx-engine'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'

// numId 1: Symbol bullet with a hanging indent; numId 2: decimal level without any w:ind
const NUMBERING =
  XML_DECL +
  '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
  '<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/>' +
  '<w:lvlText w:val="&#xF0B7;"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>' +
  '<w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol"/></w:rPr></w:lvl></w:abstractNum>' +
  '<w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/>' +
  '<w:lvlText w:val="%1."/></w:lvl></w:abstractNum>' +
  '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>' +
  '<w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>' +
  '</w:numbering>'

const li = (numId: string, text: string) =>
  `<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="${numId}"/></w:numPr></w:pPr>` +
  (text ? `<w:r><w:t>${text}</w:t></w:r>` : '') +
  '</w:p>'

const TXBX_CONTENT =
  '<w:txbxContent>' +
  li('1', 'hanging bullet') +
  li('1', '') +
  li('2', 'no level indent') +
  '<w:p/>' +
  '</w:txbxContent>'

const TEXTBOX_PARAGRAPH =
  '<w:p><w:r><mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">' +
  '<mc:Choice Requires="wps"><w:drawing>' +
  '<wp:anchor behindDoc="1" simplePos="0" locked="0" layoutInCell="1" allowOverlap="1">' +
  '<wp:simplePos x="0" y="0"/><wp:extent cx="6038850" cy="1704975"/>' +
  '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
  '<wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
  '<wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="6038850" cy="1704975"/></a:xfrm></wps:spPr>' +
  '<wps:bodyPr lIns="0" tIns="0" rIns="0" bIns="0"/>' +
  `<wps:txbx>${TXBX_CONTENT}</wps:txbx></wps:wsp>` +
  '</a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice>' +
  '<mc:Fallback><w:pict>' +
  '<v:rect xmlns:v="urn:schemas-microsoft-com:vml" id="box1" style="position:absolute">' +
  `<v:textbox>${TXBX_CONTENT}</v:textbox></v:rect>` +
  '</w:pict></mc:Fallback></mc:AlternateContent></w:r></w:p>'

describe('list markers rendered inside a textbox', () => {
  it('draws the marker with the level geometry and keeps empty items on a line', async () => {
    const parsed = await parseDocx(
      await buildDocx({
        bodyXml: `<w:p><w:r><w:t>Leading text</w:t></w:r></w:p>${TEXTBOX_PARAGRAPH}`,
        numberingXml: NUMBERING,
      }),
    )
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
    })
    editor.commands.setContent(blocksToPmDoc(parsed.blocks) as never)
    const paras = Array.from(
      editor.view.dom.querySelectorAll<HTMLElement>('.doc-textbox .doc-textbox-para'),
    )
    expect(paras.map((p) => p.getAttribute('data-marker'))).toEqual(['•', '•', '1.', null])

    const [bullet, emptyBullet, plainNumber, emptyPlain] = paras
    expect(bullet.style.getPropertyValue('margin-inline-start')).toBe('36pt')
    expect(bullet.style.getPropertyValue('--li-hang')).toBe('18pt')
    expect(bullet.style.textIndent).toBe('0pt')
    expect(bullet.style.getPropertyValue('--li-tab')).toBe('')

    // an empty list paragraph still shows its marker on a full line
    expect(emptyBullet.classList.contains('doc-textbox-para-empty')).toBe(false)
    expect(emptyPlain.classList.contains('doc-textbox-para-empty')).toBe(true)

    // no w:ind anywhere: nothing hangs, the marker runs to Word's default tab stop
    expect(plainNumber.style.getPropertyValue('--li-hang')).toBe('0pt')
    expect(plainNumber.style.getPropertyValue('--li-tab')).toBe('36pt')
    expect(plainNumber.style.getPropertyValue('margin-inline-start')).toBe('')
    editor.destroy()
  })
})
