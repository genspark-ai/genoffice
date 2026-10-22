import { describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { parseDocx } from '@genoffice/docx-engine'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc } from '../src/renderer/editor/convert'
import { effectivePPrCss, effectiveRPrCss } from '../src/renderer/editor/effective-css'
import { editorExtensions } from '../src/renderer/editor/extensions'

describe('effectivePPrCss', () => {
  it('turns resolved twips into spacing and indent declarations', () => {
    expect(
      effectivePPrCss({
        spaceBeforeTwips: 120,
        spaceAfterTwips: 0,
        indentLeftTwips: 0,
        indentRightTwips: 240,
        indentFirstLineTwips: -200,
      }),
    ).toEqual([
      'margin-top:6.0pt',
      'margin-bottom:0.0pt',
      'margin-inline-start:0pt',
      'margin-inline-end:12pt',
      'text-indent:-10pt',
    ])
    expect(effectivePPrCss({ spaceBeforeTwips: 211 }, { snap: false })).toEqual([
      'margin-top:10.55pt',
    ])
    expect(
      effectivePPrCss({ indentLeftTwips: 300, indentFirstLineTwips: 420 }, { margins: false }),
    ).toEqual(['text-indent:21pt'])
  })

  it('reads the compat flag from the model: padding instead of margin, auto spacing 0', () => {
    expect(effectivePPrCss({ spaceBeforeAuto: true, spaceAfterAuto: true })).toEqual([
      'margin-top:14.0pt',
      'margin-bottom:14.0pt',
    ])
    expect(
      effectivePPrCss({ spaceBeforeAuto: true, spaceAfterTwips: 100, additiveSpacing: true }),
    ).toEqual(['padding-top:0.0pt', 'margin-bottom:5.0pt'])
  })

  it('emits size, face and weight of a resolved run', () => {
    expect(effectiveRPrCss({ sizeHalfPoints: 28, fontFamily: 'Arial', bold: true })).toEqual([
      'font-size:14pt',
      expect.stringMatching(/^font-family:/),
      'font-weight:bold',
    ])
    expect(effectiveRPrCss({})).toEqual([])
  })
})

const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
const STYLES =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ${NS}>` +
  '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Side"><w:name w:val="Side"/>' +
  '<w:pPr><w:ind w:left="720"/></w:pPr></w:style></w:styles>'

const TXBX = (paras: string) =>
  '<w:p><w:r><mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">' +
  '<mc:Choice Requires="wps"><w:drawing>' +
  '<wp:anchor behindDoc="0" simplePos="0" locked="0" layoutInCell="1" allowOverlap="1">' +
  '<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>' +
  '<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>' +
  '<wp:extent cx="2000000" cy="1000000"/><wp:wrapSquare wrapText="bothSides"/>' +
  '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
  '<wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
  '<wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2000000" cy="1000000"/></a:xfrm></wps:spPr>' +
  `<wps:txbx><w:txbxContent>${paras}</w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp>` +
  '</a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice></mc:AlternateContent></w:r></w:p>'

describe('textbox paragraph indents', () => {
  it('a direct w:left=0 cancels the style indent; a silent paragraph keeps it through CSS', async () => {
    const parsed = await parseDocx(
      await buildDocx({
        bodyXml: TXBX(
          '<w:p><w:pPr><w:pStyle w:val="Side"/><w:ind w:left="0"/></w:pPr><w:r><w:t>Flush</w:t></w:r></w:p>' +
            '<w:p><w:pPr><w:pStyle w:val="Side"/></w:pPr><w:r><w:t>Indented</w:t></w:r></w:p>',
        ),
        stylesXml: STYLES,
      }),
    )
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
    })
    editor.commands.setContent(blocksToPmDoc(parsed.blocks) as never)
    const [flush, indented] = Array.from(
      editor.view.dom.querySelectorAll<HTMLElement>('.doc-textbox .doc-textbox-para'),
    )
    expect(flush.getAttribute('data-style')).toBe('Side')
    expect(flush.style.getPropertyValue('margin-inline-start')).toBe('0pt')
    expect(indented.style.getPropertyValue('margin-inline-start')).toBe('')
    editor.destroy()
  })
})
