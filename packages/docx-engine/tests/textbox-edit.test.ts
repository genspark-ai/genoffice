import { describe, expect, it } from 'vitest'
import {
  patchTextboxHeights,
  patchTextboxParas,
  patchTextboxSizes,
  type TextboxParaPatch,
} from '../src/generate'
import { parseDocx } from '../src/parse'
import { buildDocx } from './helpers/build-docx'

const TXBX_CONTENT =
  '<w:txbxContent>' +
  '<w:p><w:r><w:rPr><w:rFonts w:ascii="Courier New"/><w:color w:val="415461"/><w:sz w:val="22"/></w:rPr>' +
  '<w:t>**Current date:** {Month}</w:t></w:r></w:p>' +
  '<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>CRITICAL</w:t></w:r></w:p>' +
  '</w:txbxContent>'

/** anchored DrawingML code box with the VML fallback twin Word always writes */
const TEXTBOX_PARAGRAPH =
  '<w:p><w:r><mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">' +
  '<mc:Choice Requires="wps"><w:drawing>' +
  '<wp:anchor behindDoc="1" simplePos="0" locked="0" layoutInCell="1" allowOverlap="1">' +
  '<wp:simplePos x="0" y="0"/><wp:extent cx="6038850" cy="1704975"/>' +
  '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
  '<wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
  '<wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="6038850" cy="1704975"/></a:xfrm>' +
  '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>' +
  '<a:solidFill><a:srgbClr val="F5F5F5"/></a:solidFill>' +
  '<a:ln w="9524"><a:solidFill><a:srgbClr val="CCCCCC"/></a:solidFill></a:ln></wps:spPr>' +
  `<wps:txbx>${TXBX_CONTENT}</wps:txbx></wps:wsp>` +
  '</a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice>' +
  '<mc:Fallback><w:pict>' +
  '<v:rect xmlns:v="urn:schemas-microsoft-com:vml" id="box1" style="position:absolute">' +
  `<v:textbox>${TXBX_CONTENT}</v:textbox></v:rect>` +
  '</w:pict></mc:Fallback></mc:AlternateContent></w:r></w:p>'

const para = (patch: TextboxParaPatch): TextboxParaPatch => patch

describe('textbox editing', () => {
  it('extracts the box size from the shape extent', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: TEXTBOX_PARAGRAPH }))
    const box = doc.blocks[0].textboxes?.[0]
    expect(box?.widthPx).toBe(634) // 6038850 EMU / 9525
    // no spAutoFit -> fixed height, Word clips overflow
    expect(box?.heightPx).toBe(179) // 1704975 EMU / 9525
  })

  it('skips the fixed height when the shape auto-fits its text', async () => {
    const autoFitXml = TEXTBOX_PARAGRAPH.replace(
      '<wps:txbx>',
      '<wps:bodyPr><a:spAutoFit/></wps:bodyPr><wps:txbx>',
    )
    const doc = await parseDocx(await buildDocx({ bodyXml: autoFitXml }))
    const box = doc.blocks[0].textboxes?.[0]
    expect(box?.widthPx).toBe(634)
    expect(box?.heightPx).toBeUndefined()
  })

  it('extracts text insets and paragraph geometry used by Word', async () => {
    const formatted = TEXTBOX_PARAGRAPH.replace(
      '<wps:txbx>',
      '<wps:bodyPr lIns="0" tIns="47625" rIns="95250" bIns="0"/><wps:txbx>',
    ).replace(
      '<w:p><w:r><w:rPr><w:rFonts',
      '<w:p><w:pPr><w:spacing w:line="312" w:lineRule="auto" w:before="211" w:after="40"/>' +
        '<w:ind w:left="180" w:right="90" w:firstLine="60"/></w:pPr><w:r><w:rPr><w:rFonts',
    )
    const doc = await parseDocx(await buildDocx({ bodyXml: formatted }))
    const box = doc.blocks[0].textboxes?.[0]

    expect(box).toMatchObject({
      insetLeftPx: 0,
      insetTopPx: 5,
      insetRightPx: 10,
      insetBottomPx: 0,
    })
    expect(box?.paras[0]).toMatchObject({
      lineSpacing: 1.3,
      spaceBefore: 211,
      spaceAfter: 40,
      indentLeft: 180,
      indentRight: 90,
      indentFirstLine: 60,
    })
  })

  it('extracts w:pStyle of textbox paragraphs so document style CSS reaches them', async () => {
    const styled = TEXTBOX_PARAGRAPH.replace(
      '<w:p><w:pPr><w:jc w:val="center"/></w:pPr>',
      '<w:p><w:pPr><w:pStyle w:val="KeyPoint"/><w:jc w:val="center"/></w:pPr>',
    )
    const doc = await parseDocx(await buildDocx({ bodyXml: styled }))
    const box = doc.blocks[0].textboxes?.[0]
    expect(box?.paras[0].styleId).toBeUndefined()
    expect(box?.paras[1].styleId).toBe('KeyPoint')
  })

  it('keeps the anchor paragraph text sharing its paragraph with a textbox as strayRuns', async () => {
    const withHeading = TEXTBOX_PARAGRAPH.replace(
      '<w:p><w:r><mc:AlternateContent',
      '<w:p><w:pPr><w:pStyle w:val="SectionHeading"/></w:pPr><w:r><mc:AlternateContent',
    ).replace(
      '</mc:AlternateContent></w:r></w:p>',
      '</mc:AlternateContent></w:r><w:r><w:t>Objective</w:t></w:r></w:p>',
    )
    const doc = await parseDocx(await buildDocx({ bodyXml: withHeading }))
    const block = doc.blocks[0]
    expect(block.type).toBe('passthrough')
    expect(block.textboxes).toHaveLength(1)
    expect(block.strayRuns?.map((r) => r.text).join('')).toBe('Objective')
    expect(block.strayStyleId).toBe('SectionHeading')
    expect(block.previewText).toContain('Objective')
    // untouched block still saves its original bytes
    expect(block.originalXml).toBe(withHeading)
  })

  it('pins the height of an auto-fit box: drops spAutoFit and round-trips fixed', async () => {
    const autoFitXml = TEXTBOX_PARAGRAPH.replace(
      '<wps:txbx>',
      '<wps:bodyPr><a:spAutoFit/></wps:bodyPr><wps:txbx>',
    )
    const out = patchTextboxSizes(autoFitXml, [{ wPx: null, hPx: 240 }])
    expect(out).toContain('<wp:extent cx="6038850" cy="2286000"/>')
    expect(out).toContain('<a:ext cx="6038850" cy="2286000"/>')
    expect(out).not.toContain('<a:spAutoFit')
    expect(out).toContain('<a:noAutofit/>')

    // reopening now yields a fixed-height box instead of autofit
    const doc = await parseDocx(await buildDocx({ bodyXml: out }))
    const box = doc.blocks[0].textboxes?.[0]
    expect(box?.heightPx).toBe(240)
  })

  it('updates DrawingML and VML fallback heights without changing box styling', () => {
    const withFallbackHeight = TEXTBOX_PARAGRAPH.replace(
      'style="position:absolute"',
      'style="position:absolute;width:475.5pt;height:134.25pt"',
    )
    const out = patchTextboxHeights(withFallbackHeight, [240])

    expect(out).toContain('<wp:extent cx="6038850" cy="2286000"/>')
    expect(out).toContain('<a:ext cx="6038850" cy="2286000"/>')
    expect(out).toContain('width:475.5pt;height:180pt')
    expect(out).toContain('<a:srgbClr val="F5F5F5"/>')
  })

  it('regenerates changed paragraphs from rich runs and keeps unchanged bytes', () => {
    const out = patchTextboxParas(TEXTBOX_PARAGRAPH, [
      [
        para({
          runs: [
            {
              text: '**Current date:** July 21',
              font: 'Courier New',
              color: '415461',
              sizeHalfPoints: 22,
            },
          ],
        }),
        null, // untouched
      ],
    ])
    expect(out).toContain('<w:t xml:space="preserve">**Current date:** July 21</w:t>')
    expect(out).not.toContain('{Month}')
    // rich run styles regenerated from the Run model
    expect(out).toContain('w:ascii="Courier New"')
    expect(out).toContain('<w:color w:val="415461"/>')
    expect(out).toContain('<w:sz w:val="22"/>')
    // untouched paragraph keeps its original bytes
    expect(out).toContain(
      '<w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>CRITICAL</w:t></w:r>',
    )
    // box shape untouched
    expect(out).toContain('<a:srgbClr val="F5F5F5"/>')
    expect(out).toContain('<a:srgbClr val="CCCCCC"/>')
  })

  it('applies new run formatting (color, bold) to edited paragraphs', () => {
    const out = patchTextboxParas(TEXTBOX_PARAGRAPH, [
      [
        para({
          runs: [
            { text: 'red part', color: 'FF0000' },
            { text: ' bold part', bold: true, sizeHalfPoints: 28 },
          ],
        }),
        null,
      ],
    ])
    expect(out).toContain(
      '<w:r><w:rPr><w:color w:val="FF0000"/></w:rPr><w:t xml:space="preserve">red part</w:t></w:r>',
    )
    expect(out).toContain(
      '<w:r><w:rPr><w:b/><w:bCs/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr><w:t xml:space="preserve"> bold part</w:t></w:r>',
    )
  })

  it('rewrites, keeps, or removes w:jc according to align', () => {
    const centered = patchTextboxParas(TEXTBOX_PARAGRAPH, [
      [para({ runs: [{ text: 'Hi' }], align: 'center' }), null],
    ])
    expect(centered).toContain(
      '<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t xml:space="preserve">Hi</w:t></w:r></w:p>',
    )

    // align undefined keeps the template pPr (second para is centered)
    const kept = patchTextboxParas(TEXTBOX_PARAGRAPH, [
      [null, para({ runs: [{ text: 'STILL', bold: true }] })],
    ])
    expect(kept).toContain(
      '<w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/><w:bCs/></w:rPr><w:t xml:space="preserve">STILL</w:t></w:r>',
    )

    // align null strips the centered jc from the second paragraph
    const removed = patchTextboxParas(TEXTBOX_PARAGRAPH, [
      [null, para({ runs: [{ text: 'LEFT', bold: true }], align: null })],
    ])
    expect(removed).toContain(
      '<w:p><w:r><w:rPr><w:b/><w:bCs/></w:rPr><w:t xml:space="preserve">LEFT</w:t></w:r></w:p>',
    )
  })

  it('patches the VML fallback twin with the same paragraphs', () => {
    const out = patchTextboxParas(TEXTBOX_PARAGRAPH, [[para({ runs: [{ text: 'NEW' }] }), null]])
    const hits = out.match(/<w:t xml:space="preserve">NEW<\/w:t>/g) ?? []
    expect(hits).toHaveLength(2) // mc:Choice + mc:Fallback
    expect(out).not.toContain('{Month}')
  })

  it('added paragraphs inherit the last original pPr; removed ones drop', () => {
    const out = patchTextboxParas(TEXTBOX_PARAGRAPH, [
      [null, null, para({ runs: [{ text: 'line3', bold: true }] }), para({ runs: [] })],
    ])
    // new paragraph reuses the centered pPr of the last original
    expect(out).toContain(
      '<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/><w:bCs/></w:rPr><w:t xml:space="preserve">line3</w:t></w:r></w:p>',
    )
    // empty paragraph keeps pPr only
    expect(out).toContain('<w:p><w:pPr><w:jc w:val="center"/></w:pPr></w:p>')
    const fewer = patchTextboxParas(TEXTBOX_PARAGRAPH, [[para({ runs: [{ text: 'only line' }] })]])
    expect(fewer).not.toContain('CRITICAL')
  })

  it('escapes special characters and converts tabs/newlines to OOXML elements', () => {
    const out = patchTextboxParas(TEXTBOX_PARAGRAPH, [
      [para({ runs: [{ text: 'a < b & c\tD\nE' }] }), null],
    ])
    expect(out).toContain(
      '<w:t xml:space="preserve">a &lt; b &amp; c</w:t><w:tab/><w:t xml:space="preserve">D</w:t><w:br/><w:t xml:space="preserve">E</w:t>',
    )
  })

  it('round-trips: patched XML re-parses with the new rich formatting', async () => {
    const patched = patchTextboxParas(TEXTBOX_PARAGRAPH, [
      [
        para({
          runs: [{ text: '**Current date:** July 21', font: 'Courier New', color: 'FF0000' }],
        }),
        para({ runs: [{ text: 'IMPORTANT', bold: true, color: '0070C0' }], align: 'center' }),
      ],
    ])
    const doc = await parseDocx(await buildDocx({ bodyXml: patched }))
    const box = doc.blocks[0].textboxes?.[0]
    expect(box?.fill).toBe('F5F5F5')
    expect(box?.borderColor).toBe('CCCCCC')
    expect(box?.paras.map((p) => p.runs.map((r) => r.text).join(''))).toEqual([
      '**Current date:** July 21',
      'IMPORTANT',
    ])
    expect(box?.paras[0].runs[0]).toMatchObject({ font: 'Courier New', color: 'FF0000' })
    expect(box?.paras[1].runs[0]).toMatchObject({ bold: true, color: '0070C0' })
    expect(box?.paras[1].align).toBe('center')
  })

  it('leaves the paragraph untouched when every box is null', () => {
    expect(patchTextboxParas(TEXTBOX_PARAGRAPH, [null])).toBe(TEXTBOX_PARAGRAPH)
    expect(patchTextboxParas('<w:p><w:r><w:t>plain</w:t></w:r></w:p>', [])).toBe(
      '<w:p><w:r><w:t>plain</w:t></w:r></w:p>',
    )
  })
})
