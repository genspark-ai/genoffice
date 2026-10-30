import { describe, expect, it } from 'vitest'
import { ommlToLatex, ommlToMathML, parseDocx } from '../src/index'
import { ommlEndsWithBreak } from '../src/math'
import { buildDocx } from './helpers/build-docx'

const INLINE_PICTURE =
  '<w:r><w:drawing><wp:inline><wp:extent cx="5669280" cy="3166353"/>' +
  '<wp:docPr id="3" name="Picture 3"/>' +
  '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
  '<pic:pic><pic:nvPicPr><pic:cNvPr id="3" name="Picture 3"/><pic:cNvPicPr/></pic:nvPicPr>' +
  '<pic:blipFill><a:blip r:embed="rId10"/></pic:blipFill></pic:pic>' +
  '</a:graphicData></a:graphic></wp:inline></w:drawing></w:r>'

const CAPTION_RUNS =
  '<w:r><w:t xml:space="preserve">Figure </w:t></w:r>' +
  '<w:fldSimple w:instr=" SEQ Figure \\* ARABIC "><w:r><w:t>3</w:t></w:r></w:fldSimple>' +
  '<w:r><w:t>: Three pillars</w:t></w:r>'

describe('in-line picture sharing its paragraph with a caption field', () => {
  it('keeps the picture and carries the caption as the field display', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml: `<w:p><w:pPr><w:jc w:val="center"/></w:pPr>${INLINE_PICTURE}${CAPTION_RUNS}</w:p>`,
        withImage: true,
      }),
    )
    const block = doc.blocks[0]
    expect(block.type).toBe('image')
    expect(block.imageDataUrl).toMatch(/^data:image\/png;base64,/)
    expect(block.fieldDisplay?.kind).toBe('text')
    expect(block.fieldDisplay?.left).toBe('Figure 3: Three pillars')
  })

  it('leaves text before the picture on the field passthrough', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml: `<w:p>${CAPTION_RUNS}${INLINE_PICTURE}</w:p>`,
        withImage: true,
      }),
    )
    expect(doc.blocks[0].type).toBe('passthrough')
    expect(doc.blocks[0].previewText).toContain('Figure 3')
  })
})

const mathRun = (text: string, marker = '') =>
  `<m:r><w:rPr><w:rFonts w:ascii="Cambria Math"/></w:rPr>${marker}<m:t>${text}</m:t></m:r>`

const equation = (marker: string) =>
  `<m:oMathPara><m:oMath>${mathRun('x', marker)}${mathRun('=')}${mathRun('1')}</m:oMath></m:oMathPara>`

describe('equations inside table cells', () => {
  it('reach the cell runs as atomic math runs', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml:
          '<w:tbl><w:tblGrid><w:gridCol w:w="8000"/></w:tblGrid><w:tr><w:tc>' +
          `<w:p><w:pPr><w:jc w:val="center"/></w:pPr>${equation('<w:lastRenderedPageBreak/>')}</w:p>` +
          '</w:tc></w:tr></w:tbl>',
      }),
    )
    const runs = doc.blocks[0].table!.rows[0][0].richParas![0].runs
    expect(runs).toHaveLength(1)
    expect(runs[0].text).toBe('x=1')
    expect(runs[0].math?.omml).toContain('<m:oMath>')
    expect(runs[0].math?.display).toBe(true)
  })

  it('skip m:oMath copies inside mc:Fallback so the visible equation binds its own OMML', async () => {
    const shapeWithFallbackMath =
      '<w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><wp:inline>' +
      '<wp:extent cx="914400" cy="914400"/><wp:docPr id="7" name="Shape 7"/>' +
      '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
      '<wps:wsp><wps:spPr/></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></mc:Choice>' +
      `<mc:Fallback><w:pict><v:shape><m:oMath>${mathRun('y')}${mathRun('=')}${mathRun('0')}</m:oMath></v:shape></w:pict></mc:Fallback>` +
      '</mc:AlternateContent></w:r>'
    const doc = await parseDocx(
      await buildDocx({
        bodyXml:
          '<w:tbl><w:tblGrid><w:gridCol w:w="8000"/></w:tblGrid><w:tr><w:tc>' +
          `<w:p>${shapeWithFallbackMath}${equation('')}</w:p>` +
          '</w:tc></w:tr></w:tbl>',
      }),
    )
    const runs = doc.blocks[0].table!.rows[0][0].richParas![0].runs
    const math = runs.filter((r) => r.math)
    expect(math).toHaveLength(1)
    expect(math[0].text).toBe('x=1')
    expect(math[0].math?.omml).not.toContain('<m:t>y</m:t>')
  })

  it('markers inside m:r do not disturb the conversion', () => {
    const omml =
      `<m:oMath>${mathRun('x', '<w:lastRenderedPageBreak/><w:bookmarkStart w:id="0" w:name="_b"/>')}` +
      `${mathRun('=', '<w:bookmarkEnd w:id="0"/>')}${mathRun('1')}</m:oMath>`
    expect(ommlToMathML(omml)).toContain('<mi>x</mi><mo>=</mo><mn>1</mn>')
    expect(ommlToLatex(omml)).toBe('x=1')
  })
})

describe('manual line breaks inside a math zone', () => {
  const brRun = '<m:r><m:rPr><m:sty m:val="p"/></m:rPr><w:br/></m:r>'

  it('ommlEndsWithBreak sees only a trailing break run', () => {
    expect(ommlEndsWithBreak(`<m:oMath>${mathRun('x')}${brRun}</m:oMath>`)).toBe(true)
    expect(ommlEndsWithBreak(`<m:oMath>${mathRun('x')}</m:oMath>`)).toBe(false)
    expect(ommlEndsWithBreak(`<m:oMath>${brRun}${mathRun('x')}</m:oMath>`)).toBe(false)
  })

  it('flags the fragment before the break so the next m:oMath starts a new line', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml:
          '<w:tbl><w:tblGrid><w:gridCol w:w="8000"/></w:tblGrid><w:tr><w:tc>' +
          `<w:p><m:oMathPara><m:oMath>${mathRun('a')}${brRun}</m:oMath>` +
          `<m:oMath>${mathRun('b')}</m:oMath></m:oMathPara></w:p>` +
          '</w:tc></w:tr></w:tbl>',
      }),
    )
    const runs = doc.blocks[0].table!.rows[0][0].richParas![0].runs
    expect(runs.map((r) => r.math?.breakAfter ?? false)).toEqual([true, false])
    expect(runs[0].math?.omml).toContain('<w:br/>')
  })
})
