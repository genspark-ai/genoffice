import { describe, expect, it } from 'vitest'
import { parseDocx } from '../src/index'
import { buildDocx, buildKitchenSinkDocx } from './helpers/build-docx'

describe('parseDocx', () => {
  it('parses the kitchen-sink document into anchored blocks', async () => {
    const bytes = await buildKitchenSinkDocx()
    const doc = await parseDocx(bytes)

    const visible = doc.blocks.filter((b) => !b.hidden)
    expect(visible.map((b) => b.type)).toEqual([
      'heading',
      'paragraph',
      'heading',
      'listItem',
      'listItem',
      'listItem',
      'paragraph',
      'table',
      'image',
      'passthrough',
      'paragraph',
    ])

    // every block anchored to its top-level element index with exact XML slice
    doc.blocks.forEach((block, i) => {
      expect(block.docxIndex).toBe(i)
      expect(block.originalXml).toBeTruthy()
      expect(doc.internal.documentXml).toContain(block.originalXml!)
    })

    const h1 = visible[0]
    expect(h1.level).toBe(1)
    expect(h1.runs?.[0].text).toBe('第一章 概述')

    const rich = visible[1]
    const runs = rich.runs!
    expect(runs.find((r) => r.bold)?.text).toBe('加粗')
    expect(runs.find((r) => r.italic)?.text).toBe('斜体')
    expect(runs.find((r) => r.underline)?.text).toBe('下划线')
    expect(runs.find((r) => r.strike)?.text).toBe('删除线')
    const colored = runs.find((r) => r.color === 'FF0000')
    expect(colored?.text).toBe('红色14号')
    expect(colored?.sizeHalfPoints).toBe(28)
    // XML entities decoded
    expect(runs[runs.length - 1].text).toContain('& < >')

    const bullets = visible.filter((b) => b.type === 'listItem')
    expect(bullets[0].list).toEqual({ kind: 'bullet', numId: '1', ilvl: 0 })
    expect(bullets[2].list).toEqual({ kind: 'ordered', numId: '2', ilvl: 0 })

    const linkPara = visible[6]
    const linkRun = linkPara.runs!.find((r) => r.link)
    expect(linkRun?.link?.href).toBe('https://example.com/')
    expect(linkRun?.link?.rId).toBe('rId20')

    const table = visible[7]
    expect(table.label).toBe('Table 2×2')
    expect(table.previewText).toContain('A1')

    const image = visible[8]
    expect(image.imageDataUrl).toMatch(/^data:image\/png;base64,/)

    const math = visible[9]
    expect(math.label).toBe('Equation')

    // trailing sectPr is hidden
    const hidden = doc.blocks.filter((b) => b.hidden)
    expect(hidden).toHaveLength(1)
    expect(hidden[0].originalXml).toContain('<w:sectPr>')

    // style metadata for generation
    expect(doc.headingStyleIds.get(1)).toBe('Heading1')
    expect(doc.headingStyleIds.get(2)).toBe('Heading2')
    expect(doc.listParagraphStyleId).toBe('ListParagraph')
  })

  it('protects structural paragraphs (section break / fields) as passthrough', async () => {
    const { buildDocx } = await import('./helpers/build-docx')
    const bytes = await buildDocx({
      bodyXml: [
        // mid-document section break lives inside w:pPr
        '<w:p><w:pPr><w:sectPr><w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/></w:sectPr></w:pPr></w:p>',
        // TOC field (complex fields still protect the whole paragraph; simple inline
        // fields PAGE/DATE are already editable)
        '<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" </w:instrText></w:r>' +
          '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>',
        '<w:p><w:r><w:t>带脚注的段落</w:t></w:r><w:r><w:footnoteReference w:id="2"/></w:r></w:p>',
        // comments stay editable: they are tracked at run level, not as passthrough
        '<w:p><w:commentRangeStart w:id="0"/><w:r><w:t>有批注</w:t></w:r><w:commentRangeEnd w:id="0"/></w:p>',
        '<w:p><w:r><w:t>普通可编辑段落</w:t></w:r></w:p>',
      ].join(''),
    })
    const doc = await parseDocx(bytes)
    const visible = doc.blocks.filter((b) => !b.hidden)
    expect(visible.map((b) => b.type)).toEqual([
      'passthrough',
      'passthrough',
      'paragraph',
      'paragraph',
      'paragraph',
    ])
    expect(visible[0].label).toBe('Section break paragraph')
    expect(visible[1].label).toBe('Auto TOC (updates when opened in Word)')
    // footnote references stay editable as atomic noteRef runs
    expect(visible[2].runs).toEqual([
      { text: '带脚注的段落' },
      { text: '*', noteRef: { kind: 'footnote', id: '2' } },
    ])
    expect(visible[3].runs).toEqual([{ text: '有批注', commentIds: ['0'] }])
  })

  it('keeps the picture of a text-less section-break paragraph (full-bleed cover)', async () => {
    const { buildDocx } = await import('./helpers/build-docx')
    const drawing =
      '<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">' +
      '<wp:extent cx="7559040" cy="10692130"/><wp:docPr id="1" name="Cover"/>' +
      '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
      '<pic:pic><pic:nvPicPr><pic:cNvPr id="1" name="Cover"/><pic:cNvPicPr/></pic:nvPicPr>' +
      '<pic:blipFill><a:blip r:embed="rId10"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>' +
      '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="7559040" cy="10692130"/></a:xfrm>' +
      '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>' +
      '</a:graphicData></a:graphic></wp:inline></w:drawing>'
    const bytes = await buildDocx({
      withImage: true,
      bodyXml:
        '<w:p><w:pPr><w:sectPr><w:pgSz w:w="11906" w:h="16838"/>' +
        '<w:pgMar w:top="0" w:right="0" w:bottom="0" w:left="0"/></w:sectPr></w:pPr>' +
        `<w:r>${drawing}</w:r></w:p>` +
        '<w:p><w:r><w:t>body</w:t></w:r></w:p>',
    })
    const doc = await parseDocx(bytes)
    const cover = doc.blocks.filter((b) => !b.hidden)[0]
    expect(cover.type).toBe('paragraph')
    expect(cover.runs?.[0].image?.dataUrl).toMatch(/^data:image\/png;base64,/)
    expect(cover.runs?.[0].image?.widthPx).toBe(793.6)
    // the section break rides along in rawPPr and survives a regeneration
    expect(cover.rawPPr).toContain('<w:sectPr>')
  })

  it('detects headings by effective outline level, not only Heading1-style paragraphs', async () => {
    const { buildDocx } = await import('./helpers/build-docx')
    const p = (pPr: string, text: string) =>
      `<w:p><w:pPr>${pPr}</w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`
    const bytes = await buildDocx({
      extraStylesXml:
        // custom style with no own outlineLvl: inherits level 2 through basedOn
        '<w:style w:type="paragraph" w:styleId="MySub"><w:name w:val="My Sub"/><w:basedOn w:val="Heading2"/></w:style>' +
        // Word's TOCHeading pattern: basedOn Heading1 but outlineLvl 9 = body text
        '<w:style w:type="paragraph" w:styleId="TOCHeading"><w:name w:val="TOC Heading"/><w:basedOn w:val="Heading1"/>' +
        '<w:pPr><w:outlineLvl w:val="9"/></w:pPr></w:style>' +
        // outline-off root style: no basedOn to resolve, the flag must still be set
        '<w:style w:type="paragraph" w:styleId="Caption"><w:name w:val="caption"/>' +
        '<w:pPr><w:outlineLvl w:val="9"/></w:pPr></w:style>',
      bodyXml: [
        p('<w:pStyle w:val="Heading1"/>', 'h1 via built-in style'),
        p('<w:pStyle w:val="Heading2"/>', 'h2 via built-in style'),
        p('<w:pStyle w:val="Heading3"/>', 'h3 via undefined built-in styleId'),
        p('<w:outlineLvl w:val="2"/>', 'h3 via direct outlineLvl'),
        p('<w:pStyle w:val="MySub"/>', 'h2 inherited through basedOn'),
        p('<w:pStyle w:val="TOCHeading"/>', 'body text: style outlineLvl 9'),
        p(
          '<w:pStyle w:val="Heading1"/><w:outlineLvl w:val="9"/>',
          'body text: direct outlineLvl 9',
        ),
      ].join(''),
    })
    const doc = await parseDocx(bytes)
    const visible = doc.blocks.filter((b) => !b.hidden)
    expect(visible.map((b) => [b.type, b.level])).toEqual([
      ['heading', 1],
      ['heading', 2],
      ['heading', 3],
      ['heading', 3],
      ['heading', 2],
      ['paragraph', undefined],
      ['paragraph', undefined],
    ])
    expect(doc.styles.get('TOCHeading')).toMatchObject({ headingOutlineOff: true })
    expect(doc.styles.get('TOCHeading')?.headingLevel).toBeUndefined()
    expect(doc.styles.get('Caption')).toMatchObject({ headingOutlineOff: true })
    expect(doc.styles.get('Caption')?.headingLevel).toBeUndefined()
    expect(doc.styles.get('MySub')).toMatchObject({ headingLevel: 2, headingLevelInherited: true })
  })
})

describe('style-level pageBreakBefore', () => {
  it('parses pageBreakBefore into style display, inherited via basedOn, without touching paragraph format', async () => {
    const bytes = await buildDocx({
      extraStylesXml:
        '<w:style w:type="paragraph" w:styleId="ChapterTitle"><w:name w:val="Chapter Title"/>' +
        '<w:pPr><w:pageBreakBefore/></w:pPr></w:style>' +
        '<w:style w:type="paragraph" w:styleId="ChapterSub"><w:name w:val="Chapter Sub"/>' +
        '<w:basedOn w:val="ChapterTitle"/></w:style>',
      bodyXml:
        '<w:p><w:pPr><w:pStyle w:val="ChapterTitle"/></w:pPr><w:r><w:t>ch</w:t></w:r></w:p>' +
        '<w:p><w:r><w:t>body</w:t></w:r></w:p>',
    })
    const doc = await parseDocx(bytes)
    expect(doc.styles.get('ChapterTitle')?.display?.pageBreakBefore).toBe(true)
    expect(doc.styles.get('ChapterSub')?.display?.pageBreakBefore).toBe(true)
    expect(doc.styles.get('Normal')?.display?.pageBreakBefore).toBeUndefined()
    // style-level value must not leak into paragraph format (would be saved as redundant pPr)
    expect(doc.blocks[0].format?.pageBreakBefore).toBeUndefined()
  })

  it('an explicit w:val="0" overrides an inherited true (fdo#45183)', async () => {
    const bytes = await buildDocx({
      extraStylesXml:
        '<w:style w:type="paragraph" w:styleId="ChapterTitle"><w:name w:val="Chapter Title"/>' +
        '<w:pPr><w:pageBreakBefore/></w:pPr></w:style>' +
        '<w:style w:type="paragraph" w:styleId="NoBreak"><w:name w:val="No Break"/>' +
        '<w:basedOn w:val="ChapterTitle"/><w:pPr><w:pageBreakBefore w:val="0"/></w:pPr></w:style>',
      bodyXml:
        '<w:p><w:pPr><w:pStyle w:val="ChapterTitle"/><w:pageBreakBefore w:val="0"/></w:pPr>' +
        '<w:r><w:t>off</w:t></w:r></w:p>',
    })
    const doc = await parseDocx(bytes)
    expect(doc.styles.get('NoBreak')?.display?.pageBreakBefore).toBe(false)
    // direct-format off must survive so it can veto the style chain's true
    expect(doc.blocks[0].format?.pageBreakBefore).toBe(false)
  })
})

describe('empty paragraph line size', () => {
  it('records the w:sz that governs a run-less paragraph', async () => {
    const bodyXml =
      '<w:p><w:r><w:t>before</w:t></w:r></w:p>' +
      // paragraph-mark rPr wins
      '<w:p><w:pPr><w:rPr><w:sz w:val="2"/></w:rPr></w:pPr><w:r><w:rPr><w:sz w:val="4"/></w:rPr><w:t></w:t></w:r></w:p>' +
      // no pPr rPr: falls back to the (dropped) empty run
      '<w:p><w:r><w:rPr><w:sz w:val="16"/></w:rPr><w:t></w:t></w:r></w:p>'
    const doc = await parseDocx(await buildDocx({ bodyXml }))
    expect(doc.blocks[0].format?.emptyRunSizeHalfPoints).toBeUndefined()
    expect(doc.blocks[1].runs).toEqual([])
    expect(doc.blocks[1].format?.emptyRunSizeHalfPoints).toBe(2)
    expect(doc.blocks[2].format?.emptyRunSizeHalfPoints).toBe(16)
  })

  it('records the paragraph-mark w:sz of a list item (Word sizes the marker from it)', async () => {
    const numberingXml =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
      '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      '<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/>' +
      '<w:numFmt w:val="bullet"/><w:lvlText w:val="\u2022"/></w:lvl></w:abstractNum>' +
      '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>'
    const numPr = '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>'
    const mark = '<w:rPr><w:sz w:val="22"/></w:rPr>'
    const run = '<w:r><w:rPr><w:sz w:val="18"/></w:rPr><w:t>text</w:t></w:r>'
    const bodyXml =
      `<w:p><w:pPr>${numPr}${mark}</w:pPr>${run}</w:p>` +
      `<w:p><w:pPr>${numPr}</w:pPr>${run}</w:p>` +
      `<w:p><w:pPr>${mark}</w:pPr>${run}</w:p>`
    const doc = await parseDocx(await buildDocx({ bodyXml, numberingXml }))
    expect(doc.blocks[0].format?.markSizeHalfPoints).toBe(22)
    expect(doc.blocks[0].format?.emptyRunSizeHalfPoints).toBeUndefined()
    expect(doc.blocks[1].format?.markSizeHalfPoints).toBeUndefined()
    // plain paragraphs do not carry it
    expect(doc.blocks[2].format?.markSizeHalfPoints).toBeUndefined()
  })

  // Word probe 2026-09-11: a space-only paragraph lays out like an empty one,
  // sized by the paragraph mark; the space run's own size never counts
  it('a space-only paragraph takes the mark rPr only, never the space run', async () => {
    const space = '<w:r><w:rPr><w:sz w:val="8"/></w:rPr><w:t xml:space="preserve"> </w:t></w:r>'
    const bodyXml =
      `<w:p>${space}</w:p>` +
      `<w:p><w:pPr><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="8"/></w:rPr></w:pPr>${space}</w:p>` +
      '<w:p><w:r><w:rPr><w:sz w:val="8"/></w:rPr><w:t xml:space="preserve"> x</w:t></w:r></w:p>'
    const doc = await parseDocx(await buildDocx({ bodyXml }))
    expect(doc.blocks[0]?.runs?.map((r) => r.text)).toEqual([' '])
    expect(doc.blocks[0].format?.emptyRunSizeHalfPoints).toBeUndefined()
    expect(doc.blocks[0].format?.emptyRunFontFamily).toBeUndefined()
    expect(doc.blocks[1].format?.emptyRunSizeHalfPoints).toBe(8)
    expect(doc.blocks[1].format?.emptyRunFontFamily).toBe('Arial')
    expect(doc.blocks[2].format?.emptyRunSizeHalfPoints).toBeUndefined()
  })

  it('a space-only textbox paragraph records its mark w:sz too', async () => {
    const space = '<w:r><w:rPr><w:sz w:val="8"/></w:rPr><w:t xml:space="preserve"> </w:t></w:r>'
    const txbx =
      '<w:p><w:r><w:pict><v:shape id="s1" style="width:100pt;height:40pt"><v:textbox><w:txbxContent>' +
      `<w:p><w:pPr><w:rPr><w:sz w:val="8"/></w:rPr></w:pPr>${space}</w:p><w:p>${space}</w:p>` +
      '</w:txbxContent></v:textbox></v:shape></w:pict></w:r></w:p>'
    const doc = await parseDocx(await buildDocx({ bodyXml: txbx }))
    const boxes = doc.blocks.flatMap((b) => b.textboxes ?? [])
    expect(boxes.length).toBe(1)
    expect(boxes[0]?.paras[0]?.emptyRunSizeHalfPoints).toBe(8)
    expect(boxes[0]?.paras[1]?.emptyRunSizeHalfPoints).toBeUndefined()
  })

  it("a break-only paragraph takes the mark face and size, else the break run's", async () => {
    const br = '<w:br w:type="page"/>'
    const bodyXml =
      `<w:p><w:pPr><w:rPr><w:rFonts w:ascii="Tahoma" w:hAnsi="Tahoma"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Tahoma" w:hAnsi="Tahoma"/></w:rPr>${br}</w:r></w:p>` +
      `<w:p><w:r><w:rPr><w:rFonts w:ascii="Arial"/><w:sz w:val="24"/></w:rPr>${br}</w:r></w:p>` +
      `<w:p><w:r><w:rPr><w:rFonts w:ascii="Arial"/></w:rPr><w:t>text</w:t>${br}</w:r></w:p>`
    const doc = await parseDocx(await buildDocx({ bodyXml }))
    expect(doc.blocks[0].format?.emptyRunFontFamily).toBe('Tahoma')
    expect(doc.blocks[0].format?.emptyRunSizeHalfPoints).toBeUndefined()
    expect(doc.blocks[1].format?.emptyRunFontFamily).toBe('Arial')
    expect(doc.blocks[1].format?.emptyRunSizeHalfPoints).toBe(24)
    expect(doc.blocks[2].format?.emptyRunFontFamily).toBeUndefined()
    expect(doc.blocks[2].format?.emptyRunSizeHalfPoints).toBeUndefined()
  })

  it('records the w:rFonts that faces a run-less paragraph', async () => {
    const bodyXml =
      '<w:p><w:r><w:t>before</w:t></w:r></w:p>' +
      '<w:p><w:pPr><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/></w:rPr></w:pPr></w:p>' +
      // no pPr rPr: falls back to the (dropped) empty run
      '<w:p><w:r><w:rPr><w:rFonts w:ascii="Arial"/></w:rPr><w:t></w:t></w:r></w:p>' +
      // East Asian slot only: the inherited Latin face keeps sizing the line
      // (Word probe 2026-09-05: a DengXian-only mark lays a 12pt line at the
      // ascii face's 14.6pt, not DengXian's 16.4pt)
      '<w:p><w:pPr><w:rPr><w:rFonts w:eastAsia="DengXian"/><w:lang w:eastAsia="zh-CN"/></w:rPr></w:pPr></w:p>'
    const doc = await parseDocx(await buildDocx({ bodyXml }))
    expect(doc.blocks[0].format?.emptyRunFontFamily).toBeUndefined()
    expect(doc.blocks[1].format?.emptyRunFontFamily).toBe('Times New Roman')
    expect(doc.blocks[2].format?.emptyRunFontFamily).toBe('Arial')
    expect(doc.blocks[3].format?.emptyRunFontFamily).toBeUndefined()
  })
})

describe('paragraph-mark w:vanish', () => {
  it('collapses an empty paragraph with a hidden mark to an invisible marker', async () => {
    const bodyXml =
      '<w:p><w:r><w:t>before</w:t></w:r></w:p>' +
      '<w:p><w:pPr><w:ind w:left="120"/><w:rPr><w:vanish/></w:rPr></w:pPr></w:p>' +
      '<w:p><w:r><w:t>after</w:t></w:r></w:p>'
    const doc = await parseDocx(await buildDocx({ bodyXml }))
    expect(doc.blocks[1].invisibleMarker).toBe(true)
    expect(doc.blocks[1].originalXml).toContain('w:vanish')
  })

  it('keeps a paragraph with visible runs even when its mark is hidden', async () => {
    const bodyXml =
      '<w:p><w:pPr><w:rPr><w:vanish/></w:rPr></w:pPr><w:r><w:t>shown</w:t></w:r></w:p>'
    const doc = await parseDocx(await buildDocx({ bodyXml }))
    expect(doc.blocks[0].invisibleMarker).toBeUndefined()
    expect(doc.blocks[0].runs?.[0]?.text).toBe('shown')
  })

  it('keeps a text-less paragraph whose run carries a page break or note mark', async () => {
    const bodyXml =
      '<w:p><w:pPr><w:rPr><w:vanish/></w:rPr></w:pPr><w:r><w:br w:type="page"/></w:r></w:p>' +
      '<w:p><w:pPr><w:rPr><w:vanish/></w:rPr></w:pPr><w:r><w:footnoteReference w:id="2"/></w:r></w:p>'
    const doc = await parseDocx(await buildDocx({ bodyXml }))
    expect(doc.blocks[0].invisibleMarker).toBeUndefined()
    expect(doc.blocks[1].invisibleMarker).toBeUndefined()
  })

  it('still collapses when only pPr tab stops are present', async () => {
    const bodyXml =
      '<w:p><w:pPr><w:tabs><w:tab w:val="left" w:pos="720"/></w:tabs>' +
      '<w:rPr><w:vanish/></w:rPr></w:pPr></w:p>'
    const doc = await parseDocx(await buildDocx({ bodyXml }))
    expect(doc.blocks[0].invisibleMarker).toBe(true)
  })
})
