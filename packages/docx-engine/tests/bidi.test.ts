import { describe, expect, it } from 'vitest'
import { mergePPrFormat, parseDocx, saveDocx } from '../src/index'
import { mergeRPrModel } from '../src/generate'
import { buildDocx } from './helpers/build-docx'

const BIDI_BODY =
  '<w:p><w:pPr><w:bidi/><w:jc w:val="right"/></w:pPr><w:r><w:t>כותרת</w:t></w:r></w:p>' +
  '<w:p><w:pPr><w:bidi/><w:jc w:val="left"/></w:pPr><w:r><w:t>שמאל</w:t></w:r></w:p>' +
  '<w:p><w:pPr><w:bidi/></w:pPr><w:r><w:t>ברירת מחדל</w:t></w:r></w:p>' +
  '<w:p><w:pPr><w:jc w:val="right"/></w:pPr><w:r><w:t>LTR right</w:t></w:r></w:p>'

describe('RTL paragraphs (w:bidi + Word jc left/right swap)', () => {
  it('parses bidi and stores the visual alignment (jc left/right swapped)', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: BIDI_BODY }))
    expect(doc.blocks[0].format).toMatchObject({ bidi: true, align: 'left' })
    expect(doc.blocks[1].format).toMatchObject({ bidi: true, align: 'right' })
    expect(doc.blocks[2].format?.bidi).toBe(true)
    expect(doc.blocks[2].format?.align).toBeUndefined()
    // non-bidi paragraphs are not swapped
    expect(doc.blocks[3].format?.bidi).toBeUndefined()
    expect(doc.blocks[3].format?.align).toBe('right')
  })

  it('keeps untouched bidi paragraphs byte-identical', async () => {
    const bytes = await buildDocx({ bodyXml: BIDI_BODY })
    const doc = await parseDocx(bytes)
    const blocks = doc.blocks
      .filter((b) => !b.hidden && b.docxIndex !== null)
      .map((b) => ({ kind: 'original' as const, docxIndex: b.docxIndex! }))
    expect(await saveDocx(doc, blocks)).toEqual(bytes)
  })

  it('mergePPrFormat writes the logical jc back and keeps a single w:bidi', () => {
    const out = mergePPrFormat('<w:pPr><w:bidi/><w:jc w:val="right"/></w:pPr>', {
      bidi: true,
      align: 'left',
    })
    expect(out.match(/<w:bidi\/>/g)).toHaveLength(1)
    expect(out).toContain('<w:jc w:val="right"/>')
    // visual right maps back to logical left
    const out2 = mergePPrFormat('<w:pPr><w:bidi/><w:jc w:val="left"/></w:pPr>', {
      bidi: true,
      align: 'right',
    })
    expect(out2).toContain('<w:jc w:val="left"/>')
  })
})

describe('RTL tables (tblPr w:bidiVisual)', () => {
  it('parses bidiVisual into the table model', async () => {
    const xml =
      '<w:tbl><w:tblPr><w:bidiVisual/><w:tblW w:w="0" w:type="auto"/></w:tblPr>' +
      '<w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>' +
      '<w:tr><w:tc><w:p><w:r><w:t>א</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'
    const doc = await parseDocx(await buildDocx({ bodyXml: xml }))
    expect(doc.blocks[0].table?.bidiVisual).toBe(true)
  })
})

describe('complex-script bold and italic', () => {
  // Word renders Arabic and Hebrew bold from w:bCs, not w:b, so emitting only w:b makes
  // the Bold button do nothing to that text.
  const CS_RUN = '<w:rPr><w:rFonts w:cs="Traditional Arabic"/><w:szCs w:val="28"/></w:rPr>'

  it('emits the Cs twin when bold is turned on', () => {
    const out = mergeRPrModel(CS_RUN, { text: 'مرحبا', bold: true }, false)
    expect(out).toContain('<w:b/>')
    expect(out).toContain('<w:bCs/>')
  })

  it('emits the Cs twin when italic is turned on', () => {
    const out = mergeRPrModel(CS_RUN, { text: 'مرحبا', italic: true }, false)
    expect(out).toContain('<w:i/>')
    expect(out).toContain('<w:iCs/>')
  })

  it('restores it after the flag is turned off and on again', () => {
    const bold = '<w:rPr><w:rFonts w:cs="Traditional Arabic"/><w:b/><w:bCs/></w:rPr>'
    const off = mergeRPrModel(bold, { text: 'مرحبا', bold: false }, false)
    expect(off).not.toContain('<w:b/>')
    expect(off).not.toContain('<w:bCs/>')
    expect(mergeRPrModel(off, { text: 'مرحبا', bold: true }, false)).toContain('<w:bCs/>')
  })

  it('leaves an unchanged run on its original bytes', () => {
    const bold = '<w:rPr><w:rFonts w:cs="Traditional Arabic"/><w:b/><w:bCs/></w:rPr>'
    expect(mergeRPrModel(bold, { text: 'مرحبا', bold: true }, false)).toBe(bold)
  })

  it('rebuilds the twins the way size already does', () => {
    // w:sz has written its w:szCs companion from the model for as long as this path has
    // existed, so a rebuild has always restored a Cs value the original may not have carried.
    // Bold and italic were the two properties left out of that.
    const fresh = mergeRPrModel('', { text: 'مرحبا', bold: true, sizeHalfPoints: 28 }, false)
    expect(fresh).toContain('<w:szCs w:val="28"/>')
    expect(fresh).toContain('<w:b/><w:bCs/>')
  })
})
