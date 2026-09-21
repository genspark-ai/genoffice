import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'
import JSZip from 'jszip'
import { parseFileToText } from '../src/index'
import { pptxToText } from '../src/pptx'
import { xlsxToText } from '../src/xlsx'
import { resolveTarget } from '../src/opc'
import {
  buildDocxFixture,
  buildPptxFixture,
  buildXlsxFixture,
  writeFixture,
} from './helpers/fixtures'

function legacyFixture(name: string): string {
  return fileURLToPath(new URL(`fixtures/${name}`, import.meta.url))
}

describe('parseFileToText: doc', () => {
  it('extracts body text from a Word 97-2003 document', async () => {
    const result = await parseFileToText(legacyFixture('legacy-sample.doc'))
    expect(result.ok).toBe(true)
    expect(result.kind).toBe('text')
    expect(result.text).toContain('Legacy Report')
    expect(result.text).toContain('Legacy DOC body text')
    expect(result.text).toContain('Second paragraph from Word 97-2003.')
  })
})

describe('parseFileToText: docx', () => {
  it('extracts headings, paragraphs and tables', async () => {
    const path = writeFixture('report.docx', await buildDocxFixture())
    const result = await parseFileToText(path)
    expect(result.ok).toBe(true)
    expect(result.kind).toBe('text')
    expect(result.text).toContain('# Annual Report')
    expect(result.text).toContain('First paragraph hello docx')
    expect(result.text).toContain('Metric | Value')
    expect(result.text).toContain('Revenue | 100')
  })

  it('includes footnote and endnote text', async () => {
    const zip = await JSZip.loadAsync(await buildDocxFixture())
    const part = (tag: string, id: string, text: string) =>
      `<w:${tag} w:id="${id}"><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:${tag}>`
    zip.file(
      'word/footnotes.xml',
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
        part('footnote', '-1', '') +
        part('footnote', '1', 'Footnote detail') +
        '</w:footnotes>',
    )
    zip.file(
      'word/endnotes.xml',
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
        part('endnote', '1', 'Endnote detail') +
        '</w:endnotes>',
    )
    const path = writeFixture(
      'notes.docx',
      await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }),
    )
    const result = await parseFileToText(path)
    expect(result.ok).toBe(true)
    expect(result.text).toContain('Footnote detail')
    expect(result.text).toContain('Endnote detail')
  })
})

describe('parseFileToText: ppt', () => {
  it('extracts one text section per slide from a PowerPoint 97-2003 presentation', async () => {
    const result = await parseFileToText(legacyFixture('legacy-sample.ppt'))
    expect(result.ok).toBe(true)
    expect(result.kind).toBe('text')
    expect(result.text).toContain('## Slide 1')
    expect(result.text).toContain('Legacy PPT title')
    expect(result.text).toContain('First slide body')
    expect(result.text).toContain('## Slide 2')
    expect(result.text).toContain('Second legacy slide')
  })
})

describe('parseFileToText: pptx', () => {
  it('extracts one section per slide in numeric order', async () => {
    const path = writeFixture('deck.pptx', await buildPptxFixture())
    const result = await parseFileToText(path)
    expect(result.ok).toBe(true)
    expect(result.text).toContain('## Slide 1\nProductIntro\nFirst slide subtitle')
    expect(result.text).toContain('## Slide 2\nMarket Analysis')
    // slide10 must sort after slide2 (numeric, not lexicographic)
    expect(result.text!.indexOf('## Slide 10')).toBeGreaterThan(result.text!.indexOf('## Slide 2'))
    expect(result.text).toContain('## Slide 10\nSummary Slide')
  })

  it('keeps run text verbatim: leading zeros and the spaces between runs', async () => {
    const path = writeFixture('deck.pptx', await buildPptxFixture())
    const result = await parseFileToText(path)
    expect(result.text).toContain('Order 0042')
  })

  it('keeps a:br as a line break and a:fld in document order', async () => {
    const path = writeFixture('deck.pptx', await buildPptxFixture())
    const result = await parseFileToText(path)
    // two breaks, so the run text is not concatenated and the field lands between its runs;
    // the slide also carries a comment naming those tags, which must not reach the walker
    expect(result.text).toContain('## Slide 3\nBefore\n\nAfter\nPage 3 of 10')
  })

  it('takes text from a:t only, not from whitespace inside sibling elements', async () => {
    const path = writeFixture('deck.pptx', await buildPptxFixture())
    const result = await parseFileToText(path)
    // the a:br elements are written across lines; that layout whitespace is a value too
    expect(result.text).toContain('Before')
    expect(result.text).not.toMatch(/Before\n[^\S\n]/)
    // and the comment the slide carries is markup, not text
    expect(result.text).not.toContain('authoring note')
  })

  const PIC_SLIDE =
    '<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" ' +
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
    '<p:cSld><p:spTree>' +
    '<p:pic><p:blipFill><a:blip r:embed="rId2"/></p:blipFill></p:pic>' +
    '<p:grpSp><p:pic><p:blipFill><a:blip r:embed="rId3"/></p:blipFill></p:pic></p:grpSp>' +
    '<p:sp><p:txBody><a:p><a:endParaRPr/></a:p></p:txBody></p:sp>' +
    '</p:spTree></p:cSld></p:sld>'

  it('marks picture-only slides instead of emitting a bare heading', async () => {
    const zip = new JSZip()
    zip.file(
      'ppt/slides/slide1.xml',
      '<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" ' +
        'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
        '<p:cSld><p:spTree><p:sp><p:txBody><a:p><a:r><a:t>Agenda</a:t></a:r></a:p>' +
        '</p:txBody></p:sp></p:spTree></p:cSld></p:sld>',
    )
    zip.file('ppt/slides/slide2.xml', PIC_SLIDE)
    const text = await pptxToText(await zip.generateAsync({ type: 'uint8array' }))
    expect(text).toContain('## Slide 1\nAgenda')
    expect(text).toContain('## Slide 2\n[picture-only slide: 2 images, no extractable text]')
    expect(text).not.toContain('No extractable text')
  })

  it('leads with a deck-level note when every slide is picture-only', async () => {
    const zip = new JSZip()
    zip.file('ppt/slides/slide1.xml', PIC_SLIDE)
    zip.file('ppt/slides/slide2.xml', PIC_SLIDE)
    const text = await pptxToText(await zip.generateAsync({ type: 'uint8array' }))
    expect(text.startsWith('[No extractable text: none of the 2 slides carries text')).toBe(true)
    expect(text).toContain('## Slide 1\n[picture-only slide: 2 images, no extractable text]')
  })

  it('leaves a blank slide (no text, no pictures) as a bare heading', async () => {
    const zip = new JSZip()
    zip.file(
      'ppt/slides/slide1.xml',
      '<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">' +
        '<p:cSld><p:spTree/></p:cSld></p:sld>',
    )
    expect(await pptxToText(await zip.generateAsync({ type: 'uint8array' }))).toBe('## Slide 1')
  })

  it('keeps a:tab as a tab between runs', async () => {
    const zip = new JSZip()
    zip.file(
      'ppt/slides/slide1.xml',
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" ' +
        'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
        '<p:cSld><p:spTree><p:sp><p:txBody><a:p>' +
        '<a:r><a:t>Col1</a:t></a:r><a:tab/><a:r><a:t>Col2</a:t></a:r>' +
        '</a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>',
    )
    const bytes = await zip.generateAsync({ type: 'uint8array' })
    expect(await pptxToText(bytes)).toContain('Col1\tCol2')
  })

  async function presentationFixture(slideIds: string, relationships: string): Promise<JSZip> {
    const zip = await JSZip.loadAsync(await buildPptxFixture())
    zip.file(
      'ppt/presentation.xml',
      '<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" ' +
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
        `<p:sldIdLst>${slideIds}</p:sldIdLst></p:presentation>`,
    )
    zip.file(
      'ppt/_rels/presentation.xml.rels',
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        relationships +
        '</Relationships>',
    )
    return zip
  }

  function slideRelationship(id: string, target: string, extra = ''): string {
    return (
      `<Relationship Id="${id}" Target="${target}" ${extra} ` +
      'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"/>'
    )
  }

  it('follows presentation order with positional numbering and excludes orphan slides', async () => {
    const zip = await presentationFixture(
      "<p:sldId id='265' r:id='rId10'/><p:sldId id='256' r:id='rId1'/>",
      slideRelationship('rId1', 'slides/slide1.xml') +
        slideRelationship('rId10', 'slides/slide10.xml') +
        slideRelationship('unused', 'slides/slide2.xml'),
    )
    const bytes = await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' })
    const original = bytes.slice()
    expect(await pptxToText(bytes)).toBe(
      '## Slide 1\nSummary Slide\n\n## Slide 2\nProductIntro\nFirst slide subtitle',
    )
    expect(bytes).toEqual(original)
  })

  it.each([
    'slides/custom.xml',
    '/ppt/slides/custom.xml',
    './slides/../slides/custom.xml',
    '../ppt/slides/custom.xml',
    ' slides/custom.xml ',
  ])('resolves a single slide relationship target %s with a custom part name', async (target) => {
    const zip = await presentationFixture(
      '<p:sldId id="256" r:id="custom"/>',
      slideRelationship('custom', target),
    )
    zip.file('ppt/slides/custom.xml', await zip.file('ppt/slides/slide10.xml')!.async('text'))
    expect(await pptxToText(await zip.generateAsync({ type: 'uint8array' }))).toBe(
      '## Slide 1\nSummary Slide',
    )
  })

  it.each(['slides\\custom.xml', 'slides\\..\\slides\\custom.xml'])(
    'resolves a Windows-style backslash relationship target %s',
    async (target) => {
      const zip = await presentationFixture(
        '<p:sldId id="256" r:id="custom"/>',
        slideRelationship('custom', target),
      )
      zip.file('ppt/slides/custom.xml', await zip.file('ppt/slides/slide10.xml')!.async('text'))
      expect(await pptxToText(await zip.generateAsync({ type: 'uint8array' }))).toBe(
        '## Slide 1\nSummary Slide',
      )
    },
  )

  it('clamps above-root dot-dot chains at the zip root', async () => {
    const zip = await presentationFixture(
      '<p:sldId id="256" r:id="custom"/>',
      slideRelationship('custom', '../../ppt/slides/custom.xml'),
    )
    zip.file('ppt/slides/custom.xml', await zip.file('ppt/slides/slide10.xml')!.async('text'))
    expect(await pptxToText(await zip.generateAsync({ type: 'uint8array' }))).toBe(
      '## Slide 1\nSummary Slide',
    )
  })

  it('preserves positions across missing ids, parts, blank slides and rejected relationships', async () => {
    const zip = await presentationFixture(
      '<p:sldId id="256"/><p:sldId r:id=""/><p:sldId r:id="unknown"/>' +
        '<p:sldId r:id="missing"/><p:sldId r:id="external"/><p:sldId r:id="wrongType"/>' +
        '<p:sldId r:id="noTarget"/><p:sldId r:id="blank"/><p:sldId r:id="last"/>',
      slideRelationship('', 'slides/slide1.xml') +
        slideRelationship('missing', 'slides/missing.xml') +
        slideRelationship('external', 'slides/slide2.xml', 'TargetMode="External"') +
        '<Relationship Id="wrongType" Target="slides/slide3.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide"/>' +
        '<Relationship Target="slides/slide1.xml"/>' +
        '<Relationship Id="noTarget" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"/>' +
        slideRelationship('blank', 'slides/blank.xml') +
        slideRelationship('last', 'slides/slide10.xml'),
    )
    zip.file(
      'ppt/slides/blank.xml',
      '<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree/></p:cSld></p:sld>',
    )
    expect(await pptxToText(await zip.generateAsync({ type: 'uint8array' }))).toBe(
      '## Slide 8\n\n## Slide 9\nSummary Slide',
    )
  })

  it.each(['empty list', 'missing list', 'missing relationships'])(
    'does not fall back to orphan slides when the manifest has %s',
    async (kind) => {
      const zip = await presentationFixture('<p:sldId r:id="rId1"/>', '')
      if (kind === 'empty list') {
        zip.file('ppt/presentation.xml', '<p:presentation><p:sldIdLst/></p:presentation>')
      } else if (kind === 'missing list') {
        zip.file('ppt/presentation.xml', '<p:presentation/>')
      } else {
        zip.remove('ppt/_rels/presentation.xml.rels')
      }
      expect(await pptxToText(await zip.generateAsync({ type: 'uint8array' }))).toBe('')
    },
  )

  it('retains numeric fallback without a manifest and leaves input bytes unchanged', async () => {
    const bytes = await buildPptxFixture()
    const original = bytes.slice()
    const text = await pptxToText(bytes)
    expect(text.match(/^## Slide \d+$/gm)).toEqual([
      '## Slide 1',
      '## Slide 2',
      '## Slide 3',
      '## Slide 10',
    ])
    expect(bytes).toEqual(original)
  })
})

describe('parseFileToText: xlsx', () => {
  it('extracts sheet name, shared/inline strings, numbers, booleans and empty columns', async () => {
    const path = writeFixture('table.xlsx', await buildXlsxFixture())
    const result = await parseFileToText(path)
    expect(result.ok).toBe(true)
    expect(result.text).toContain('# Grades')
    expect(result.text).toContain('Name | Scores')
    // C2 is missing so the boolean in D2 lands in the 4th column
    expect(result.text).toContain('Alice | 95 |  | TRUE')
  })

  it('keeps cell text verbatim: leading zeros and the spaces between rich-text runs', async () => {
    const path = writeFixture('table.xlsx', await buildXlsxFixture())
    const result = await parseFileToText(path)
    expect(result.text).toContain('02139 | Total due')
  })

  it('keeps the spaces in a <v> value (cached formula string, error literal)', async () => {
    const path = writeFixture('table.xlsx', await buildXlsxFixture())
    const result = await parseFileToText(path)
    expect(result.text).toContain('\n Alice pts \n #N/A ')
  })

  it('parses .xlsm through the same xlsx path', async () => {
    const path = writeFixture('table.xlsm', await buildXlsxFixture())
    const result = await parseFileToText(path)
    expect(result.ok).toBe(true)
    expect(result.text).toContain('# Grades')
  })

  it('fails gracefully on a corrupt file', async () => {
    const path = writeFixture('broken.xlsx', Buffer.from('not a zip'))
    const result = await parseFileToText(path)
    expect(result.ok).toBe(false)
    expect(result.error).toBeTruthy()
  })

  it('marks an image-only sheet and counts the pictures in its drawing part', async () => {
    const zip = new JSZip()
    zip.file(
      'xl/workbook.xml',
      '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
        '<sheets><sheet name="Chart" sheetId="1" r:id="rId1"/><sheet name="Data" sheetId="2" r:id="rId2"/></sheets></workbook>',
    )
    zip.file(
      'xl/_rels/workbook.xml.rels',
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>' +
        '</Relationships>',
    )
    zip.file(
      'xl/worksheets/sheet1.xml',
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
        '<sheetData/><drawing r:id="rId1"/></worksheet>',
    )
    zip.file(
      'xl/worksheets/_rels/sheet1.xml.rels',
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>' +
        '</Relationships>',
    )
    zip.file(
      'xl/drawings/drawing1.xml',
      '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing">' +
        '<xdr:twoCellAnchor><xdr:pic/></xdr:twoCellAnchor>' +
        '</xdr:wsDr>',
    )
    zip.file(
      'xl/worksheets/sheet2.xml',
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>' +
        '<row r="1"><c r="A1"><v>7</v></c></row></sheetData></worksheet>',
    )
    const text = await xlsxToText(await zip.generateAsync({ type: 'uint8array' }))
    expect(text).toContain('# Chart\n[image-only sheet: 1 image, no cell data]')
    expect(text).toContain('# Data\n7')
    expect(text).not.toContain('No extractable text')
  })

  it('leads with a workbook-level note when every sheet is image-only', async () => {
    const zip = new JSZip()
    zip.file(
      'xl/workbook.xml',
      '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
        '<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>',
    )
    zip.file(
      'xl/_rels/workbook.xml.rels',
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
        '</Relationships>',
    )
    // no sheet rels at all: the drawing is still reported, just without a count
    zip.file(
      'xl/worksheets/sheet1.xml',
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
        '<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t> </t></is></c></row></sheetData>' +
        '<drawing r:id="rId1"/></worksheet>',
    )
    const text = await xlsxToText(await zip.generateAsync({ type: 'uint8array' }))
    expect(text.startsWith('[No extractable text: none of the 1 sheet holds cell data')).toBe(true)
    expect(text).toContain('# Sheet1\n \n[image-only sheet: a drawing but no cell data]')
  })

  it('reads lowercase cell refs at the right columns', async () => {
    const zip = new JSZip()
    zip.file(
      'xl/workbook.xml',
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
        '<sheets><sheet name="S1" sheetId="1" r:id="rId1"/></sheets></workbook>',
    )
    zip.file(
      'xl/_rels/workbook.xml.rels',
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
        '</Relationships>',
    )
    zip.file(
      'xl/worksheets/sheet1.xml',
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>' +
        '<row r="1"><c r="a1"><v>1</v></c><c r="c1"><v>3</v></c></row>' +
        '</sheetData></worksheet>',
    )
    const bytes = await zip.generateAsync({ type: 'uint8array' })
    expect(await xlsxToText(bytes)).toContain('1 |  | 3')
  })

  it('keeps an empty shared-string cell empty instead of leaking shared[0]', async () => {
    const zip = new JSZip()
    zip.file(
      'xl/workbook.xml',
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
        '<sheets><sheet name="S1" sheetId="1" r:id="rId1"/></sheets></workbook>',
    )
    zip.file(
      'xl/_rels/workbook.xml.rels',
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
        '</Relationships>',
    )
    zip.file(
      'xl/sharedStrings.xml',
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="2" uniqueCount="2">' +
        '<si><t>First</t></si><si><t>Second</t></si></sst>',
    )
    zip.file(
      'xl/worksheets/sheet1.xml',
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>' +
        '<row r="1"><c r="A1" t="s"><v>0</v></c>' +
        '<c r="B1" t="s"><v></v></c>' +
        '<c r="C1" t="s"/>' +
        '<c r="D1" t="s"><v>   </v></c>' +
        '<c r="E1" t="s"><v>99</v></c>' +
        '<c r="F1" t="s"><v>not-a-number</v></c></row>' +
        '</sheetData></worksheet>',
    )
    const bytes = await zip.generateAsync({ type: 'uint8array' })
    const text = await xlsxToText(bytes)
    // Only the valid index 0 survives; every malformed shared ref degrades to empty.
    expect(text).toContain('First |  |  |  |  | ')
  })

  it('appends cells with malformed refs instead of dropping their text', async () => {
    const zip = new JSZip()
    zip.file(
      'xl/workbook.xml',
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
        '<sheets><sheet name="S1" sheetId="1" r:id="rId1"/></sheets></workbook>',
    )
    zip.file(
      'xl/_rels/workbook.xml.rels',
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
        '</Relationships>',
    )
    zip.file(
      'xl/worksheets/sheet1.xml',
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>' +
        '<row r="1"><c r="A1"><v>ok</v></c><c r="1"><v>orphan</v></c></row>' +
        '</sheetData></worksheet>',
    )
    const bytes = await zip.generateAsync({ type: 'uint8array' })
    const text = await xlsxToText(bytes)
    expect(text).toContain('ok | orphan')
  })

  it('clamps wild column refs instead of padding millions of cells', async () => {
    const zip = new JSZip()
    zip.file(
      'xl/workbook.xml',
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
        '<sheets><sheet name="S1" sheetId="1" r:id="rId1"/></sheets></workbook>',
    )
    zip.file(
      'xl/_rels/workbook.xml.rels',
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
        '</Relationships>',
    )
    zip.file(
      'xl/worksheets/sheet1.xml',
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>' +
        '<row r="1"><c r="A1"><v>ok</v></c><c r="XXXXXXX99"><v>bomb</v></c></row>' +
        '</sheetData></worksheet>',
    )
    const bytes = await zip.generateAsync({ type: 'uint8array' })
    const text = await xlsxToText(bytes)
    expect(text).toContain('ok | bomb')
    expect(text.length).toBeLessThan(1000)
  })

  it('resolves workbook rel targets against xl/workbook.xml', () => {
    const wb = (t: string) => resolveTarget('xl/workbook.xml', t)
    expect(wb('worksheets/sheet1.xml')).toBe('xl/worksheets/sheet1.xml')
    expect(wb('worksheets\\sheet1.xml')).toBe('xl/worksheets/sheet1.xml')
    expect(wb('/xl/worksheets/sheet1.xml')).toBe('xl/worksheets/sheet1.xml')
    expect(wb('../customXml/item1.xml')).toBe('customXml/item1.xml')
    expect(wb('../../xl/worksheets/sheet1.xml')).toBe('xl/worksheets/sheet1.xml')
  })

  it('resolves sheets through backslash rel targets from Windows producers', async () => {
    const zip = new JSZip()
    zip.file(
      'xl/workbook.xml',
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
        '<sheets><sheet name="S1" sheetId="1" r:id="rId1"/></sheets></workbook>',
    )
    zip.file(
      'xl/_rels/workbook.xml.rels',
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets\\sheet1.xml"/>' +
        '</Relationships>',
    )
    zip.file(
      'xl/worksheets/sheet1.xml',
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>' +
        '<row r="1"><c r="a1"><v>1</v></c><c r="c1"><v>3</v></c></row>' +
        '</sheetData></worksheet>',
    )
    const bytes = await zip.generateAsync({ type: 'uint8array' })
    expect(await xlsxToText(bytes)).toContain('1 |  | 3')
  })
})
