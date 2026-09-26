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

  it('accepts arbitrary namespace prefixes for presentation and slide parts', async () => {
    const zip = new JSZip()
    zip.file(
      'ppt/presentation.xml',
      '<p-x:presentation xmlns:p-x="http://schemas.openxmlformats.org/presentationml/2006/main" ' +
        'xmlns:rel-x="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
        '<p-x:sldIdLst><p-x:sldId id="not-rel-1" rel-x:id="rId1"/>' +
        '<p-x:sldId rel-x:id="rId2" id="not-rel-2"/>' +
        '</p-x:sldIdLst></p-x:presentation>',
    )
    zip.file(
      'ppt/_rels/presentation.xml.rels',
      '<rel-p:Relationships xmlns:rel-p="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<rel-p:Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>' +
        '<rel-p:Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide2.xml"/>' +
        '</rel-p:Relationships>',
    )
    zip.file(
      'ppt/slides/slide1.xml',
      '<s-x:sld xmlns:s-x="http://schemas.openxmlformats.org/presentationml/2006/main" ' +
        'xmlns:t-x="http://schemas.openxmlformats.org/drawingml/2006/main">' +
        '<s-x:cSld><s-x:spTree><s-x:sp><s-x:txBody><t-x:p><t-x:r><t-x:t>Prefixed text</t-x:t>' +
        '</t-x:r></t-x:p></s-x:txBody></s-x:sp></s-x:spTree></s-x:cSld></s-x:sld>',
    )
    zip.file(
      'ppt/slides/slide2.xml',
      '<s-x:sld xmlns:s-x="http://schemas.openxmlformats.org/presentationml/2006/main" ' +
        'xmlns:w-x="http://schemas.openxmlformats.org/presentationml/2006/main">' +
        '<s-x:cSld><s-x:spTree><w-x:pic/><w-x:pic/></s-x:spTree></s-x:cSld></s-x:sld>',
    )
    const text = await pptxToText(await zip.generateAsync({ type: 'uint8array' }))
    expect(text).toContain('## Slide 1\nPrefixed text')
    expect(text).toContain('## Slide 2\n[picture-only slide: 2 images, no extractable text]')
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

  /** minimal slide part: the txBody the walkers read, in the shape tree they walk to find it */
  function slideXml(body: string): string {
    return (
      '<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" ' +
      'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
      'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">' +
      `<p:cSld><p:spTree><p:sp><p:txBody>${body}</p:txBody></p:sp></p:spTree></p:cSld></p:sld>`
    )
  }

  /** the two branches of one shape's mc:AlternateContent; no fallback at all when it is omitted */
  function altContent(choice: string, fallback?: string): string {
    return (
      '<mc:AlternateContent>' +
      `<mc:Choice Requires="a14">${choice}</mc:Choice>` +
      (fallback === undefined ? '' : `<mc:Fallback>${fallback}</mc:Fallback>`) +
      '</mc:AlternateContent>'
    )
  }

  it('reads one branch of an mc:AlternateContent, never both', async () => {
    const zip = new JSZip()
    zip.file(
      'ppt/slides/slide1.xml',
      slideXml(
        // the branches carry different runs on purpose: identical ones would let a
        // walker read both without the duplication showing up here
        `<a:p><a:r><a:t xml:space="preserve">Total: </a:t></a:r>` +
          altContent('<a:r><a:t>21</a:t></a:r>', '<a:r><a:t>20</a:t></a:r>') +
          `</a:p>` +
          // a Choice with no Fallback is the branch a consumer that understands the
          // required namespaces would take, so it is the text that survives
          `<a:p>${altContent('<a:r><a:t>ChoiceOnly</a:t></a:r>')}</a:p>` +
          '<a:p><mc:AlternateContent/></a:p>' +
          '<a:p><a:r><a:t>Plain</a:t></a:r></a:p>',
      ),
    )
    expect(await pptxToText(await zip.generateAsync({ type: 'uint8array' }))).toBe(
      '## Slide 1\nTotal: 20\nChoiceOnly\nPlain',
    )
  })

  it('counts a picture its two branches both carry as the one image', async () => {
    const zip = new JSZip()
    zip.file('ppt/slides/slide1.xml', slideXml('<a:p><a:r><a:t>Agenda</a:t></a:r></a:p>'))
    const pic = '<p:pic><p:blipFill><a:blip r:embed="rId2"/></p:blipFill></p:pic>'
    zip.file(
      'ppt/slides/slide2.xml',
      '<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" ' +
        'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
        'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">' +
        `<p:cSld><p:spTree>${altContent(pic, pic)}</p:spTree></p:cSld></p:sld>`,
    )
    expect(await pptxToText(await zip.generateAsync({ type: 'uint8array' }))).toContain(
      '## Slide 2\n[picture-only slide: 1 image, no extractable text]',
    )
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

  function notesSlideXml(body: string[], slideNum: string): string {
    const paras = (runs: string[]) =>
      runs.map((t) => `<a:p><a:r><a:t xml:space="preserve">${t}</a:t></a:r></a:p>`).join('')
    return (
      '<p:notes xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" ' +
      'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree>' +
      '<p:sp><p:nvSpPr><p:nvPr><p:ph type="sldImg"/></p:nvPr></p:nvSpPr></p:sp>' +
      `<p:sp><p:nvSpPr><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:txBody>${paras(body)}</p:txBody></p:sp>` +
      '<p:sp><p:nvSpPr><p:nvPr><p:ph type="sldNum" sz="quarter" idx="10"/></p:nvPr></p:nvSpPr>' +
      `<p:txBody><a:p><a:fld id="{N}" type="slidenum"><a:t>${slideNum}</a:t></a:fld></a:p></p:txBody></p:sp>` +
      '</p:spTree></p:cSld></p:notes>'
    )
  }

  function notesRelationship(target: string): string {
    return (
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      `<Relationship Id="rId2" Target="${target}" ` +
      'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide"/>' +
      '</Relationships>'
    )
  }

  it('appends speaker notes after each slide and skips the slide-number field', async () => {
    const zip = await presentationFixture(
      "<p:sldId id='256' r:id='rId1'/><p:sldId id='257' r:id='rId2'/>",
      slideRelationship('rId1', 'slides/slide1.xml') +
        slideRelationship('rId2', 'slides/slide2.xml'),
    )
    zip.file(
      'ppt/slides/_rels/slide1.xml.rels',
      notesRelationship('../notesSlides/notesSlide1.xml'),
    )
    zip.file(
      'ppt/notesSlides/notesSlide1.xml',
      notesSlideXml(['Remember to greet the audience', 'Mention the Q3 numbers'], '1'),
    )
    expect(await pptxToText(await zip.generateAsync({ type: 'uint8array' }))).toBe(
      '## Slide 1\nProductIntro\nFirst slide subtitle\n### Notes\nRemember to greet the audience\nMention the Q3 numbers' +
        '\n\n## Slide 2\nMarket Analysis\nOrder 0042',
    )
  })

  it('counts notes as text on a slide that otherwise only holds pictures', async () => {
    const zip = new JSZip()
    zip.file(
      'ppt/slides/slide1.xml',
      '<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">' +
        '<p:cSld><p:spTree><p:pic/></p:spTree></p:cSld></p:sld>',
    )
    zip.file(
      'ppt/slides/_rels/slide1.xml.rels',
      notesRelationship('../notesSlides/notesSlide1.xml'),
    )
    zip.file('ppt/notesSlides/notesSlide1.xml', notesSlideXml(['The chart shows revenue'], '1'))
    expect(await pptxToText(await zip.generateAsync({ type: 'uint8array' }))).toBe(
      '## Slide 1\n[picture-only slide: 1 image, no extractable text]\n### Notes\nThe chart shows revenue',
    )
  })

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

  it('accepts arbitrary namespace prefixes for workbook and drawing parts', async () => {
    const zip = new JSZip()
    zip.file(
      'xl/workbook.xml',
      '<q:workbook xmlns:q="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
        'xmlns:rel="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
        '<q:sheets><q:sheet name="Data" sheetId="1" id="not-rel-1" rel:id="rId1"/>' +
        '<q:sheet rel:id="rId2" id="not-rel-2" name="Pictures" sheetId="2"/></q:sheets></q:workbook>',
    )
    zip.file(
      'xl/_rels/workbook.xml.rels',
      '<rel:Relationships xmlns:rel="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<rel:Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
        '<rel:Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>' +
        '</rel:Relationships>',
    )
    zip.file(
      'xl/worksheets/sheet1.xml',
      '<q:worksheet xmlns:q="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
        '<q:sheetData><q:row><q:c r="A1"><q:v>7</q:v></q:c></q:row></q:sheetData></q:worksheet>',
    )
    zip.file(
      'xl/worksheets/sheet2.xml',
      '<q:worksheet xmlns:q="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
        'xmlns:rel="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
        '<q:sheetData/><q:drawing id="not-rel-drawing" rel:id="rId1"/></q:worksheet>',
    )
    zip.file(
      'xl/worksheets/_rels/sheet2.xml.rels',
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>' +
        '</Relationships>',
    )
    zip.file(
      'xl/drawings/drawing1.xml',
      '<x-dr:wsDr xmlns:x-dr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing">' +
        '<x-dr:twoCellAnchor><x-dr:pic/></x-dr:twoCellAnchor>' +
        '<x-dr:twoCellAnchor><x-dr:pic/></x-dr:twoCellAnchor></x-dr:wsDr>',
    )
    const text = await xlsxToText(await zip.generateAsync({ type: 'uint8array' }))
    expect(text).toContain('# Data\n7')
    expect(text).toContain('# Pictures\n[image-only sheet: 2 images, no cell data]')
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

  it('surfaces formula text for cells without a cached value', async () => {
    const zip = new JSZip()
    zip.file(
      'xl/workbook.xml',
      '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
        '<sheets><sheet name="Calc" sheetId="1" r:id="rId1"/></sheets></workbook>',
    )
    zip.file(
      'xl/_rels/workbook.xml.rels',
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
        '</Relationships>',
    )
    zip.file(
      'xl/worksheets/sheet1.xml',
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>' +
        '<row r="1"><c r="A1"><v>1</v></c><c r="B1"><f>SUM(A1:A3)</f></c><c r="C1"><f>A1*2</f><v></v></c></row>' +
        '<row r="2"><c r="A2"><v>2</v></c><c r="B2"><f t="shared" ref="B2:B3" si="0">A2+1</f></c><c r="C2"><f>A2*2</f><v>4</v></c></row>' +
        '<row r="3"><c r="A3"><v>3</v></c><c r="B3"><f t="shared" si="0"/></c><c r="C3"><f t="shared" ref="C3:C4" si="1">A3*10</f><v>30</v></c></row>' +
        '<row r="4"><c r="C4"><f t="shared" si="1"/></c></row>' +
        '</sheetData></worksheet>',
    )
    const text = await xlsxToText(await zip.generateAsync({ type: 'uint8array' }))
    expect(text).toBe(
      '# Calc\n1 | =SUM(A1:A3) | =A1*2\n2 | =A2+1 | 4\n3 | =A2+1 | 30\n |  | =A3*10',
    )
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
