import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import {
  applyRevisionResolutions,
  countTrackedChangesXml,
  parseDocx,
  patchTextboxParas,
  saveDocx,
  type SaveBlock,
} from '../src/index'
import { buildDocx } from './helpers/build-docx'

/** accept / reject of tracked changes in footnotes, header parts and textboxes */

const ANN = ' w:id="7" w:author="Ann" w:date="2026-09-29T12:23:00Z"'
const BOB = ' w:id="8" w:author="Bob" w:date="2026-09-30T08:00:00Z"'
const NS = ' xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'

const FOOTNOTES_XML =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  `<w:footnotes${NS}>` +
  '<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>' +
  '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>' +
  '<w:footnote w:id="2"><w:p><w:r><w:footnoteRef/></w:r><w:r><w:t xml:space="preserve"> Kept </w:t></w:r>' +
  `<w:del${ANN}><w:r><w:delText>gone</w:delText></w:r></w:del>` +
  `<w:ins${BOB}><w:r><w:t>added</w:t></w:r></w:ins></w:p></w:footnote>` +
  '</w:footnotes>'

const FOOTNOTES_PART = {
  extraRels:
    '<Relationship Id="rId40" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>',
  extraParts: [
    {
      path: 'word/footnotes.xml',
      xml: FOOTNOTES_XML,
      contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml',
    },
  ],
}

const DELETED_DRAWING = `<w:del${ANN}><w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><wp:extent cx="914400" cy="914400"/></wp:inline></w:drawing></w:r></w:del>`

const HEADER_XML =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr${NS}>` +
  '<w:p><w:r><w:t xml:space="preserve">Kept </w:t></w:r>' +
  `<w:del${ANN}><w:r><w:delText>gone</w:delText></w:r></w:del>` +
  `<w:ins${BOB}><w:r><w:t>added</w:t></w:r></w:ins></w:p>` +
  `<w:p>${DELETED_DRAWING}</w:p></w:hdr>`

const HEADER_PART = {
  extraRels:
    '<Relationship Id="rId20" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>',
  sectPrExtra: '<w:headerReference w:type="default" r:id="rId20"/>',
  extraParts: [
    {
      path: 'word/header1.xml',
      xml: HEADER_XML,
      contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml',
    },
  ],
}

const TXBX_PARAGRAPH =
  '<w:p><w:r><w:drawing><wp:anchor behindDoc="0" simplePos="0" locked="0" layoutInCell="1" allowOverlap="1">' +
  '<wp:simplePos x="0" y="0"/><wp:extent cx="6038850" cy="1704975"/>' +
  '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
  '<wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
  '<wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="6038850" cy="1704975"/></a:xfrm>' +
  '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr>' +
  '<wps:txbx><w:txbxContent><w:p><w:r><w:t xml:space="preserve">Box </w:t></w:r>' +
  `<w:del${ANN}><w:r><w:delText>old</w:delText></w:r></w:del>` +
  `<w:ins${BOB}><w:r><w:t>new</w:t></w:r></w:ins></w:p></w:txbxContent></wps:txbx></wps:wsp>` +
  '</a:graphicData></a:graphic></wp:anchor></w:drawing></w:r></w:p>'

const partOf = async (bytes: Uint8Array, path: string) =>
  (await JSZip.loadAsync(bytes)).file(path)!.async('string')

describe('resolutions replayed on a part', () => {
  it('a w:id filter picks one change and leaves the other pending', () => {
    const out = applyRevisionResolutions(HEADER_XML, [{ mode: 'accept', filter: { ids: ['7'] } }])
    expect(out).not.toContain('gone')
    expect(out).not.toContain('<w:drawing>')
    expect(out).toMatch(/<w:ins w:id="8"[^>]*><w:r><w:t>added<\/w:t>/)
    expect(countTrackedChangesXml(out)).toBe(1)
  })

  it('an exact date filter and an unfiltered pass', () => {
    const rejected = applyRevisionResolutions(HEADER_XML, [
      { mode: 'reject', filter: { author: 'bob', date: '2026-09-30T08:00:00Z' } },
    ])
    expect(rejected).not.toContain('added')
    expect(rejected).toContain('<w:delText>gone</w:delText>')
    const all = applyRevisionResolutions(HEADER_XML, [{ mode: 'reject' }])
    expect(countTrackedChangesXml(all)).toBe(0)
    expect(all).toContain('<w:t>gone</w:t>')
    expect(all).toContain('<w:drawing>')
  })
})

describe('footnote revisions', () => {
  const open = async () =>
    parseDocx(
      await buildDocx({
        bodyXml: '<w:p><w:r><w:t>Body</w:t></w:r><w:r><w:footnoteReference w:id="2"/></w:r></w:p>',
        ...FOOTNOTES_PART,
      }),
    )
  const blocksOf = (parsed: Awaited<ReturnType<typeof parseDocx>>): SaveBlock[] => [
    { kind: 'original', docxIndex: parsed.blocks[0].docxIndex! },
  ]

  it('accept all leaves no markup in footnotes.xml and keeps the entry bytes otherwise', async () => {
    const parsed = await open()
    const note = parsed.footnotes[0]
    const saved = await saveDocx(parsed, blocksOf(parsed), {
      footnotes: [
        {
          ...note,
          richParas: [[{ text: 'Kept ' }, { text: 'added' }]],
          revisionsResolved: [{ mode: 'accept' }],
        },
      ],
    })
    const xml = await partOf(saved, 'word/footnotes.xml')
    expect(countTrackedChangesXml(xml)).toBe(0)
    expect(xml).toContain('<w:r><w:t>added</w:t></w:r>')
    expect(xml).not.toContain('gone')
    expect(xml).toContain('<w:footnote w:type="separator"')
    expect((await parseDocx(saved)).footnotes[0].text).toBe('Kept added')
  })

  it('reject of one change by id restores the deleted text and keeps the other pending', async () => {
    const parsed = await open()
    const note = parsed.footnotes[0]
    const saved = await saveDocx(parsed, blocksOf(parsed), {
      footnotes: [
        {
          ...note,
          text: 'Kept goneadded',
          revisionsResolved: [{ mode: 'reject', filter: { ids: ['7'] } }],
        },
      ],
    })
    const xml = await partOf(saved, 'word/footnotes.xml')
    expect(xml).toContain('<w:r><w:t>gone</w:t></w:r>')
    expect(xml).toMatch(/<w:ins w:id="8"/)
    const again = await parseDocx(saved)
    expect(again.footnotes[0].text).toBe('Kept goneadded')
    expect(again.footnotes[0].richParas![0].map((r) => !!r.ins)).toEqual([false, true])
  })
})

describe('header revisions', () => {
  it('keepContent replays the resolutions on the part bytes; the deleted drawing goes only on accept', async () => {
    const parsed = await parseDocx(
      await buildDocx({ bodyXml: '<w:p><w:r><w:t>Body</w:t></w:r></w:p>', ...HEADER_PART }),
    )
    const blocks: SaveBlock[] = [{ kind: 'original', docxIndex: parsed.blocks[0].docxIndex! }]
    const paras = parsed.headerParas!
    const accepted = await saveDocx(parsed, blocks, {
      header: {
        text: 'Kept added',
        paras,
        keepContent: true,
        revisionsResolved: [{ mode: 'accept', filter: { ids: ['8'] } }],
      },
    })
    let xml = await partOf(accepted, 'word/header1.xml')
    expect(xml).toContain('<w:r><w:t>added</w:t></w:r>')
    expect(xml).toMatch(/<w:del w:id="7"/)
    expect(xml).toContain('<w:drawing>')
    expect(countTrackedChangesXml(xml)).toBe(2)

    const all = await saveDocx(parsed, blocks, {
      header: {
        text: 'Kept added',
        paras,
        keepContent: true,
        revisionsResolved: [{ mode: 'accept' }],
      },
    })
    xml = await partOf(all, 'word/header1.xml')
    expect(countTrackedChangesXml(xml)).toBe(0)
    expect(xml).not.toContain('<w:drawing>')
    expect(xml).not.toContain('gone')

    const rejected = await saveDocx(parsed, blocks, {
      header: {
        text: 'Kept gone',
        paras,
        keepContent: true,
        revisionsResolved: [{ mode: 'reject' }],
      },
    })
    xml = await partOf(rejected, 'word/header1.xml')
    expect(countTrackedChangesXml(xml)).toBe(0)
    expect(xml).toContain('<w:t>gone</w:t>')
    expect(xml).toContain('<w:drawing>')
    expect(xml).not.toContain('added')
  })

  it('a regenerated part (typed edit after a resolution) writes the resolved runs', async () => {
    const parsed = await parseDocx(
      await buildDocx({ bodyXml: '<w:p><w:r><w:t>Body</w:t></w:r></w:p>', ...HEADER_PART }),
    )
    const blocks: SaveBlock[] = [{ kind: 'original', docxIndex: parsed.blocks[0].docxIndex! }]
    const para = parsed.headerParas![0]
    const saved = await saveDocx(parsed, blocks, {
      header: {
        text: 'Kept added typed',
        paras: [{ ...para, runs: [{ text: 'Kept ' }, { text: 'added' }, { text: ' typed' }] }],
        revisionsResolved: [{ mode: 'accept' }],
      },
    })
    const xml = await partOf(saved, 'word/header1.xml')
    expect(countTrackedChangesXml(xml)).toBe(0)
    expect(xml).not.toContain('gone')
    expect(xml).toContain('typed')
  })
})

describe('textbox revisions', () => {
  it('resolved runs patched into w:txbxContent leave no markup', async () => {
    const parsed = await parseDocx(await buildDocx({ bodyXml: TXBX_PARAGRAPH }))
    const box = parsed.blocks[0].textboxes![0]
    const accepted = patchTextboxParas(parsed.blocks[0].originalXml!, {
      byIndex: [[{ runs: [{ text: 'Box ' }, { text: 'new' }], align: null, numPr: null }]],
    })
    expect(countTrackedChangesXml(accepted)).toBe(0)
    expect(accepted).toContain('new')
    expect(accepted).not.toContain('old')
    const rejected = patchTextboxParas(parsed.blocks[0].originalXml!, {
      byIndex: [
        [
          {
            runs: [{ text: 'Box ' }, { text: 'old' }],
            align: box.paras[0].align ?? null,
            numPr: null,
          },
        ],
      ],
    })
    expect(countTrackedChangesXml(rejected)).toBe(0)
    expect(rejected).toContain('<w:t xml:space="preserve">old</w:t>')
    expect(rejected).not.toContain('new')
  })
})
