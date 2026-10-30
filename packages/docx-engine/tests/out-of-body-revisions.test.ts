import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { parseDocx, saveDocx, type SaveBlock } from '../src/index'
import { buildNotesXml } from '../src/notes'
import { buildDocx } from './helpers/build-docx'

/** w:ins / w:del outside the body: footnotes, textboxes, header parts */

const REV = ' w:id="7" w:author="Ann" w:date="2026-09-29T12:23:00Z"'
const NS = ' xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'

const FOOTNOTES_XML =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  `<w:footnotes${NS}>` +
  '<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>' +
  '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>' +
  '<w:footnote w:id="2"><w:p><w:r><w:footnoteRef/></w:r><w:r><w:t xml:space="preserve"> Kept </w:t></w:r>' +
  `<w:del${REV}><w:r><w:delText>gone</w:delText></w:r></w:del>` +
  `<w:ins${REV}><w:r><w:t>added</w:t></w:r></w:ins></w:p></w:footnote>` +
  '</w:footnotes>'

const TXBX =
  '<w:txbxContent><w:p><w:r><w:t xml:space="preserve">Box </w:t></w:r>' +
  `<w:del${REV}><w:r><w:delText>old</w:delText></w:r></w:del>` +
  `<w:ins${REV}><w:r><w:t>new</w:t></w:r></w:ins></w:p></w:txbxContent>`

const TEXTBOX_PARAGRAPH =
  '<w:p><w:r><w:drawing><wp:anchor behindDoc="0" simplePos="0" locked="0" layoutInCell="1" allowOverlap="1">' +
  '<wp:simplePos x="0" y="0"/><wp:extent cx="6038850" cy="1704975"/>' +
  '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
  '<wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
  '<wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="6038850" cy="1704975"/></a:xfrm>' +
  '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr>' +
  `<wps:txbx>${TXBX}</wps:txbx></wps:wsp>` +
  '</a:graphicData></a:graphic></wp:anchor></w:drawing></w:r></w:p>'

const HEADER_XML =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr${NS}>` +
  '<w:p><w:r><w:t xml:space="preserve">Kept </w:t></w:r>' +
  `<w:del${REV}><w:r><w:delText>gone</w:delText></w:r></w:del>` +
  `<w:ins${REV}><w:r><w:t>added</w:t></w:r></w:ins></w:p></w:hdr>`

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

describe('tracked changes outside the body', () => {
  it('footnote runs carry ins / del instead of flattening w:delText into live text', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml: '<w:p><w:r><w:t>Body</w:t></w:r><w:r><w:footnoteReference w:id="2"/></w:r></w:p>',
        extraRels:
          '<Relationship Id="rId40" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>',
        extraParts: [
          {
            path: 'word/footnotes.xml',
            xml: FOOTNOTES_XML,
            contentType:
              'application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml',
          },
        ],
      }),
    )
    const runs = doc.footnotes[0].richParas![0]
    expect(runs.map((r) => [r.text, !!r.del, !!r.ins])).toEqual([
      ['Kept ', false, false],
      ['gone', true, false],
      ['added', false, true],
    ])
    expect(runs[1].del).toMatchObject({ author: 'Ann', date: '2026-09-29T12:23:00Z' })
    expect(doc.footnotes[0].text).toBe('Kept added')
  })

  it('a rich note rebuild shares one w:del / w:ins wrapper per tracked change, ids unique', () => {
    const rev = { author: 'Ann', date: '2026-09-29T12:23:00Z', id: '7' }
    const xml = buildNotesXml(
      'footnote',
      [
        {
          id: '2',
          text: 'Kept added',
          richParas: [
            [
              { text: 'Kept ' },
              { text: 'go', del: rev },
              { text: 'ne', del: rev, bold: true },
              { text: 'added', ins: rev },
            ],
            [{ text: 'second', del: rev }],
          ],
        },
        { id: '3', text: 'other', richParas: [[{ text: 'other', ins: rev }]] },
      ],
      null,
    )
    expect(xml.match(/<w:del /g)).toHaveLength(2)
    expect(xml).toContain('<w:delText xml:space="preserve">go</w:delText></w:r><w:r><w:rPr><w:b/>')
    const ids = [...xml.matchAll(/<w:(?:ins|del) w:id="(\d+)"/g)].map((m) => m[1])
    expect(new Set(ids).size).toBe(ids.length)
  })

  it('textbox runs keep their revision info', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: TEXTBOX_PARAGRAPH }))
    const runs = doc.blocks[0].textboxes![0].paras[0].runs
    expect(runs.map((r) => [r.text, !!r.del, !!r.ins])).toEqual([
      ['Box ', false, false],
      ['old', true, false],
      ['new', false, true],
    ])
  })

  it('header del / ins runs survive a part regeneration as w:del / w:delText and w:ins', async () => {
    const parsed = await parseDocx(
      await buildDocx({ bodyXml: '<w:p><w:r><w:t>Body</w:t></w:r></w:p>', ...HEADER_PART }),
    )
    const paras = parsed.headerParas!
    expect(paras[0].runs.map((r) => [r.text, !!r.del, !!r.ins])).toEqual([
      ['Kept ', false, false],
      ['gone', true, false],
      ['added', false, true],
    ])
    expect(parsed.headerText).toBe('Kept added')

    const saveBlocks: SaveBlock[] = [{ kind: 'original', docxIndex: parsed.blocks[0].docxIndex! }]
    const saved = await saveDocx(parsed, saveBlocks, {
      header: {
        text: 'Kept goneadded',
        paras: [{ ...paras[0], runs: [...paras[0].runs, { text: ' typed' }] }],
      },
    })
    const xml = await (await JSZip.loadAsync(saved)).file('word/header1.xml')!.async('string')
    expect(xml).toMatch(/<w:del w:id="\d+" w:author="Ann" w:date="2026-09-29T12:23:00Z">/)
    expect(xml).toContain('<w:delText xml:space="preserve">gone</w:delText>')
    expect(xml).toMatch(/<w:ins w:id="\d+" w:author="Ann"[^>]*><w:r><w:t[^>]*>added<\/w:t>/)
    expect(xml).not.toContain('<w:t xml:space="preserve">gone</w:t>')

    const reparsed = await parseDocx(saved)
    expect(reparsed.headerParas![0].runs.map((r) => [r.text, !!r.del, !!r.ins])).toEqual([
      ['Kept ', false, false],
      ['gone', true, false],
      ['added', false, true],
      [' typed', false, false],
    ])
  })
})
