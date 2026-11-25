/**
 * A header or footer paragraph keeps its Header / Footer style through a save,
 * and a document that has no such style gets Word's built-in definition when
 * a new part names it: Word for Mac 16 writes the style on first use with a
 * center stop at 4680 and a right stop at 9360 twips, no space after and
 * single spacing (probe 2026-11-20).
 */
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { BUILT_IN_HF_TAB_STOPS, parseDocx, saveDocx, type SaveBlock } from '../src/index'
import { buildDocx } from './helpers/build-docx'

const HEADER_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml'
const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
const OWN_HEADER_STYLE =
  '<w:style w:type="paragraph" w:styleId="Header"><w:name w:val="header"/><w:basedOn w:val="Normal"/>' +
  '<w:pPr><w:tabs><w:tab w:val="center" w:pos="4513"/><w:tab w:val="right" w:pos="9026"/></w:tabs>' +
  '<w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:style>'

async function docWithHeader(pPr: string, extraStylesXml = OWN_HEADER_STYLE) {
  return parseDocx(
    await buildDocx({
      bodyXml: '<w:p><w:r><w:t>body</w:t></w:r></w:p>',
      extraStylesXml,
      extraRels:
        '<Relationship Id="rId60" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>',
      extraParts: [
        {
          path: 'word/header1.xml',
          contentType: HEADER_TYPE,
          xml: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr ${W}><w:p><w:pPr>${pPr}</w:pPr><w:r><w:t>Left</w:t></w:r></w:p></w:hdr>`,
        },
      ],
      sectPrExtra: '<w:headerReference w:type="default" r:id="rId60"/>',
    }),
  )
}

async function partsOf(saved: Uint8Array) {
  const zip = await JSZip.loadAsync(saved)
  const read = async (p: string) => (await zip.file(p)?.async('string')) ?? ''
  return {
    header: await read('word/header1.xml'),
    footer: await read('word/footer1.xml'),
    styles: await read('word/styles.xml'),
  }
}

const body = (doc: Awaited<ReturnType<typeof parseDocx>>): SaveBlock[] => [
  { kind: 'original', docxIndex: doc.blocks[0].docxIndex! },
]

describe('header and footer paragraph styles', () => {
  it('reads the paragraph style of a strip line, only when the document defines it', async () => {
    const own = await docWithHeader('<w:pStyle w:val="Header"/>')
    expect(own.headerParas![0].styleId).toBe('Header')
    const missing = await docWithHeader('<w:pStyle w:val="Header"/>', '')
    expect(missing.headerParas![0].styleId).toBeUndefined()
  })

  it('writes the style back when the line is edited', async () => {
    const doc = await docWithHeader('<w:pStyle w:val="Header"/>')
    const para = doc.headerParas![0]
    const saved = await saveDocx(doc, body(doc), {
      header: { text: '', paras: [{ ...para, runs: [{ text: 'Edited' }] }] },
    })
    const { header, styles } = await partsOf(saved)
    expect(header).toMatch(/<w:p><w:pPr><w:pStyle w:val="Header"\/>/)
    expect(header).toContain('Edited')
    // the document's own definition stays as it is
    expect(styles).toContain('<w:tab w:val="center" w:pos="4513"/>')
    expect(styles.match(/w:styleId="Header"/g)).toHaveLength(1)
  })

  it("adds Word's built-in Header and Footer to a document that has neither", async () => {
    const doc = await parseDocx(
      await buildDocx({ bodyXml: '<w:p><w:r><w:t>body</w:t></w:r></w:p>' }),
    )
    const line = (styleId: string, text: string) => ({
      styleId,
      tabStops: [...BUILT_IN_HF_TAB_STOPS],
      spaceAfter: 0,
      runs: [{ text }],
    })
    const saved = await saveDocx(doc, body(doc), {
      header: { text: '', paras: [line('Header', 'Top')] },
      footer: { text: '', paras: [line('Footer', 'Bottom')] },
    })
    const { header, footer, styles } = await partsOf(saved)
    expect(header).toMatch(/<w:pStyle w:val="Header"\/>/)
    expect(footer).toMatch(/<w:pStyle w:val="Footer"\/>/)
    for (const [id, name] of [
      ['Header', 'header'],
      ['Footer', 'footer'],
    ]) {
      const def = new RegExp(`<w:style w:type="paragraph" w:styleId="${id}">.*?</w:style>`).exec(
        styles,
      )?.[0]
      expect(def).toContain(`<w:name w:val="${name}"/>`)
      expect(def).toContain('<w:basedOn w:val="Normal"/>')
      expect(def).toContain(
        '<w:tab w:val="center" w:pos="4680"/><w:tab w:val="right" w:pos="9360"/>',
      )
      expect(def).toContain('w:after="0"')
      expect(def).not.toContain('w:customStyle')
      expect(def).not.toContain('<w:qFormat/>')
    }
    // the reopened document resolves the new styles and keeps the reference
    const reopened = await parseDocx(saved)
    expect(reopened.styles.get('Header')?.display?.tabStops).toEqual([
      { pos: 4680, val: 'center' },
      { pos: 9360, val: 'right' },
    ])
    expect(reopened.headerParas![0].styleId).toBe('Header')
  })

  it('leaves styles.xml alone for a part without a style reference', async () => {
    const doc = await parseDocx(
      await buildDocx({ bodyXml: '<w:p><w:r><w:t>body</w:t></w:r></w:p>' }),
    )
    const saved = await saveDocx(doc, body(doc), {
      header: { text: '', paras: [{ runs: [{ text: 'Plain' }] }] },
    })
    const { header, styles } = await partsOf(saved)
    expect(header).not.toContain('w:pStyle')
    expect(styles).not.toContain('w:styleId="Header"')
  })
})
