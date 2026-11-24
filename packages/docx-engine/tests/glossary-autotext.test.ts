import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { docxAutoTextBody, docxAutoTextEntries, parseDocx } from '../src/index'
import { buildDocx } from './helpers/build-docx'

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
const r = (text: string, rPr = '') =>
  `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}<w:t xml:space="preserve">${text}</w:t></w:r>`
const part = (name: string, gallery: string, body: string) =>
  `<w:docPart><w:docPartPr><w:name w:val="${name}"/><w:category><w:name w:val="General"/>` +
  `<w:gallery w:val="${gallery}"/></w:category><w:behaviors><w:behavior w:val="content"/></w:behaviors>` +
  `</w:docPartPr><w:docPartBody>${body}</w:docPartBody></w:docPart>`

const withGlossary = () =>
  buildDocx({
    bodyXml: `<w:p>${r('Body.')}</w:p>`,
    extraRels:
      '<Relationship Id="rId90" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/glossaryDocument" Target="glossary/document.xml"/>',
    extraParts: [
      {
        path: 'word/glossary/document.xml',
        contentType:
          'application/vnd.openxmlformats-officedocument.wordprocessingml.document.glossary+xml',
        xml:
          `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:glossaryDocument ${W}><w:docParts>` +
          part(
            'Sign off',
            'autoTxt',
            `<w:p>${r('Kind regards, ')}${r('R&amp;D', '<w:b/>')}</w:p>`,
          ) +
          part('A placeholder', 'placeholder', `<w:p>${r('Click here')}</w:p>`) +
          part(
            'Address',
            'autoTxt',
            `<w:p>${r('1 Main Street')}</w:p><w:p>${r('Springfield')}</w:p>`,
          ) +
          '</w:docParts></w:glossaryDocument>',
      },
    ],
  })

describe('AutoText in the glossary part', () => {
  it('lists the entries of the AutoText gallery only', async () => {
    expect(await docxAutoTextEntries(await withGlossary())).toEqual([
      { name: 'Sign off', text: 'Kind regards, R&D' },
      { name: 'Address', text: '1 Main Street\nSpringfield' },
    ])
    expect(await docxAutoTextEntries(await buildDocx({ bodyXml: '<w:p/>' }))).toEqual([])
  })

  it('gives one entry as a document the parse reads, with its formatting', async () => {
    const bytes = await withGlossary()
    const one = await docxAutoTextBody(bytes, 'Sign off')
    expect(one?.inline).toBe(true)
    const parsed = await parseDocx(one!.bytes)
    const runs = parsed.blocks.flatMap((b) => ('runs' in b ? (b.runs ?? []) : []))
    expect(runs.map((run) => run.text).join('')).toBe('Kind regards, R&D')
    expect(runs.find((run) => run.text === 'R&D')?.bold).toBe(true)

    const two = await docxAutoTextBody(bytes, 'Address')
    expect(two?.inline).toBe(false)
    expect((await parseDocx(two!.bytes)).blocks.filter((b) => 'runs' in b)).toHaveLength(2)
    expect(await docxAutoTextBody(bytes, 'A placeholder')).toBeNull()
  })

  it('keeps a dollar sign of a relationship target as it is written', async () => {
    const bytes = await buildDocx({
      bodyXml: `<w:p>${r('Body.')}</w:p>`,
      extraRels:
        '<Relationship Id="rId90" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/glossaryDocument" Target="glossary/document.xml"/>',
      extraParts: [
        {
          path: 'word/glossary/document.xml',
          contentType:
            'application/vnd.openxmlformats-officedocument.wordprocessingml.document.glossary+xml',
          xml:
            `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:glossaryDocument ${W}><w:docParts>` +
            part('Logo', 'autoTxt', `<w:p>${r('Logo')}</w:p>`) +
            '</w:docParts></w:glossaryDocument>',
        },
        {
          path: 'word/glossary/_rels/document.xml.rels',
          contentType: 'application/vnd.openxmlformats-package.relationships+xml',
          xml:
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/a$&amp;b$$.png"/>' +
            '</Relationships>',
        },
      ],
    })
    const one = await docxAutoTextBody(bytes, 'Logo')
    const rels = await (
      await JSZip.loadAsync(one!.bytes)
    )
      .file('word/_rels/document.xml.rels')!
      .async('string')
    expect(rels).toContain('Target="glossary/media/a$&amp;b$$.png"')
  })
})
