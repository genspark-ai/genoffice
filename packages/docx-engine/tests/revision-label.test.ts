import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import {
  compareDocuments,
  labelUnattributedRevisions,
  labelUnattributedRevisionsXml,
  parseDocx,
} from '../src/index'
import { buildDocx } from './helpers/build-docx'

const r = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`
const p = (content: string) => `<w:p>${content}</w:p>`

describe('tracked changes that name no author', () => {
  it('take the name, an empty or a missing attribute alike, and a named one stays', () => {
    const xml =
      '<w:ins w:id="1" w:author="" w:date="2026-01-01T00:00:00Z"><w:r/></w:ins>' +
      '<w:del w:id="2"><w:r/></w:del>' +
      '<w:ins w:id="3" w:author="Abe"><w:r/></w:ins>' +
      '<w:rPr><w:ins w:id="4"/></w:rPr>' +
      '<w:pPrChange w:id="5" w:author=""><w:pPr/></w:pPrChange>' +
      '<w:insideH w:val="single"/><w:delText>x</w:delText>'
    const out = labelUnattributedRevisionsXml(xml, 'O "Riginal" & Co')
    expect(out.count).toBe(4)
    const name = 'w:author="O &quot;Riginal&quot; &amp; Co"'
    expect(out.xml).toBe(
      `<w:ins w:id="1" ${name} w:date="2026-01-01T00:00:00Z"><w:r/></w:ins>` +
        `<w:del w:id="2" ${name}><w:r/></w:del>` +
        '<w:ins w:id="3" w:author="Abe"><w:r/></w:ins>' +
        `<w:rPr><w:ins w:id="4" ${name}/></w:rPr>` +
        `<w:pPrChange w:id="5" ${name}><w:pPr/></w:pPrChange>` +
        '<w:insideH w:val="single"/><w:delText>x</w:delText>',
    )
  })

  it('a package with none comes back as it was', async () => {
    const bytes = await buildDocx({ bodyXml: p(r('Plain.')) })
    const out = await labelUnattributedRevisions(bytes, 'Someone')
    expect(out.count).toBe(0)
    expect(out.bytes).toBe(bytes)
  })

  it('a combination carries them under the name given for the original', async () => {
    const original = await buildDocx({
      bodyXml:
        p(r('Kept as it is.')) +
        p(
          r('The plan ') +
            '<w:ins w:id="7" w:date="2026-11-01T08:00:00Z">' +
            r('quietly ') +
            '</w:ins>' +
            '<w:ins w:id="8" w:author="Abe Origin" w:date="2026-11-01T08:00:00Z">' +
            r('and slowly ') +
            '</w:ins>' +
            r('moves on.'),
        ),
    })
    const revised = await buildDocx({
      bodyXml: p(r('Kept as it is.')) + p(r('The plan moves on today.')),
    })
    const named = await labelUnattributedRevisions(original, 'Olga Original')
    expect(named.count).toBe(1)
    const result = await compareDocuments(await parseDocx(named.bytes), await parseDocx(revised), {
      author: 'Rae Visor',
      date: '2026-11-23T00:00:00Z',
      keepRevisions: true,
    })
    const xml = await (
      await JSZip.loadAsync(result.bytes)
    )
      .file('word/document.xml')!
      .async('string')
    const authorOf = (text: string) =>
      new RegExp(`<w:ins [^>]*w:author="([^"]*)"[^>]*>(?:(?!</w:ins>).)*${text}`).exec(xml)?.[1]
    expect(authorOf('quietly')).toBe('Olga Original')
    expect(authorOf('and slowly')).toBe('Abe Origin')
    expect(authorOf('today')).toBe('Rae Visor')
  })
})
