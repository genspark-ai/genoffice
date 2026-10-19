import { describe, expect, it } from 'vitest'
import { applyTrackedChangesXml, countTrackedChangesXml } from '../src/revision-xml'

const AT = 'w:author="X" w:date="2026-01-01T00:00:00Z"'
const r = (t: string) => `<w:r><w:t xml:space="preserve">${t}</w:t></w:r>`
const dr = (t: string) => `<w:r><w:delText xml:space="preserve">${t}</w:delText></w:r>`
const p = (inner: string, pPr = '') => `<w:p>${pPr}${inner}</w:p>`
const hdr = (body: string) => `<w:hdr xmlns:w="w">${body}</w:hdr>`
const texts = (xml: string) =>
  [...xml.matchAll(/<w:p>[\s\S]*?<\/w:p>/g)].map((m) =>
    [...m[0].matchAll(/<w:(?:t|delText)[^>]*>([^<]*)</g)].map((t) => t[1]).join(''),
  )

describe('applyTrackedChangesXml', () => {
  it('accepts a move: text leaves the source and stays at the destination', () => {
    const xml = hdr(
      p(
        r('P1 alpha ') +
          `<w:moveFromRangeStart w:id="1" ${AT} w:name="m"/><w:moveFrom w:id="2" ${AT}>${dr('MOVED ')}</w:moveFrom><w:moveFromRangeEnd w:id="1"/>` +
          r('beta'),
      ) +
        p(
          r('P2 gamma ') +
            `<w:moveToRangeStart w:id="3" ${AT} w:name="m"/><w:moveTo w:id="4" ${AT}>${r('MOVED ')}</w:moveTo><w:moveToRangeEnd w:id="3"/>` +
            r('delta'),
        ),
    )
    const out = applyTrackedChangesXml(xml, 'accept')
    expect(out.applied).toBe(2)
    expect(out.remaining).toBe(0)
    expect(texts(out.xml)).toEqual(['P1 alpha beta', 'P2 gamma MOVED delta'])
    expect(out.xml).not.toMatch(/moveFrom|moveTo/)
    const back = applyTrackedChangesXml(xml, 'reject')
    expect(texts(back.xml)).toEqual(['P1 alpha MOVED beta', 'P2 gamma delta'])
    expect(back.xml).not.toMatch(/delText|moveFrom|moveTo/)
  })

  it('accepts ins / del and nested ins+del; reject restores deleted text', () => {
    const xml = hdr(
      p(
        r('a ') +
          `<w:ins w:id="1" ${AT}>${r('b ')}</w:ins>` +
          `<w:del w:id="2" ${AT}>${dr('c ')}</w:del>` +
          `<w:ins w:id="3" ${AT}><w:del w:id="4" ${AT}>${dr('d ')}</w:del></w:ins>` +
          r('e'),
      ),
    )
    const acc = applyTrackedChangesXml(xml, 'accept')
    expect(texts(acc.xml)).toEqual(['a b e'])
    expect(acc.remaining).toBe(0)
    const rej = applyTrackedChangesXml(xml, 'reject')
    expect(texts(rej.xml)).toEqual(['a c e'])
    expect(rej.xml).not.toMatch(/<w:ins|<w:del|delText/)
  })

  it('a deleted paragraph mark joins the paragraph into the next one on accept', () => {
    const xml = hdr(
      p(
        r('one'),
        `<w:pPr><w:jc w:val="center"/><w:rPr><w:del w:id="1" ${AT}/><w:b/></w:rPr></w:pPr>`,
      ) +
        p(r('two'), '<w:pPr><w:jc w:val="right"/></w:pPr>') +
        p(r('three')),
    )
    const acc = applyTrackedChangesXml(xml, 'accept')
    expect(texts(acc.xml)).toEqual(['onetwo', 'three'])
    expect(acc.xml).toContain('<w:jc w:val="center"/>')
    expect(acc.xml).not.toContain('w:val="right"')
    expect(acc.xml).not.toMatch(/<w:del\b/)
    const rej = applyTrackedChangesXml(xml, 'reject')
    expect(texts(rej.xml)).toEqual(['one', 'two', 'three'])
    expect(rej.xml).toContain('<w:rPr><w:b/></w:rPr>')
  })

  it('an inserted paragraph mark stays on accept and joins on reject; the last paragraph only loses the marker', () => {
    const xml = hdr(
      p(r('one'), `<w:pPr><w:rPr><w:ins w:id="1" ${AT}/></w:rPr></w:pPr>`) +
        p(r('two'), `<w:pPr><w:rPr><w:del w:id="2" ${AT}/></w:rPr></w:pPr>`),
    )
    const acc = applyTrackedChangesXml(xml, 'accept')
    expect(texts(acc.xml)).toEqual(['one', 'two'])
    expect(acc.xml).not.toMatch(/<w:ins\b|<w:del\b|<w:rPr>/)
    const rej = applyTrackedChangesXml(xml, 'reject')
    expect(texts(rej.xml)).toEqual(['onetwo'])
  })

  it('formatting records: accept drops them, reject restores the old properties', () => {
    const xml = hdr(
      p(
        `<w:r><w:rPr><w:b/><w:rPrChange w:id="1" ${AT}><w:rPr><w:i/></w:rPr></w:rPrChange></w:rPr><w:t>x</w:t></w:r>`,
        `<w:pPr><w:jc w:val="center"/><w:rPr><w:b/></w:rPr><w:pPrChange w:id="2" ${AT}><w:pPr><w:jc w:val="left"/></w:pPr></w:pPrChange></w:pPr>`,
      ),
    )
    const acc = applyTrackedChangesXml(xml, 'accept')
    expect(acc.applied).toBe(2)
    expect(acc.xml).toBe(
      hdr(
        p(
          '<w:r><w:rPr><w:b/></w:rPr><w:t>x</w:t></w:r>',
          '<w:pPr><w:jc w:val="center"/><w:rPr><w:b/></w:rPr></w:pPr>',
        ),
      ),
    )
    const rej = applyTrackedChangesXml(xml, 'reject')
    expect(rej.xml).toBe(
      hdr(
        p(
          '<w:r><w:rPr><w:i/></w:rPr><w:t>x</w:t></w:r>',
          '<w:pPr><w:jc w:val="left"/><w:rPr><w:b/></w:rPr></w:pPr>',
        ),
      ),
    )
  })

  it('a deleted table row goes on accept; an inserted row goes on reject', () => {
    const row = (mark: string, t: string) =>
      `<w:tr>${mark ? `<w:trPr>${mark}</w:trPr>` : ''}<w:tc><w:p>${r(t)}</w:p></w:tc></w:tr>`
    const xml = hdr(
      `<w:tbl>${row('', 'a')}${row(`<w:del w:id="1" ${AT}/>`, 'b')}${row(`<w:ins w:id="2" ${AT}/>`, 'c')}</w:tbl>`,
    )
    const acc = applyTrackedChangesXml(xml, 'accept')
    expect(texts(acc.xml)).toEqual(['a', 'c'])
    expect(acc.xml).not.toContain('<w:trPr>')
    const rej = applyTrackedChangesXml(xml, 'reject')
    expect(texts(rej.xml)).toEqual(['a', 'b'])
  })

  it('a filter leaves other authors’ changes pending and counts what remains', () => {
    const xml = hdr(
      p(
        `<w:ins w:id="1" w:author="A" w:date="2026-01-01T00:00:00Z">${r('a')}</w:ins>` +
          `<w:ins w:id="2" w:author="B" w:date="2026-02-01T00:00:00Z">${r('b')}</w:ins>`,
      ),
    )
    expect(countTrackedChangesXml(xml)).toBe(2)
    const byAuthor = applyTrackedChangesXml(xml, 'accept', { author: 'a' })
    expect(byAuthor.applied).toBe(1)
    expect(byAuthor.remaining).toBe(1)
    expect(byAuthor.xml).toMatch(/<w:ins w:id="2"/)
    expect(byAuthor.xml).not.toMatch(/<w:ins w:id="1"/)
    const byDate = applyTrackedChangesXml(xml, 'reject', { before: '2026-01-15' })
    expect(texts(byDate.xml)).toEqual(['b'])
    const byType = applyTrackedChangesXml(xml, 'accept', { type: 'deletion' })
    expect(byType.applied).toBe(0)
    expect(byType.xml).toBe(xml)
  })

  it('an author filter takes a move’s unattributed end markers along with their start', () => {
    const xml = hdr(
      p(
        `<w:moveFromRangeStart w:id="1" w:author="A" w:date="2026-01-01T00:00:00Z" w:name="m"/>` +
          `<w:moveFrom w:id="2" w:author="A" w:date="2026-01-01T00:00:00Z">${dr('x')}</w:moveFrom>` +
          `<w:moveFromRangeEnd w:id="1"/>` +
          `<w:ins w:id="3" w:author="B" w:date="2026-01-01T00:00:00Z">${r('y')}</w:ins>`,
      ),
    )
    const out = applyTrackedChangesXml(xml, 'accept', { author: 'A' })
    expect(out.applied).toBe(1)
    expect(out.xml).not.toMatch(/moveFrom/)
    expect(out.xml).toMatch(/<w:ins w:id="3"/)
  })

  it('leaves a part without revisions byte-identical', () => {
    const xml = hdr(p(r('plain')))
    expect(applyTrackedChangesXml(xml, 'accept')).toEqual({ xml, applied: 0, remaining: 0 })
  })
})
