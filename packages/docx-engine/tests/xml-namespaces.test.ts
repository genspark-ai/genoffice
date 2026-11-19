import { describe, expect, it } from 'vitest'
import { bindUnboundPrefixes, rootNamespaces } from '../src/xml-namespaces'

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
const DECL = '<?xml version="1.0" encoding="UTF-8"?>'
const A = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
const C = 'xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"'
const chart =
  '<w:drawing><wp:inline><wp:extent cx="1" cy="1"/>' +
  `<a:graphic ${A}><a:graphicData uri="x"><c:chart ${C} r:id="rId1"/></a:graphicData></a:graphic>` +
  '</wp:inline></w:drawing>'
const doc = (root: string, body: string) =>
  `${DECL}<w:document ${root}><w:body><w:p><w:r>${body}</w:r></w:p></w:body></w:document>`

describe('bindUnboundPrefixes', () => {
  it('declares on the root the prefixes a chart uses that nothing declares', () => {
    const out = bindUnboundPrefixes(doc(W, chart))
    expect(out).toContain(
      `<w:document ${W}` +
        ' xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"' +
        ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
    )
    // a: and c: are declared where they are used
    expect(out).not.toMatch(/<w:document[^>]*xmlns:a=/)
    expect(out).not.toMatch(/<w:document[^>]*xmlns:c=/)
  })

  it('leaves a document whose root declares them byte for byte', () => {
    const full = doc(`${W} xmlns:wp="urn:wp" xmlns:r="urn:r"`, chart)
    expect(bindUnboundPrefixes(full)).toBe(full)
  })

  it('leaves a document that only uses a prefix inside its declaration alone', () => {
    const local = doc(W, `<a:graphic ${A}><a:graphicData uri="x"/></a:graphic>`)
    expect(bindUnboundPrefixes(local)).toBe(local)
  })

  it('sees a use after the declaring element has closed', () => {
    const after = doc(W, `<a:graphic ${A}/><a:blip/>`)
    expect(bindUnboundPrefixes(after)).toMatch(/<w:document[^>]*xmlns:a=/)
  })

  it('is not fooled by a prefix-like word in text or an attribute value', () => {
    const text = doc(
      W,
      '<w:t xml:space="preserve">see wp:inline and r:id="x"</w:t><w:x w:val=" r:id"/>',
    )
    expect(bindUnboundPrefixes(text)).toBe(text)
  })

  it('does not take a declaration or a use inside a comment or CDATA for markup', () => {
    const hidden = `<!-- <a:graphic ${A}> --><w:t><![CDATA[<a:x ${A}>]]></w:t><a:blip/>`
    expect(bindUnboundPrefixes(doc(W, hidden))).toMatch(/<w:document[^>]*xmlns:a=/)
    const unused = doc(W, '<!-- <wp:inline/> --><w:t><![CDATA[<wp:inline>]]></w:t>')
    expect(bindUnboundPrefixes(unused)).toBe(unused)
  })

  it('finds the root after a comment that holds a tag', () => {
    const led = `${DECL}<!-- <note> --><w:document ${W}><w:body><wp:inline/></w:body></w:document>`
    expect(bindUnboundPrefixes(led)).toContain(`<w:document ${W} xmlns:wp=`)
  })

  it('handles a self-closed root and a missing declaration line', () => {
    expect(bindUnboundPrefixes(`<w:document ${W}/>`)).toBe(`<w:document ${W}/>`)
    const bare = `<w:document ${W}><w:body><wp:inline/></w:body></w:document>`
    expect(bindUnboundPrefixes(bare)).toMatch(/^<w:document [^>]*xmlns:wp=/)
  })
})

describe('bindUnboundPrefixes on a part the save creates', () => {
  const MC = 'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"'
  const hdr = (root: string, body: string) => `${DECL}<w:hdr ${root}>${body}</w:hdr>`

  it('declares a prefix that only a Requires or Ignorable list names', () => {
    const out = bindUnboundPrefixes(
      hdr(
        `${W} ${MC} mc:Ignorable="w14"`,
        '<mc:AlternateContent><mc:Choice Requires="wps"/></mc:AlternateContent>',
      ),
    )
    expect(out).toContain('xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"')
    expect(out).toContain(
      'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"',
    )
  })

  it('binds an unknown prefix from the part the content came from', () => {
    const body = '<w:p vendor:id="1"/>'
    expect(bindUnboundPrefixes(hdr(W, body))).toBe(hdr(W, body))
    expect(bindUnboundPrefixes(hdr(W, body), { vendor: 'urn:example:vendor' })).toBe(
      hdr(`${W} xmlns:vendor="urn:example:vendor"`, body),
    )
  })

  it('reads the declarations of a root', () => {
    expect(rootNamespaces(hdr(`${W} xmlns:r='urn:r'`, '<w:p xmlns:a="urn:a"/>'))).toEqual({
      w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
      r: 'urn:r',
    })
  })
})
