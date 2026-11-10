import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { parseDocx, saveDocx } from '../src/index'
import { buildDocx } from './helpers/build-docx'

/**
 * fast-xml-parser resolves the five named XML entities on its own, so text read off
 * the parse tree only needs numeric character references resolved on top of it.
 * Decoding the named entities a second time rewrites a document's own literal
 * "&lt;" / "&amp;" text into "<" / "&", and regenerating the paragraph then persists
 * that damage to word/document.xml.
 */
describe('entity decoding in run text', () => {
  // XML 1.0 stores literal entity text with the ampersand escaped; Word displays
  // this run as: Escape &amp; as &lt;b&gt; ok
  const STORED = 'Escape &amp;amp; as &amp;lt;b&amp;gt; ok'
  const DISPLAYED = 'Escape &amp; as &lt;b&gt; ok'

  it('keeps text that is literally an entity reference', async () => {
    const bytes = await buildDocx({
      bodyXml: `<w:p><w:r><w:t xml:space="preserve">${STORED}</w:t></w:r></w:p>`,
    })
    const doc = await parseDocx(bytes)
    expect(doc.blocks[0].runs!.map((r) => r.text).join('')).toBe(DISPLAYED)
  })

  it('still resolves a real numeric character reference beside literal entity text', async () => {
    const bytes = await buildDocx({
      bodyXml: '<w:p><w:r><w:t xml:space="preserve">&amp;lt; then &#x2713;</w:t></w:r></w:p>',
    })
    const doc = await parseDocx(bytes)
    expect(doc.blocks[0].runs!.map((r) => r.text).join('')).toBe('&lt; then ✓')
  })

  it('does not rewrite an untouched run when a sibling run is edited', async () => {
    const bytes = await buildDocx({
      bodyXml:
        `<w:p><w:r><w:t xml:space="preserve">${STORED} </w:t></w:r>` +
        '<w:r><w:rPr><w:b/></w:rPr><w:t>tail</w:t></w:r></w:p>',
    })
    const doc = await parseDocx(bytes)
    const runs = doc.blocks[0].runs!.map((r) => (r.bold ? { ...r, text: 'edited' } : r))
    const saved = await saveDocx(doc, [{ kind: 'generated', block: { type: 'paragraph', runs } }])
    const zip = await JSZip.loadAsync(saved)
    const xml = await zip.file('word/document.xml')!.async('string')
    expect(xml).toContain(`${STORED} `)
    expect(xml).toContain('edited')
  })
})
