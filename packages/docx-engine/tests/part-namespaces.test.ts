/**
 * Regenerated parts splice original entry bytes back in, so their root has to keep declaring
 * every prefix those bytes use. A literal namespace list leaves prefixes unbound, which Word
 * reports as unreadable content. fast-xml-parser's validator does not catch this: it is not
 * namespace-aware, so the check here has to be.
 */
import { describe, expect, it } from 'vitest'
import { buildNotesXml, rootAttributes } from '../src/notes'

/** every prefix used on an element or attribute must be declared on the root */
function unboundPrefixes(xml: string): string[] {
  const root = /<([a-zA-Z][\w.-]*:[\w.-]+)\b([^>]*)>/.exec(xml.replace(/<\?[\s\S]*?\?>/, ''))
  const declared = new Set(
    [...(root?.[2] ?? '').matchAll(/xmlns:([\w.-]+)=/g)].map((m) => m[1]!).concat('xml'),
  )
  const used = new Set<string>()
  for (const m of xml.matchAll(/<\/?([a-zA-Z][\w.-]*):/g)) used.add(m[1]!)
  for (const m of xml.matchAll(/[\s"']([a-zA-Z][\w.-]*):[\w.-]+=["']/g)) used.add(m[1]!)
  return [...used].filter((p) => !declared.has(p) && p !== 'xmlns')
}

describe('regenerated part namespaces', () => {
  it('the check itself catches an unbound prefix', () => {
    expect(unboundPrefixes('<w:x xmlns:w="w"><w:p w14:paraId="1"/></w:x>')).toEqual(['w14'])
    expect(unboundPrefixes('<w:x xmlns:w="w" xmlns:w14="b"><w:p w14:paraId="1"/></w:x>')).toEqual(
      [],
    )
  })

  it('keeps the original root attributes when a part is regenerated', () => {
    const original =
      '<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"' +
      ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"' +
      ' xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">' +
      '<w:footnote w:id="2"><w:p w14:paraId="4A2B1C0D" w14:textId="77BB0011">' +
      '<w:r><w:t>original</w:t></w:r></w:p></w:footnote></w:footnotes>'
    const out = buildNotesXml(
      'footnote',
      [
        { id: '2', text: 'original' },
        { id: '3', text: 'added' },
      ],
      original,
    )
    expect(out).toContain('w14:paraId="4A2B1C0D"')
    expect(unboundPrefixes(out)).toEqual([])
  })

  it('falls back to the literal namespaces when there is no original part', () => {
    const out = buildNotesXml('footnote', [{ id: '2', text: 'fresh' }], null)
    expect(unboundPrefixes(out)).toEqual([])
    expect(out).toContain('xmlns:w=')
  })

  it('rootAttributes ignores a root that declares no namespaces', () => {
    expect(rootAttributes('<w:comments>', 'w:comments', 'FALLBACK')).toBe('FALLBACK')
    expect(rootAttributes(null, 'w:comments', 'FALLBACK')).toBe('FALLBACK')
    expect(rootAttributes('<w:comments xmlns:w="a" xmlns:r="b">', 'w:comments', 'FALLBACK')).toBe(
      'xmlns:w="a" xmlns:r="b"',
    )
  })
})
