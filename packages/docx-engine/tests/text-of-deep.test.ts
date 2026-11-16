import { describe, expect, it } from 'vitest'
import { deepXmlParser, textOf, type XNode } from '../src/xml-utils'

describe('textOf', () => {
  it('walks a tree far deeper than the JS stack could recurse', () => {
    // deepXmlParser legitimately allows 100k nesting levels (POI stress
    // files); fxp builds the tree iteratively, so the walk must not pay a
    // stack frame per level.
    const depth = 20_000
    let xml = '<root>'
    xml += '<a>'.repeat(depth)
    xml += 'deep'
    xml += '</a>'.repeat(depth)
    xml += '</root>'
    const node = (deepXmlParser.parse(xml) as XNode[])[0]!
    expect(textOf(node)).toBe('deep')
  })

  it('keeps text in document order across nesting and siblings', () => {
    const node = (deepXmlParser.parse('<root>x<a>y<b>z</b>w</a>q</root>') as XNode[])[0]!
    expect(textOf(node)).toBe('xyzwq')
  })
})
