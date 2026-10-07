import { describe, expect, it } from 'vitest'
import { mergeLevelXml } from '../src/blank'

/**
 * `CT_PPrGeneral` orders its children, and Word validates that order strictly.
 * A producer can put any paragraph child in a `w:lvl`'s `w:pPr` — `w:spacing`
 * and `w:jc` are the realistic ones — and the merge has to emit them in schema
 * order, not just park them around the two children this function writes.
 */
describe('mergeLevelXml keeps w:pPr children in schema order', () => {
  const LEVEL = { numFmt: 'bullet', lvlText: '–', indentLeft: 720 }

  /** the DIRECT child order inside the emitted <w:pPr> (w:tab nests in w:tabs) */
  const pPrOrder = (xml: string): string[] => {
    const body = xml.match(/<w:pPr>(.*?)<\/w:pPr>/)?.[1] ?? ''
    const out: string[] = []
    const re = /<\/?w:(\w+)([^>]*)>/g
    let depth = 0
    let m: RegExpExecArray | null
    while ((m = re.exec(body))) {
      const closing = m[0].startsWith('</')
      const selfClosing = m[2].trimEnd().endsWith('/')
      if (closing) depth--
      else if (depth === 0) out.push(`w:${m[1]}`)
      if (!closing && !selfClosing) depth++
    }
    return out
  }

  it('leaves an original w:spacing after w:tabs and w:ind, not before them', () => {
    const existing =
      '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="–"/>' +
      '<w:pPr><w:spacing w:before="0"/><w:tabs><w:tab w:val="num" w:pos="720"/></w:tabs></w:pPr></w:lvl>'
    const order = pPrOrder(mergeLevelXml(existing, LEVEL, 0))
    // CT_PPrGeneral: tabs (10) < spacing (24) < ind (25)
    expect(order).toEqual(['w:tabs', 'w:spacing', 'w:ind'])
  })

  it('keeps the child it writes itself in the right place', () => {
    const existing =
      '<w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/><w:lvlText w:val="–"/>' +
      '<w:pPr><w:tabs><w:tab w:val="num" w:pos="720"/></w:tabs></w:pPr></w:lvl>'
    expect(pPrOrder(mergeLevelXml(existing, { ...LEVEL, tabStop: 720 }, 0))).toEqual([
      'w:tabs',
      'w:ind',
    ])
  })

  it('still emits ind when the original pPr had children it does not know', () => {
    const existing =
      '<w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/><w:lvlText w:val="–"/>' +
      '<w:pPr><w:spacing w:after="120"/></w:pPr></w:lvl>'
    const order = pPrOrder(mergeLevelXml(existing, LEVEL, 0))
    expect(order).toContain('w:ind')
    expect(order.indexOf('w:spacing')).toBeLessThan(order.indexOf('w:ind'))
  })

  it('sorts every schema child by its CT_PPrGeneral position, not around tabs/ind', () => {
    const existing =
      '<w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/><w:lvlText w:val="–"/>' +
      '<w:pPr><w:jc w:val="left"/><w:keepNext/><w:spacing w:after="120"/>' +
      '<w:tabs><w:tab w:val="num" w:pos="720"/></w:tabs></w:pPr></w:lvl>'
    // schema: keepNext (1) < tabs (10) < spacing (24) < ind (25) < jc (26)
    expect(pPrOrder(mergeLevelXml(existing, LEVEL, 0))).toEqual([
      'w:keepNext',
      'w:tabs',
      'w:spacing',
      'w:ind',
      'w:jc',
    ])
  })
})
