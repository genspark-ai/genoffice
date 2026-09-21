import { describe, expect, it } from 'vitest'
import { ancestorsOf, buildParseMap, type ParseMap } from '../src/renderer/document/parse-map'
import { label, MAX_CRUMB_ID_CHARS } from '../src/renderer/components/Breadcrumb'

describe('ancestorsOf', () => {
  it('walks the real chain root-first', () => {
    const map = buildParseMap('<div><main><p>hi</p></main></div>', 1)
    const p = map.elements.find((e) => e.tag === 'p')!
    const chain = ancestorsOf(map, p.sid).map((e) => e.tag)
    expect(chain[chain.length - 1]).toBe('main')
    expect(chain).toContain('div')
  })

  it('terminates on a corrupt parentSid cycle instead of hanging', () => {
    const map = buildParseMap('<div><p>hi</p></div>', 1)
    const div = map.elements.find((e) => e.tag === 'div')!
    const p = map.elements.find((e) => e.tag === 'p')!
    const cyclic: ParseMap = {
      ...map,
      bySid: new Map(map.bySid),
    }
    cyclic.bySid.set(div.sid, { ...div, parentSid: p.sid })
    cyclic.bySid.set(p.sid, { ...p, parentSid: div.sid })
    const chain = ancestorsOf(cyclic, p.sid)
    expect(chain.length).toBeLessThan(10)
  })
})

describe('breadcrumb label', () => {
  it('truncates KB-long ids', () => {
    const long = label(`<div id="${'x'.repeat(500)}" class="a b c">`, 'div')
    expect(long.length).toBeLessThan(100)
    expect(long).toContain('…')
    expect(label('<p>hi</p>', 'p')).toBe('p')
    expect(MAX_CRUMB_ID_CHARS).toBe(48)
  })
})
