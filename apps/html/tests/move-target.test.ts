import { describe, expect, it } from 'vitest'
import { buildParseMap } from '../src/renderer/document/parse-map'
import { moveTarget } from '../src/renderer/document/move-target'

const DOC = `<!doctype html>
<html><body>
<section id="a"><h1 id="t">T</h1><p id="p">P</p></section>
<section id="b"><p id="q">Q</p></section>
<ul id="list"><li id="l1">1</li><li id="l2">2</li></ul>
<table id="tbl"><tr id="row"><td id="c1">c</td></tr></table>
</body></html>`

const map = buildParseMap(DOC, 1, null)
const sid = (id: string) =>
  map.elements.find((e) => DOC.slice(...e.startTag).includes(`id="${id}"`))!.sid

describe('moveTarget', () => {
  it('swaps with the neighbouring sibling when there is one', () => {
    expect(moveTarget(map, sid('p'), -1)).toEqual({ position: 'before', ref_sid: sid('t') })
    expect(moveTarget(map, sid('t'), 1)).toEqual({ position: 'after', ref_sid: sid('p') })
  })

  it('leaves the parent at the first / last position', () => {
    expect(moveTarget(map, sid('t'), -1)).toEqual({ position: 'before', ref_sid: sid('a') })
    expect(moveTarget(map, sid('p'), 1)).toEqual({ position: 'after', ref_sid: sid('a') })
  })

  it('never moves out of body', () => {
    expect(moveTarget(map, sid('a'), -1)).toBeNull()
    expect(moveTarget(map, sid('tbl'), 1)).toBeNull()
  })

  it('list and table parts swap with siblings but never leave their parent', () => {
    expect(moveTarget(map, sid('l2'), -1)).toEqual({ position: 'before', ref_sid: sid('l1') })
    expect(moveTarget(map, sid('l1'), -1)).toBeNull()
    expect(moveTarget(map, sid('l2'), 1)).toBeNull()
    expect(moveTarget(map, sid('c1'), -1)).toBeNull()
    expect(moveTarget(map, sid('row'), 1)).toBeNull()
  })
})
