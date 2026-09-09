import { describe, expect, it } from 'vitest'
import { buildParseMap } from '../src/renderer/document/parse-map'
import {
  buildQueueInstruction,
  buildQueueSummary,
  buildSelectionInstruction,
  liveItems,
  resolveQueue,
  resolveQueueItem,
  type EditQueueItem,
} from '../src/renderer/ai/edit-queue'

const html = [
  '<!doctype html>',
  '<html><body>',
  '<h1 id="t">Hello &amp; welcome</h1>',
  '<p class="lead">First paragraph</p>',
  '<img src="a.png" alt="">',
  '</body></html>',
].join('\n')

function sidOf(tag: string) {
  const map = buildParseMap(html, 1)
  return map.elements.find((e) => e.tag === tag)!.sid
}

const item = (qid: string, tag: string, instruction: string): EditQueueItem => ({
  qid,
  sid: sidOf(tag),
  tag,
  capturedText: 'captured',
  instruction,
})

describe('resolveQueueItem', () => {
  it('reads the live excerpt from the source and labels textless elements by tag', () => {
    const map = buildParseMap(html, 1)
    expect(resolveQueueItem(html, map, item('a', 'h1', 'x')).target?.excerpt).toBe(
      'Hello & welcome',
    )
    expect(resolveQueueItem(html, map, item('b', 'img', 'x')).target?.excerpt).toBe('<img>')
  })

  it('marks an item stale when its sid is gone or now names another tag', () => {
    const map = buildParseMap(html, 1)
    const gone = { ...item('a', 'p', 'x'), sid: 999 }
    expect(resolveQueueItem(html, map, gone).target).toBeNull()
    const retagged = { ...item('b', 'p', 'x'), tag: 'div' }
    expect(resolveQueueItem(html, map, retagged).target).toBeNull()
  })
})

describe('buildQueueInstruction', () => {
  it('lists edits bottom-up by source position with sid, tag and excerpt', () => {
    const map = buildParseMap(html, 1)
    const entries = liveItems(
      resolveQueue(html, map, [item('a', 'h1', 'shorten'), item('b', 'p', 'expand')]),
    )
    const text = buildQueueInstruction(entries)
    const pIdx = text.indexOf(`1. sid=${sidOf('p')} <p> "First paragraph"`)
    const hIdx = text.indexOf(`2. sid=${sidOf('h1')} <h1> "Hello & welcome"`)
    expect(pIdx).toBeGreaterThan(-1)
    expect(hIdx).toBeGreaterThan(pIdx)
    expect(text).toContain('Requested change: expand')
    expect(text).toContain('Requested change: shorten')
  })

  it('summarizes in queue order for the chat bubble', () => {
    const map = buildParseMap(html, 1)
    const entries = liveItems(
      resolveQueue(html, map, [item('a', 'h1', 'shorten'), item('b', 'p', 'expand')]),
    )
    expect(buildQueueSummary('Batch:', entries)).toBe(
      'Batch:\n1. Hello & welcome — shorten\n2. First paragraph — expand',
    )
  })

  it('pins a send-now instruction to its element', () => {
    const map = buildParseMap(html, 1)
    const [entry] = liveItems(resolveQueue(html, map, [item('a', 'p', 'make it bold')]))
    expect(buildSelectionInstruction(entry!.target, 'make it bold')).toBe(
      `Apply this to the element sid=${sidOf('p')} <p> ("First paragraph"), addressing it by sid in apply_ops:\nmake it bold`,
    )
  })
})
