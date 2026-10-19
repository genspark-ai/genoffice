/**
 * Rewriting an existing text box must not silently change its font size: plain
 * runs inherit the original size, and a request that carries an explicit size
 * which differs from the box's returns a note so the AI can undo the drift.
 */
import { describe, it, expect, beforeEach } from 'vitest'
import { addElement, createBlankPptx, openPptx, type OpenedPptx } from '@genoffice/pptx-engine'
import { runTxn } from '../src/ops/executor'
import '../src/ops/index'

let opened: OpenedPptx
let textId: string

beforeEach(async () => {
  opened = await openPptx(await createBlankPptx())
  const el = addElement(opened.deck.slides[0]!, {
    kind: 'textbox',
    offset: { x: 0, y: 0, cx: 1828800, cy: 914400 },
    paragraphs: [{ runs: [{ text: 'XXX,', fontSize: 15 }] }],
  })
  textId = el.id
})

const setText = (runs: Array<{ text: string; fontSize?: number }>) =>
  runTxn(opened, {
    ops: [{ op: 'setText', target: { slide: 0, el: textId }, paragraphs: [{ runs }] }],
  })

const note = (r: ReturnType<typeof runTxn>) => (r.records![0]!.after as { note?: string }).note

describe('setText font size on an existing box', () => {
  it('plain runs inherit the original size and produce no note', () => {
    const r = setText([{ text: 'A summary sentence' }])
    expect(r.applied).toBe(true)
    expect(note(r)).toBeUndefined()
    const runs = (opened.deck.slides[0]!.elements.find((e) => e.id === textId) as any).text
      .paragraphs[0].runs
    expect(runs[0].fontSize).toBe(15)
  })

  it('an explicit different size is applied but flagged', () => {
    const r = setText([{ text: 'A summary sentence', fontSize: 14 }])
    expect(r.applied).toBe(true)
    expect(note(r)).toContain('15pt → 14pt')
  })

  it('an explicit size equal to the original is silent', () => {
    const r = setText([{ text: 'A summary sentence', fontSize: 15 }])
    expect(note(r)).toBeUndefined()
  })
})
