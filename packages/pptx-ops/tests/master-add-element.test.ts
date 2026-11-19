/**
 * Part-addressed addElement: the master edit view (and CLI ops) insert text
 * boxes and shapes onto master/layout parts through the same op the slide
 * path runs. The part model grows, the created id lands in the journal, and
 * the archive entry is re-serialized with the new shape at transaction end.
 */
import { describe, it, expect } from 'vitest'
import { createBlankPptx, openPptx, parseMasterPart } from '@genoffice/pptx-engine'
import { runTxn } from '../src/ops/executor'
import '../src/ops/index'

const MASTER = 'ppt/slideMasters/slideMaster1.xml'
const LAYOUT = 'ppt/slideLayouts/slideLayout1.xml'

describe('addElement on master/layout parts', () => {
  it('inserts a text box onto the seeded master part and writes the entry back', async () => {
    const opened = await openPptx(await createBlankPptx())
    const master = parseMasterPart(opened.archive, MASTER)
    expect(master).not.toBeNull()
    const before = master!.elements.length
    const r = runTxn(opened, {
      ops: [
        {
          op: 'addElement',
          target: { part: MASTER },
          kind: 'textbox',
          offset: { x: 914400, y: 914400, cx: 1828800, cy: 914400 },
          paragraphs: [{ runs: [{ text: 'Master box' }] }],
        },
      ],
      parts: new Map([[MASTER, master!]]),
    })
    expect(r.applied).toBe(true)
    expect(r.records?.[0]?.created?.[0]).toBeTruthy()
    expect(master!.elements.length).toBe(before + 1)
    expect(opened.archive.readText(MASTER)).toContain('Master box')
  })

  it('parses an unseeded layout part on demand and inserts a preset shape', async () => {
    const opened = await openPptx(await createBlankPptx())
    const r = runTxn(opened, {
      ops: [
        {
          op: 'addElement',
          target: { part: LAYOUT },
          kind: 'roundRect',
          offset: { x: 914400, y: 914400, cx: 1828800, cy: 914400 },
        },
      ],
    })
    expect(r.applied).toBe(true)
    expect(r.records?.[0]?.created?.[0]).toBeTruthy()
    expect(opened.archive.readText(LAYOUT)).toContain('prst="roundRect"')
  })
})
