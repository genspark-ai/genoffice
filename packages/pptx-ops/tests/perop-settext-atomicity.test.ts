/**
 * A rejected op contributes no model or package changes, while preceding and
 * following successful operations stay (genoffice#2821): setText rewrote the runs and
 * their link rels before the picture-bullet landing failed, so a per_op batch
 * kept the rejected op's text and hyperlink. The bullet format is now refused
 * in validate — the plan phase — where a per_op batch skips the op whole.
 */
import { describe, it, expect, beforeAll } from 'vitest'
import {
  addElement,
  createBlankPptx,
  openPptx,
  patchSlideXml,
  type OpenedPptx,
} from '@genoffice/pptx-engine'
import { runTxn } from '../src/ops/executor'
import '../src/ops/index'

let opened: OpenedPptx
let textId: string

beforeAll(async () => {
  opened = await openPptx(await createBlankPptx())
  const slide = opened.deck.slides[0]!
  const el = addElement(slide, {
    kind: 'textbox',
    offset: { x: 0, y: 0, cx: 1828800, cy: 914400 },
    paragraphs: [{ runs: [{ text: 'Original text' }] }],
  })
  textId = el.id
})

describe('per_op keeps a rejected setText whole', () => {
  it('a failing setText leaves text and rels untouched; later ops still land', () => {
    const r = runTxn(opened, {
      isolation: 'per_op',
      ops: [
        {
          op: 'setText',
          target: { slide: 0, el: textId },
          paragraphs: [
            {
              runs: [
                {
                  text: 'Rejected replacement',
                  link: { kind: 'url', url: 'https://example.com/rejected' },
                },
              ],
              bullet: 'blip',
              bulletImage: { base64: 'AA==', ext: 'not-real' },
            },
          ],
        },
        {
          op: 'addElement',
          target: { slide: 0 },
          kind: 'textbox',
          offset: { x: 0, y: 1828800, cx: 1828800, cy: 914400 },
          paragraphs: [{ runs: [{ text: 'Successful marker' }] }],
        },
      ],
    })
    expect(r.applied).toBe(true)
    expect(r.records).toHaveLength(1)
    expect(r.failures).toHaveLength(1)
    expect(r.failures![0]!.index).toBe(0)
    expect(r.failures![0]!.error).toContain('unsupported picture bullet format ".not-real"')

    const xml = patchSlideXml(opened.deck.slides[0]!)
    expect(xml).toContain('Original text')
    expect(xml).not.toContain('Rejected replacement')
    expect(xml).not.toContain('example.com/rejected')
    expect(xml).toContain('Successful marker')
    // the rejected bullet's media part never lands either
    for (const path of opened.archive.entries.keys()) {
      expect(path).not.toMatch(/ppt\/media\/image\d+\.not-real/)
    }
  })
})
