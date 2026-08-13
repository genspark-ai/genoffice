import { describe, expect, it } from 'vitest'
import {
  BLANK_BULLET_NUM_ID,
  BLANK_ORDERED_NUM_ID,
  buildBlankDocx,
  parseDocx,
  saveDocx,
  type SaveBlock,
} from '../src/index'

describe('blank document template (new-document / AI generation base)', () => {
  it('parses to a single empty paragraph with standard heading styles available', async () => {
    const bytes = await buildBlankDocx()
    const doc = await parseDocx(bytes)
    const visible = doc.blocks.filter((b) => !b.hidden)
    expect(visible).toHaveLength(1)
    expect(visible[0].type).toBe('paragraph')
    for (const styleId of ['Heading1', 'Heading2', 'Heading3', 'Heading6']) {
      expect(doc.styles.has(styleId), styleId).toBe(true)
    }
  })

  it('AI-generated content (headings, paragraphs, lists) saves and reparses correctly', async () => {
    const bytes = await buildBlankDocx()
    const doc = await parseDocx(bytes)

    const generated: SaveBlock[] = [
      { kind: 'generated', block: { type: 'heading', level: 1, runs: [{ text: '生成的标题' }] } },
      {
        kind: 'generated',
        block: {
          type: 'paragraph',
          runs: [
            { text: '正文段落,', bold: false },
            { text: '重点', bold: true },
          ],
        },
      },
      {
        kind: 'generated',
        block: {
          type: 'listItem',
          list: { kind: 'bullet', numId: BLANK_BULLET_NUM_ID, ilvl: 0 },
          runs: [{ text: '无序项' }],
        },
      },
      {
        kind: 'generated',
        block: {
          type: 'listItem',
          list: { kind: 'ordered', numId: BLANK_ORDERED_NUM_ID, ilvl: 0 },
          runs: [{ text: '有序项' }],
        },
      },
    ]
    const saved = await saveDocx(doc, generated)
    const reparsed = await parseDocx(saved)
    const visible = reparsed.blocks.filter((b) => !b.hidden)

    expect(visible.map((b) => b.type)).toEqual(['heading', 'paragraph', 'listItem', 'listItem'])
    expect(visible[0].level).toBe(1)
    expect(visible[0].runs?.map((r) => r.text).join('')).toBe('生成的标题')
    expect(visible[1].runs?.some((r) => r.bold && r.text === '重点')).toBe(true)
    expect(visible[2].list?.kind).toBe('bullet')
    expect(visible[3].list?.kind).toBe('ordered')
  })
})
