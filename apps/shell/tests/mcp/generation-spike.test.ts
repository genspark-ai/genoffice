import { describe, expect, it } from 'vitest'
import {
  BLANK_BULLET_NUM_ID,
  BLANK_ORDERED_NUM_ID,
  buildBlankDocx,
  parseDocx,
  saveDocx,
  type SaveBlock,
} from '@genoffice/docx-engine'

/**
 * M0 de-risking spike: the whole docx generation recipe
 * (blank template -> parse -> saveDocx) must run in a plain Node environment,
 * with no DOM available. If this fails, the MCP server cannot generate
 * documents from the main process and the plan changes.
 *
 * Mirrors the engine's own blank-template test, but asserts the headless
 * precondition explicitly.
 */
describe('mcp docx generation spike (headless)', () => {
  it('runs without any DOM globals', () => {
    expect(typeof document).toBe('undefined')
    expect(typeof window).toBe('undefined')
  })

  it('builds a docx from the blank template and reparses with the expected structure', async () => {
    const parsed = await parseDocx(await buildBlankDocx())

    const blocks: SaveBlock[] = [
      {
        kind: 'generated',
        block: { type: 'heading', level: 1, runs: [{ text: 'Quarterly Report' }] },
      },
      {
        kind: 'generated',
        block: {
          type: 'paragraph',
          runs: [
            { text: 'Plain ' },
            { text: 'bold', bold: true },
            { text: ' and ' },
            { text: 'italic', italic: true },
            { text: '.' },
          ],
        },
      },
      {
        kind: 'generated',
        block: {
          type: 'listItem',
          list: { kind: 'bullet', numId: BLANK_BULLET_NUM_ID, ilvl: 0 },
          runs: [{ text: 'First point' }],
        },
      },
      {
        kind: 'generated',
        block: {
          type: 'listItem',
          list: { kind: 'ordered', numId: BLANK_ORDERED_NUM_ID, ilvl: 0 },
          runs: [{ text: 'Step one' }],
        },
      },
    ]

    const saved = await saveDocx(parsed, blocks)
    expect(saved).toBeInstanceOf(Uint8Array)
    expect(saved.byteLength).toBeGreaterThan(0)

    const reparsed = await parseDocx(saved)
    const visible = reparsed.blocks.filter((b) => !b.hidden)
    expect(visible.map((b) => b.type)).toEqual(['heading', 'paragraph', 'listItem', 'listItem'])
    expect(visible[0].level).toBe(1)
    expect(visible[0].runs?.map((r) => r.text).join('')).toBe('Quarterly Report')
    expect(visible[1].runs?.some((r) => r.bold && r.text === 'bold')).toBe(true)
    expect(visible[1].runs?.some((r) => r.italic && r.text === 'italic')).toBe(true)
    expect(visible[2].list?.kind).toBe('bullet')
    expect(visible[3].list?.kind).toBe('ordered')
  })
})
