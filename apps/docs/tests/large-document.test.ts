import { describe, expect, it } from 'vitest'
import type { Block } from '@genoffice/docx-engine'
import {
  CHARS_PER_BLOCK,
  HUGE_DOC_REFUSE_WEIGHT,
  LARGE_DOC_READ_ONLY_WEIGHT,
  LITE_DOC_WEIGHT,
  docTextLength,
  docWeight,
  openTierFor,
} from '../src/renderer/large-document'

describe('openTierFor', () => {
  it('tiers by weight with inclusive lower bounds', () => {
    expect(openTierFor(0)).toBe('normal')
    expect(openTierFor(LITE_DOC_WEIGHT)).toBe('normal')
    expect(openTierFor(LITE_DOC_WEIGHT + 1)).toBe('lite')
    expect(openTierFor(LARGE_DOC_READ_ONLY_WEIGHT)).toBe('lite')
    expect(openTierFor(LARGE_DOC_READ_ONLY_WEIGHT + 1)).toBe('readOnly')
    expect(openTierFor(HUGE_DOC_REFUSE_WEIGHT)).toBe('readOnly')
    expect(openTierFor(HUGE_DOC_REFUSE_WEIGHT + 1)).toBe('refuse')
  })
})

const para = (text: string): Block =>
  ({ id: 'p', type: 'paragraph', docxIndex: null, originalXml: null, runs: [{ text }] }) as Block

const table = (cells: string[][], nested?: Block): Block =>
  ({
    id: 't',
    type: 'table',
    docxIndex: null,
    originalXml: null,
    table: {
      rows: cells.map((row) =>
        row.map((text) => ({
          paras: [text],
          ...(nested ? { nestedTables: [nested.table] } : {}),
        })),
      ),
    },
  }) as unknown as Block

describe('docWeight', () => {
  it('counts blocks plus characters of paragraph runs and table cells, nested tables included', () => {
    const blocks = [para('a'.repeat(220)), table([['bb', 'cc']], table([['dddd']]))]
    expect(docTextLength(blocks)).toBe(220 + 4 + 4 + 4)
    expect(docWeight(blocks)).toBe(2 + Math.round(232 / CHARS_PER_BLOCK))
  })

  it('a long-paragraph document outweighs a short one with more blocks', () => {
    const short = Array.from({ length: 300 }, () => para('x'.repeat(100)))
    const long = Array.from({ length: 100 }, () => para('x'.repeat(700)))
    expect(docWeight(long)).toBeGreaterThan(docWeight(short))
  })
})
