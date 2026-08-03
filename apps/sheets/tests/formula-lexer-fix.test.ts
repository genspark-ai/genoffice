import { LexerTreeBuilder, sequenceNodeType } from '@univerjs/engine-formula'
import { describe, expect, it } from 'vitest'

import { fixSequenceNodes, wrapLexer, type SequenceEntry } from '../src/renderer/formula-lexer-fix'

const lexer = new LexerTreeBuilder()

function fixedNodes(formula: string): SequenceEntry[] {
  return fixSequenceNodes(
    formula,
    lexer.sequenceNodesBuilder(formula) as SequenceEntry[] | undefined,
  ) as SequenceEntry[]
}

/**
 * Replica of sheets-ui's normalizeFormulaString splice: FUNCTION and
 * REFERENCE tokens are upper-cased and written back by node index. With the
 * broken indices this is what rewrote `="a"""&B1` into `="a"""B11`.
 */
function normalizeLikeCommit(formula: string): string {
  const nodes = fixedNodes(formula)
  let normalized = formula
  const totalOffset = 0
  for (const node of nodes) {
    if (typeof node === 'string') continue
    if (
      node.nodeType === sequenceNodeType.FUNCTION ||
      node.nodeType === sequenceNodeType.REFERENCE
    ) {
      const start = node.startIndex + totalOffset + 1
      const end = node.endIndex + totalOffset + 2
      normalized =
        normalized.substring(0, start) + node.token.toUpperCase() + normalized.substring(end)
    }
  }
  return normalized
}

// The 7 minimal cases from the original bug report — typed text must round-trip identically.
const CASES = [
  '="a"""&B1',
  '="a""b"&UPPER(B1)',
  '="a"&UPPER(B1)',
  '=UPPER(B1)&"a"""',
  '="a"""&"b"',
  '=LEFT(B1,FIND(",",B1)-1)',
  '="say ""hi"""',
]

describe('fixSequenceNodes', () => {
  it.each(CASES)('round-trips %s through commit normalization', (formula) => {
    expect(normalizeLikeCommit(formula)).toBe(formula)
  })

  it('re-derives exact node indices after escaped quotes', () => {
    const nodes = fixedNodes('="a"""&B1')
    const body = '"a"""&B1'
    for (const node of nodes) {
      if (typeof node === 'string') continue
      expect(body.slice(node.startIndex, node.endIndex + 1)).toBe(node.token)
    }
    const reference = nodes.find(
      (node) => typeof node !== 'string' && node.nodeType === sequenceNodeType.REFERENCE,
    ) as { startIndex: number; endIndex: number; token: string }
    expect(reference.token).toBe('B1')
    expect(reference.startIndex).toBe(6)
    expect(reference.endIndex).toBe(7)
  })

  it('restores the swallowed quote in STRING tokens', () => {
    const nodes = fixedNodes('="a"""&B1')
    const literal = nodes.find(
      (node) => typeof node !== 'string' && node.nodeType === sequenceNodeType.STRING,
    ) as { token: string }
    expect(literal.token).toBe('"a"""')
  })

  it('handles multiple escape pairs (drift grows per pair)', () => {
    const formula = '="x""y""z"&A1+B2'
    expect(normalizeLikeCommit(formula)).toBe(formula)
  })

  it('leaves clean formulas untouched (same array identity)', () => {
    const raw = lexer.sequenceNodesBuilder('="a"&UPPER(B1)') as SequenceEntry[]
    expect(fixSequenceNodes('="a"&UPPER(B1)', raw)).toBe(raw)
  })

  it('keeps quoted sheet names with escaped single quotes intact', () => {
    // The replica upper-cases the whole reference token (the real
    // normalizer preserves sheet-name case); assert index integrity only.
    const formula = '=\'it\'\'s\'!A1&"x""y"'
    expect(normalizeLikeCommit(formula).toLowerCase()).toBe(formula.toLowerCase())
  })

  it('repairs autofill offset rewrites through the instance wrap', () => {
    const local = new LexerTreeBuilder()
    expect(local.moveFormulaRefOffset('="a"""&B1', 0, 1)).toBe('="a""&B2') // broken today
    const wrap = wrapLexer(local)
    expect(local.moveFormulaRefOffset('="a"""&B1', 0, 1)).toBe('="a"""&B2')
    expect(local.moveFormulaRefOffset('="a"&B1', 0, 1)).toBe('="a"&B2')
    wrap.dispose()
    expect(local.moveFormulaRefOffset('="a"""&B1', 0, 1)).toBe('="a""&B2')
  })

  it('never mutates the nodes Univer cached', () => {
    const raw = lexer.sequenceNodesBuilder('="a"""&B1') as SequenceEntry[]
    const before = JSON.stringify(raw)
    fixSequenceNodes('="a"""&B1', raw)
    expect(JSON.stringify(raw)).toBe(before)
  })
})
