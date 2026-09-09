import { LexerTreeBuilder, sequenceNodeType } from '@univerjs/engine-formula'
import { describe, expect, it } from 'vitest'

import {
  fixSequenceNodes,
  guardLiteralNewlines,
  wrapLexer,
  type SequenceEntry,
} from '../src/renderer/formula-lexer-fix'

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

function stringTokens(nodes: SequenceEntry[] | undefined | null): string[] {
  return (nodes ?? []).flatMap((node) =>
    typeof node !== 'string' && node.nodeType === sequenceNodeType.STRING ? [node.token] : [],
  )
}

function treeLeaves(node: unknown): string[] {
  const tree = node as { getChildren(): unknown[] }
  return tree
    .getChildren()
    .flatMap((child) => (typeof child === 'string' ? [child] : treeLeaves(child)))
}

// Univer caches lexer output per formula text at module scope, so the stock
// and wrapped checks below must never share a formula string.
describe('newlines inside string literals', () => {
  const CASES = [
    ['LF', '="a\nb"&B1', '"a\nb"'],
    ['CRLF', '="a\r\nb"&B1', '"a\r\nb"'],
    ['LF after an escaped quote', '="x""\ny"&B1', '"x""\ny"'],
  ] as const

  it.each(CASES)('%s: the stock lexer flattens the literal to spaces', (_name, formula) => {
    const local = new LexerTreeBuilder()
    const stock = formula.replace('B1', 'Z1')
    expect(stringTokens(local.sequenceNodesBuilder(stock) as SequenceEntry[])[0]).toMatch(
      /^"[^\r\n]*"$/,
    )
    expect(treeLeaves(local.treeBuilder(stock)).join('')).not.toMatch(/[\r\n]/)
  })

  it.each(CASES)('%s: the wrapped lexer keeps the literal verbatim', (_name, formula, literal) => {
    const local = new LexerTreeBuilder()
    const wrap = wrapLexer(local)
    const nodes = local.sequenceNodesBuilder(formula) as SequenceEntry[]
    expect(stringTokens(nodes)).toEqual([literal])
    const body = formula.slice(1)
    for (const node of nodes) {
      if (typeof node === 'string') continue
      expect(body.slice(node.startIndex, node.endIndex + 1)).toBe(node.token)
    }
    expect(normalizeLikeCommit(formula)).toBe(formula)
    wrap.dispose()
  })

  it.each(CASES)('%s: the evaluation tree carries the break', (_name, formula, literal) => {
    const local = new LexerTreeBuilder()
    const wrap = wrapLexer(local)
    // The tree stores an escaped "" as a single quote.
    expect(treeLeaves(local.treeBuilder(formula))).toContain(literal.replace(/""/g, '"'))
    wrap.dispose()
  })

  it('autofill offset rewrites keep the break', () => {
    const local = new LexerTreeBuilder()
    expect(local.moveFormulaRefOffset('="a\nb"&Y1', 0, 1)).toBe('="a b"&Y2') // Univer today
    const wrap = wrapLexer(local)
    expect(local.moveFormulaRefOffset('="a\nb"&B1', 0, 1)).toBe('="a\nb"&B2')
    wrap.dispose()
    expect(local.moveFormulaRefOffset('="a\nb"&X1', 0, 1)).toBe('="a b"&X2')
  })

  it('still flattens a newline between tokens to whitespace', () => {
    const local = new LexerTreeBuilder()
    const wrap = wrapLexer(local)
    const formula = '=A1&\n"b"\r\n'
    const nodes = local.sequenceNodesBuilder(formula) as SequenceEntry[]
    expect(nodes.map((node) => (typeof node === 'string' ? node : node.token)).join('')).toBe(
      'A1& "b" ',
    )
    expect(stringTokens(nodes).map((token) => token.trim())).toEqual(['"b"'])
    expect(treeLeaves(local.treeBuilder(formula)).map((leaf) => leaf.trim())).toContain('"b"')
    expect(local.moveFormulaRefOffset('=A1&\n"b"', 0, 1)).toBe('=A2& "b"')
    wrap.dispose()
  })

  it('guards only the literal breaks and leaves quoted sheet names and table refs alone', () => {
    expect(guardLiteralNewlines('="a\nb"&\nB1')).toBe('="a\uFDD1b"&\nB1')
    expect(guardLiteralNewlines('=Table1[["\n"]]&"\r"')).toBe('=Table1[["\n"]]&"\uFDD0"')
    expect(guardLiteralNewlines("='it\n''s'!A1")).toBe("='it\n''s'!A1")
    expect(guardLiteralNewlines('="\uFDD1\n"')).toBe('="\uFDD1\n"')
  })
})
