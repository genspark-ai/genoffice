/**
 * Escaped-quote lexer fix. Univer's LexerTreeBuilder._nodeMaker emits
 * one sequence slot for the two characters of an escaped quote ("") inside a
 * string literal, and getSequenceNode treats slot index as character
 * position, so every node after an escape pair reports startIndex/endIndex
 * shifted left by the number of pairs before it (and STRING tokens drop a
 * quote). sheets-ui's normalizeFormulaString then splices the formula text
 * by those indices on commit, silently rewriting `="a"""&B1` into
 * `="a"""B11`; autofill's moveFormulaRefOffset and F4's
 * convertRefersToAbsolute misassemble the same way.
 *
 * The wrap re-scans the formula with the same quote/bracket state machine,
 * builds the slot→character map the lexer should have produced, and returns
 * cloned nodes with corrected indices and tokens. Formulas whose nodes
 * already reassemble to the source text pass through untouched.
 *
 * Newline fix. _nodeMaker flattens every CR/LF in the formula to a space so
 * line breaks between tokens act as whitespace — but that also rewrites
 * `="a"&CHAR(10)`-style literals typed with a real line break (Excel keeps
 * them; JP form templates build whole paragraphs this way), so the engine
 * evaluates and autofill/F4 rewrite the text with the break gone. The wrap
 * swaps CR/LF inside "..." for private guard characters before the lexer
 * sees the text and puts them back in every token it produced.
 */
import { LexerTreeBuilder, type sequenceNodeType } from '@univerjs/engine-formula'

import type { UniverRuntime } from './univer-state'

export interface SequenceNodeLike {
  nodeType: sequenceNodeType
  token: string
  startIndex: number
  endIndex: number
}

export type SequenceEntry = SequenceNodeLike | string

/// Unicode noncharacters: never legitimate formula text, so a formula that
/// already contains one is left to the stock lexer rather than corrupted.
const CR_GUARD = '\uFDD0'
const LF_GUARD = '\uFDD1'
const GUARD_PATTERN = /[\uFDD0\uFDD1]/g

/// CR/LF inside "..." literals become guard characters; everything else is
/// untouched so _nodeMaker still flattens the between-token breaks.
export function guardLiteralNewlines(formula: string): string {
  if (!/[\r\n]/.test(formula) || GUARD_PATTERN.test(formula)) return formula
  let out = ''
  let inDouble = false
  let inSingle = false
  let bracketDepth = 0
  for (let at = 0; at < formula.length; at += 1) {
    const char = formula[at]
    if (inDouble) {
      if (char === '\r') out += CR_GUARD
      else if (char === '\n') out += LF_GUARD
      else {
        out += char
        if (char === '"') {
          if (formula[at + 1] === '"') {
            out += '"'
            at += 1
          } else inDouble = false
        }
      }
      continue
    }
    out += char
    if (inSingle) {
      if (char === "'") {
        if (formula[at + 1] === "'") {
          out += "'"
          at += 1
        } else inSingle = false
      }
    } else if (char === '"' && bracketDepth === 0) {
      inDouble = true
    } else if (char === "'") {
      inSingle = true
    } else if (char === '[') {
      bracketDepth += 1
    } else if (char === ']' && bracketDepth > 0) {
      bracketDepth -= 1
    }
  }
  return out
}

function restoreGuards(text: string): string {
  return text.replace(GUARD_PATTERN, (guard) => (guard === CR_GUARD ? '\r' : '\n'))
}

/// Mirror of the guarded _nodeMaker's input preprocessing.
function formulaBody(formula: string): string {
  const flattened = restoreGuards(guardLiteralNewlines(formula).replace(/\r\n$|\r|\n/g, ' '))
  return flattened.startsWith('=') ? flattened.slice(1) : flattened
}

/// Slot index → character index in `body`, replicating the lexer's states:
/// an escaped "" inside a string yields one slot; '' inside a quoted sheet
/// name yields two; quotes are literal inside [table refs] and vice versa.
function buildSlotMap(body: string): number[] {
  const map: number[] = []
  let inDouble = false
  let inSingle = false
  let bracketDepth = 0
  for (let at = 0; at < body.length; at += 1) {
    const char = body[at]
    map.push(at)
    if (inDouble) {
      if (char === '"') {
        if (body[at + 1] === '"') at += 1
        else inDouble = false
      }
    } else if (inSingle) {
      if (char === "'") {
        if (body[at + 1] === "'") {
          map.push(at + 1)
          at += 1
        } else inSingle = false
      }
    } else if (char === '"' && bracketDepth === 0) {
      inDouble = true
    } else if (char === "'") {
      inSingle = true
    } else if (char === '[') {
      bracketDepth += 1
    } else if (char === ']' && bracketDepth > 0) {
      bracketDepth -= 1
    }
  }
  return map
}

export function fixSequenceNodes(
  formula: string,
  nodes: SequenceEntry[] | undefined | null,
): SequenceEntry[] | undefined | null {
  if (!nodes) return nodes
  const body = formulaBody(formula)
  const reassembled = nodes.map((node) => (typeof node === 'string' ? node : node.token)).join('')
  if (reassembled === body) return nodes
  const map = buildSlotMap(body)
  let changed = false
  const fixed = nodes.map((node) => {
    if (typeof node === 'string') return node
    const startIndex = map[node.startIndex]
    const boundary = map[node.endIndex + 1] ?? body.length
    if (startIndex === undefined) return node
    const endIndex = boundary - 1
    if (startIndex === node.startIndex && endIndex === node.endIndex) return node
    changed = true
    return { ...node, startIndex, endIndex, token: body.slice(startIndex, endIndex + 1) }
  })
  return changed ? fixed : nodes
}

interface LexerNodeLike {
  getToken(): string
  setToken(token: string): void
  getChildren(): (LexerNodeLike | string)[]
  setChildren(children: (LexerNodeLike | string)[]): void
  getParent(): LexerNodeLike | null | undefined
}

interface SequenceArrayEntry {
  segment: string
  currentString: string
}

type NodeMaker = (
  formula: string,
  sequenceArray?: SequenceArrayEntry[],
  matchCurrentNodeIndex?: number,
) => unknown

interface LexerInternals {
  _nodeMaker: NodeMaker
  _currentLexerNode: LexerNodeLike
}

function restoreTree(node: LexerNodeLike): void {
  node.setToken(restoreGuards(node.getToken()))
  node.setChildren(
    node.getChildren().map((child) => {
      if (typeof child === 'string') return restoreGuards(child)
      restoreTree(child)
      return child
    }),
  )
}

function wrapNodeMaker(lexer: LexerInternals): () => void {
  const original = lexer._nodeMaker
  lexer._nodeMaker = function (this: LexerInternals, formula, sequenceArray, matchIndex) {
    const guarded = guardLiteralNewlines(formula)
    if (guarded === formula) return original.call(this, formula, sequenceArray, matchIndex)
    const firstNew = sequenceArray?.length ?? 0
    const result = original.call(this, guarded, sequenceArray, matchIndex)
    let top = this._currentLexerNode
    while (top?.getParent()) top = top.getParent() as LexerNodeLike
    if (top) restoreTree(top)
    for (const entry of sequenceArray?.slice(firstNew) ?? []) {
      entry.segment = restoreGuards(entry.segment)
      entry.currentString = restoreGuards(entry.currentString)
    }
    // matchCurrentNodeIndex mode returns [node, char] for the cursor position.
    if (Array.isArray(result) && typeof result[1] === 'string') {
      return [result[0], restoreGuards(result[1])]
    }
    return result
  }
  return () => {
    lexer._nodeMaker = original
  }
}

export function wrapLexer(lexer: LexerTreeBuilder): { dispose(): void } {
  const original = lexer.sequenceNodesBuilder.bind(lexer) as (
    formula: string,
  ) => SequenceEntry[] | undefined | null
  const wrapped = (formula: string) => fixSequenceNodes(formula, original(formula))
  ;(lexer as { sequenceNodesBuilder: typeof wrapped }).sequenceNodesBuilder = wrapped
  const unwrapNodeMaker = wrapNodeMaker(lexer as unknown as LexerInternals)
  return {
    dispose() {
      delete (lexer as { sequenceNodesBuilder?: typeof wrapped }).sequenceNodesBuilder
      unwrapNodeMaker()
    },
  }
}

/// Wrap the injector's LexerTreeBuilder singleton — sheets-ui (commit
/// normalization), sheets-formula-ui (highlight), and engine-formula
/// (autofill/F4 offset rewrites) all resolve the same instance.
export function installFormulaLexerFix(runtime: UniverRuntime): { dispose(): void } {
  return wrapLexer(runtime.univer.__getInjector().get(LexerTreeBuilder))
}
