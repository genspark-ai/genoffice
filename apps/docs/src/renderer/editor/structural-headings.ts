/**
 * Structural headings for documents that carry no heading styles.
 *
 * Legal letters often mark a section with a short bold or all-caps line, or
 * with a section number ("I. BACKGROUND", "II. RENEWAL TERMS"), on an ordinary
 * paragraph. The outline and story views read that structure here without
 * touching the document: a ref is a view-time guess, never a style change.
 *
 * The rules are deliberately narrow. A candidate must be a top-level text
 * paragraph (list items and table cells are left alone), fit on one line, end
 * without sentence punctuation, and be followed by an ordinary paragraph.
 * Numbering adds a level: Roman numerals sit above Arabic numbers and letters.
 */
import type { Node as PmNode } from '@tiptap/pm/model'
import { collectHeadings, type HeadingRef, type HeadingStyles } from './headings'

export interface StructuralHeadingRef extends HeadingRef {
  inferred: true
  /** what marked the line as a heading */
  basis: 'bold' | 'caps' | 'number'
}

type NumberKind = 'roman' | 'arabic' | 'letter'

const MAX_HEADING_CHARS = 80
const MAX_HEADING_WORDS = 12

const ROMAN = /^M{0,3}(CM|CD|D?C{0,3})(XC|XL|L?X{0,3})(IX|IV|V?I{0,3})$/
const MINOR_WORDS = new Set([
  'a',
  'an',
  'and',
  'as',
  'at',
  'but',
  'by',
  'for',
  'from',
  'in',
  'into',
  'nor',
  'of',
  'on',
  'or',
  'the',
  'to',
  'with',
  'vs',
  'v',
])

function isRoman(token: string): boolean {
  return token.length > 0 && ROMAN.test(token)
}

function numberPrefix(text: string): { kind: NumberKind; rest: string } | null {
  const match = /^([A-Z0-9]{1,7})\.(?=\s|$)/.exec(text)
  if (!match) return null
  const token = match[1]
  const rest = text.slice(match[0].length)
  if (/^\d{1,3}$/.test(token)) return { kind: 'arabic', rest }
  if (isRoman(token)) return { kind: 'roman', rest }
  if (/^[A-Z]$/.test(token)) return { kind: 'letter', rest }
  return null
}

/** one line of plain text: inline pictures, equations and breaks disqualify the block */
function plainText(node: PmNode): string | null {
  let text = ''
  let plain = true
  node.forEach((child) => {
    if (!child.isText) {
      plain = false
      return
    }
    text += child.text ?? ''
  })
  return plain ? text : null
}

function textOf(node: PmNode): string | null {
  const text = plainText(node)
  return text == null ? null : text.trim()
}

function isShort(text: string): boolean {
  if (text.length > MAX_HEADING_CHARS) return false
  return text.split(/\s+/).filter(Boolean).length <= MAX_HEADING_WORDS
}

function sentenceEnd(text: string): boolean {
  return /[.!?;,]$/.test(text)
}

function allBold(node: PmNode): boolean {
  let sawText = false
  let bold = true
  node.forEach((child) => {
    if (!child.isText) return
    const text = child.text ?? ''
    if (!text.trim()) return
    sawText = true
    const marked = child.marks.some((mark) => mark.type.name === 'bold')
    const off = child.marks.some(
      (mark) => mark.type.name === 'docTextStyle' && mark.attrs.boldOff === true,
    )
    if (!marked || off) bold = false
  })
  return sawText && bold
}

function capsMarked(child: PmNode): boolean {
  return child.marks.some((mark) => mark.type.name === 'docTextStyle' && mark.attrs.caps === 'all')
}

function allCaps(node: PmNode): boolean {
  let sawCased = false
  let caps = true
  node.forEach((child) => {
    if (!child.isText) return
    const text = child.text ?? ''
    const marked = capsMarked(child)
    for (const char of text) {
      // only letters that have case count as evidence; CJK text never qualifies
      if (char.toLowerCase() === char.toUpperCase()) continue
      sawCased = true
      if (!marked && char !== char.toUpperCase()) caps = false
    }
  })
  return sawCased && caps
}

function titleCase(text: string): boolean {
  const words = text.split(/\s+/).filter(Boolean)
  let sawWord = false
  for (const [index, word] of words.entries()) {
    const core = word.replace(/^[^\p{L}\p{N}]+|[^\p{L}\p{N}]+$/gu, '')
    const first = [...core].find((char) => /\p{L}/u.test(char))
    if (!first) continue
    sawWord = true
    if (index > 0 && MINOR_WORDS.has(core.toLowerCase())) continue
    if (first !== first.toUpperCase()) return false
  }
  return sawWord
}

function numberedHeading(text: string, node: PmNode): NumberKind | null {
  const prefix = numberPrefix(text)
  if (!prefix) return null
  if (!prefix.rest.trim()) return null
  if (allBold(node) || allCaps(node)) return prefix.kind
  // a numbered sentence ("1. The tenant shall pay rent") is not a heading
  return titleCase(prefix.rest) ? prefix.kind : null
}

/** whether the paragraph's own shape could be a heading, before the follow-up check */
function headingShape(
  node: PmNode,
): { basis: StructuralHeadingRef['basis']; number?: NumberKind } | null {
  if (node.type.name !== 'docParagraph') return null
  const text = textOf(node)
  if (!text || !isShort(text) || sentenceEnd(text)) return null
  const bold = allBold(node)
  const caps = allCaps(node)
  const number = numberedHeading(text, node)
  if (number) return { basis: 'number', number }
  if (bold) return { basis: 'bold' }
  if (caps) return { basis: 'caps' }
  return null
}

function isOrdinaryParagraph(node: PmNode | undefined): boolean {
  if (node?.type.name !== 'docParagraph') return false
  const text = textOf(node)
  return text != null && text.length > 0 && headingShape(node) == null
}

/** the first block after the candidate that is not an empty paragraph */
function nextContentBlock(
  blocks: ReadonlyArray<{ node: PmNode; pos: number; shape: ReturnType<typeof headingShape> }>,
  index: number,
): PmNode | undefined {
  for (let at = index + 1; at < blocks.length; at++) {
    const node = blocks[at].node
    if (node.type.name === 'docParagraph' && !textOf(node)) continue
    return node
  }
  return undefined
}

function numberedLevels(
  candidates: ReadonlyArray<{ number?: NumberKind }>,
): Map<NumberKind, number> {
  const levels = new Map<NumberKind, number>()
  let next = 1
  for (const kind of ['roman', 'arabic', 'letter'] as const) {
    if (candidates.some((candidate) => candidate.number === kind)) levels.set(kind, next++)
  }
  return levels
}

/**
 * Headings read from the document's shape, or an empty list when the document
 * already has heading styles (a real heading is always preferred over a guess).
 */
export function collectStructuralHeadings(
  doc: PmNode,
  styles?: HeadingStyles,
): StructuralHeadingRef[] {
  if (collectHeadings(doc, styles).length > 0) return []
  const blocks: Array<{ node: PmNode; pos: number; shape: ReturnType<typeof headingShape> }> = []
  doc.forEach((node, pos) => blocks.push({ node, pos, shape: headingShape(node) }))
  const accepted = blocks.filter(
    (block, index) => block.shape != null && isOrdinaryParagraph(nextContentBlock(blocks, index)),
  )
  const levels = numberedLevels(accepted.map((block) => block.shape!))

  return accepted.map((block) => {
    const shape = block.shape!
    return {
      text: textOf(block.node) ?? '',
      level: shape.number ? (levels.get(shape.number) ?? 1) : 1,
      pos: block.pos,
      inferred: true as const,
      basis: shape.basis,
    }
  })
}

/** the margin class the outline view gives an inferred heading */
export function structuralHeadingClass(level: number): string {
  return `inferred-heading inferred-h${Math.min(Math.max(level, 1), 6)}`
}

/** the marker the story view shows in the gutter: faded, like the heading */
export function structuralMarker(level: number): string {
  return '#'.repeat(Math.min(Math.max(level, 1), 6))
}
