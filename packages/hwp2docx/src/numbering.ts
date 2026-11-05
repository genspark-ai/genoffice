/**
 * Hangul paragraph heads (hh:heading: bullets, paragraph numbers, outline
 * numbers) written as the literal text Hangul shows, like the "- " or "가."
 * in front of a line. The text stays editable but no longer renumbers itself;
 * the conversion reports that as `auto-numbering-static`.
 */

import { attr, children, descendants, elements, intAttr, textOf, type XNode } from './xml'

interface ParaHead {
  start: number
  numFormat: string
  /** template, ^n = the counter of level n */
  text: string
  /** autoIndent: the following lines start where the text after the head does */
  autoIndent: boolean
  /** gap between the head and the text: a share of the glyph size (PERCENT) or HWPUNIT */
  textOffset: { percent: boolean; value: number }
}

/** a paragraph head as it shows in front of the text */
export interface Head {
  text: string
  autoIndent: boolean
  textOffset: { percent: boolean; value: number }
}

export interface HeadingTables {
  bullet(id: string | undefined): Head | undefined
  level(numberingId: string | undefined, level: number): ParaHead | undefined
}

function headLayout(node: XNode | undefined): Omit<Head, 'text'> {
  return {
    autoIndent: attr(node, 'autoIndent') === '1',
    textOffset: {
      percent: (attr(node, 'textOffsetType') ?? 'PERCENT') === 'PERCENT',
      value: node ? intAttr(node, 'textOffset', 50) : 50,
    },
  }
}

export function readHeadings(root: XNode[]): HeadingTables {
  const bullets = new Map<string, Head>()
  for (const node of descendants(root, 'hh:bullet')) {
    const id = attr(node, 'id')
    const char = attr(node, 'char')
    const layout = headLayout(elements(node).find((child) => child.name === 'hh:paraHead')?.node)
    if (id !== undefined && char) bullets.set(id, { text: char, ...layout })
  }
  const numberings = new Map<string, Map<number, ParaHead>>()
  for (const node of descendants(root, 'hh:numbering')) {
    const id = attr(node, 'id')
    if (id === undefined) continue
    const levels = new Map<number, ParaHead>()
    for (const head of elements(node)) {
      if (head.name !== 'hh:paraHead') continue
      levels.set(intAttr(head.node, 'level', 1), {
        start: intAttr(head.node, 'start', 1),
        numFormat: attr(head.node, 'numFormat') ?? 'DIGIT',
        text: textContent(head.node),
        ...headLayout(head.node),
      })
    }
    numberings.set(id, levels)
  }
  return {
    bullet: (id) => (id !== undefined ? bullets.get(id) : undefined),
    level: (id, level) => (id !== undefined ? numberings.get(id)?.get(level) : undefined),
  }
}

function textContent(node: XNode): string {
  return children(node)
    .map((child) => textOf(child) ?? '')
    .join('')
}

/** running counters per numbering, one per level (1-based) */
export class HeadingCounter {
  private counters = new Map<string, number[]>()

  constructor(private readonly tables: HeadingTables) {}

  /** the head of a paragraph, advancing the counters; undefined when there is none */
  next(
    kind: 'bullet' | 'number' | 'outline',
    idRef: string | undefined,
    level0: number,
  ): Head | undefined {
    if (kind === 'bullet') return this.tables.bullet(idRef)
    const level = Math.min(10, Math.max(1, level0 + 1))
    const head = this.tables.level(idRef, level)
    if (!head || idRef === undefined) return undefined
    const counts = this.counters.get(idRef) ?? []
    for (let l = 1; l < level; l++) counts[l] ??= (this.tables.level(idRef, l)?.start ?? 1) - 1
    counts[level] = (counts[level] ?? head.start - 1) + 1
    counts.length = level + 1 // deeper levels restart
    this.counters.set(idRef, counts)
    // a caret without a level number shows nothing ("^  □" draws "  □")
    const text = head.text.replace(/\^(\d*)/g, (_, n: string) => {
      if (!n) return ''
      const l = Number(n)
      const fmt = this.tables.level(idRef, l)?.numFormat ?? 'DIGIT'
      return formatNumber(Math.max(1, counts[l] ?? 1), fmt)
    })
    return text ? { text, autoIndent: head.autoIndent, textOffset: head.textOffset } : undefined
  }
}

const HANGUL_SYLLABLES = '가나다라마바사아자차카타파하'
const HANGUL_JAMO = 'ㄱㄴㄷㄹㅁㅂㅅㅇㅈㅊㅋㅌㅍㅎ'
const CIRCLED_HANGUL_SYLLABLES = '㉮㉯㉰㉱㉲㉳㉴㉵㉶㉷㉸㉹㉺㉻'
const CIRCLED_HANGUL_JAMO = '㉠㉡㉢㉣㉤㉥㉦㉧㉨㉩㉪㉫㉬㉭'
const HANJA_DIGITS = '一二三四五六七八九十' // public-hygiene: allow

function roman(n: number): string {
  const table: Array<[number, string]> = [
    [1000, 'M'],
    [900, 'CM'],
    [500, 'D'],
    [400, 'CD'],
    [100, 'C'],
    [90, 'XC'],
    [50, 'L'],
    [40, 'XL'],
    [10, 'X'],
    [9, 'IX'],
    [5, 'V'],
    [4, 'IV'],
    [1, 'I'],
  ]
  let out = ''
  let rest = n
  for (const [value, glyph] of table) {
    while (rest >= value) {
      out += glyph
      rest -= value
    }
  }
  return out
}

const cycle = (set: string, n: number): string => [...set][(n - 1) % [...set].length]

export function formatNumber(n: number, format: string): string {
  switch (format) {
    case 'CIRCLED_DIGIT':
      return n <= 20 ? String.fromCodePoint(0x2460 + n - 1) : String(n)
    case 'ROMAN_CAPITAL':
      return roman(n)
    case 'ROMAN_SMALL':
      return roman(n).toLowerCase()
    case 'LATIN_CAPITAL':
      return String.fromCharCode(65 + ((n - 1) % 26))
    case 'LATIN_SMALL':
      return String.fromCharCode(97 + ((n - 1) % 26))
    case 'CIRCLED_LATIN_SMALL':
      return n <= 26 ? String.fromCodePoint(0x24d0 + n - 1) : String(n)
    case 'HANGUL_SYLLABLE':
      return cycle(HANGUL_SYLLABLES, n)
    case 'HANGUL_JAMO':
      return cycle(HANGUL_JAMO, n)
    case 'CIRCLED_HANGUL_SYLLABLE':
      return cycle(CIRCLED_HANGUL_SYLLABLES, n)
    case 'CIRCLED_HANGUL_JAMO':
      return cycle(CIRCLED_HANGUL_JAMO, n)
    case 'HANJA_DIGIT':
      return n <= 10 ? [...HANJA_DIGITS][n - 1] : String(n)
    default:
      return String(n)
  }
}
