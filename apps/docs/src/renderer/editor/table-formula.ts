/**
 * Word's formulas: a `=` field, e.g. `=SUM(ABOVE) \# "0.00"` in a table cell
 * or `=(2+3)*4` in body text. The field's text is its cached result; the
 * instruction is what Word reads and recomputes. Only what is evaluated here
 * is ever rewritten: a formula outside this subset keeps the result it was
 * opened with.
 */
import { Extension, type Editor } from '@tiptap/core'
import type { Node as PmNode } from '@tiptap/pm/model'
import { Plugin, PluginKey, type EditorState, type Transaction } from '@tiptap/pm/state'
import { TableMap } from '@tiptap/pm/tables'
import type { Block } from '@genoffice/docx-engine'
import { bookmarkRange } from './bookmark-range'
import { collectFields, isFieldLocked, type FieldRange } from './field-codes'
import { textOfXml } from './field-results'
import { TRACK_IGNORE } from './revisions'

/** the functions the dialog offers, each evaluated below */
export const FORMULA_FUNCTIONS = [
  'ABS',
  'AND',
  'AVERAGE',
  'COUNT',
  'DEFINED',
  'FALSE',
  'IF',
  'INT',
  'MAX',
  'MIN',
  'MOD',
  'NOT',
  'OR',
  'PRODUCT',
  'ROUND',
  'SIGN',
  'SUM',
  'TRUE',
] as const

/** the number formats Word's Formula dialog lists */
export const FORMULA_NUMBER_FORMATS = [
  '#,##0',
  '#,##0.00',
  '$#,##0.00;($#,##0.00)',
  '0',
  '0%',
  '0.00',
  '0.00%',
] as const

const DIRECTIONS = ['ABOVE', 'BELOW', 'LEFT', 'RIGHT'] as const
type Direction = (typeof DIRECTIONS)[number]

/** the table around a formula: cell texts by grid slot, null where a merged cell continues */
export interface FormulaGrid {
  rows: number
  cols: number
  text(row: number, col: number): string | null
}

export interface FormulaCell {
  row: number
  col: number
}

/** a formula in body text has no cells around it: a cell reference or a direction does not evaluate */
export const NO_TABLE: FormulaGrid = { rows: 0, cols: 0, text: () => null }
const NO_CELL: FormulaCell = { row: -1, col: -1 }

/** what a formula reads outside its own table, by upper-cased bookmark name */
export interface FormulaContext {
  /** the text a bookmark marks; null when there is no such bookmark */
  bookmark(name: string): string | null
  /** the table a bookmark sits in, for `SUM(Name A1:B2)` */
  table(name: string): FormulaGrid | null
}

export function isFormulaInstr(instr: string): boolean {
  return instr.trimStart().startsWith('=')
}

/** `=SUM(ABOVE) \# "0.00"` → the expression and the number picture */
export function parseFormulaInstr(instr: string): { expr: string; format: string | null } | null {
  const text = instr.trim()
  if (!text.startsWith('=')) return null
  const body = text.slice(1)
  const at = body.search(/\\#/)
  if (at < 0) return /\\[*@!]/.test(body) ? null : { expr: body.trim(), format: null }
  const picture = /^\\#\s*(?:"([^"]*)"|(\S+))\s*$/.exec(body.slice(at))
  if (!picture) return null
  return { expr: body.slice(0, at).trim(), format: picture[1] ?? picture[2] ?? '' }
}

export function formulaInstr(expr: string, format: string): string {
  const body = expr.trim().replace(/^=?\s*/, '=')
  return format.trim() ? `${body} \\# "${format.trim()}"` : body
}

/** the first number a cell shows: `$1,234.50` is 1234.5, `(12)` is -12; null for none */
export function cellNumber(text: string): number | null {
  const match = /(\()?\s*([-−]?)\s*[$€£¥]?\s*(\d[\d,]*(?:\.\d+)?|\.\d+)/.exec(text)
  if (!match) return null
  const value = Number(match[3].replace(/,/g, ''))
  if (!Number.isFinite(value)) return null
  const negative =
    match[2] !== '' || (match[1] !== undefined && /^\s*%?\s*\)/.test(restAfter(text, match)))
  return negative ? -value : value
}

function restAfter(text: string, match: RegExpExecArray): string {
  return text.slice(match.index + match[0].length)
}

type Token =
  { kind: 'num'; value: number } | { kind: 'name'; value: string } | { kind: 'op'; value: string }

function tokenize(expr: string): Token[] | null {
  const tokens: Token[] = []
  const pattern =
    /\s*(?:(\d+(?:\.\d+)?|\.\d+)|([\p{L}_][\p{L}\p{N}_]*)|(<>|<=|>=|[-+*/^(),;:=<>]))/uy
  let at = 0
  while (at < expr.length) {
    if (/^\s*$/.test(expr.slice(at))) break
    pattern.lastIndex = at
    const match = pattern.exec(expr)
    if (!match) return null
    if (match[1] !== undefined) tokens.push({ kind: 'num', value: Number(match[1]) })
    else if (match[2] !== undefined) tokens.push({ kind: 'name', value: match[2].toUpperCase() })
    else tokens.push({ kind: 'op', value: match[3] })
    at = pattern.lastIndex
  }
  return tokens
}

const CELL_REF = /^([A-Z]{1,2})(\d+)$/
const COLUMN_REF = /^[A-Z]{1,2}$/

function columnIndex(letters: string): number {
  let col = 0
  for (const letter of letters) col = col * 26 + (letter.charCodeAt(0) - 64)
  return col - 1
}

function cellRef(name: string): FormulaCell | null {
  const match = CELL_REF.exec(name)
  if (!match) return null
  const row = Number(match[2])
  return row >= 1 ? { row: row - 1, col: columnIndex(match[1]) } : null
}

class Unsupported extends Error {}

/**
 * The cells a direction names, nearest first. As in Word, the run ends at the
 * first empty cell, and a cell without a number adds nothing.
 */
function directionValues(grid: FormulaGrid, self: FormulaCell, direction: Direction): number[] {
  const step =
    direction === 'ABOVE'
      ? [-1, 0]
      : direction === 'BELOW'
        ? [1, 0]
        : direction === 'LEFT'
          ? [0, -1]
          : [0, 1]
  const values: number[] = []
  let row = self.row + step[0]
  let col = self.col + step[1]
  while (row >= 0 && row < grid.rows && col >= 0 && col < grid.cols) {
    const text = grid.text(row, col)
    if (text !== null) {
      if (text.trim() === '') break
      const value = cellNumber(text)
      if (value !== null) values.push(value)
    }
    row += step[0]
    col += step[1]
  }
  return values
}

const COMPARISONS: Record<string, (a: number, b: number) => boolean> = {
  '=': (a, b) => a === b,
  '<>': (a, b) => a !== b,
  '<': (a, b) => a < b,
  '<=': (a, b) => a <= b,
  '>': (a, b) => a > b,
  '>=': (a, b) => a >= b,
}

// an error inside a value (a division by zero) travels as NaN, so the branch
// IF does not take cannot fail the formula
function evaluate(
  tokens: Token[],
  grid: FormulaGrid,
  self: FormulaCell,
  context: FormulaContext | undefined,
): number {
  let at = 0
  const peek = () => tokens[at]
  const isOp = (value: string) => peek()?.kind === 'op' && peek()?.value === value
  const expect = (value: string) => {
    if (!isOp(value)) throw new Unsupported()
    at++
  }
  const endsArgument = (token: Token | undefined) =>
    !token || (token.kind === 'op' && /^[,;)]$/.test(token.value))

  /** a cell, a range, whole rows or whole columns of `from` at the cursor; null when there is none */
  const reference = (from: FormulaGrid, own: FormulaCell | null): number[] | null => {
    const first = tokens[at]
    const colon = tokens[at + 1]
    const last = tokens[at + 2]
    const ranged = colon?.kind === 'op' && colon.value === ':'
    const block = (top: number, left: number, bottom: number, right: number): number[] => {
      if (top < 0 || left < 0 || bottom >= from.rows || right >= from.cols) throw new Unsupported()
      const values: number[] = []
      for (let row = top; row <= bottom; row++) {
        for (let col = left; col <= right; col++) {
          if (own && row === own.row && col === own.col) continue
          const text = from.text(row, col)
          const value = text === null ? null : cellNumber(text)
          if (value !== null) values.push(value)
        }
      }
      return values
    }
    if (first?.kind === 'num' && ranged && last?.kind === 'num') {
      if (!Number.isInteger(first.value) || !Number.isInteger(last.value)) throw new Unsupported()
      at += 3
      const [top, bottom] = [first.value, last.value].sort((a, b) => a - b)
      return block(top - 1, 0, bottom - 1, from.cols - 1)
    }
    if (first?.kind !== 'name') return null
    if (ranged && last?.kind === 'name') {
      if (COLUMN_REF.test(first.value) && COLUMN_REF.test(last.value)) {
        at += 3
        const [left, right] = [columnIndex(first.value), columnIndex(last.value)].sort(
          (a, b) => a - b,
        )
        return block(0, left, from.rows - 1, right)
      }
      const a = cellRef(first.value)
      const b = cellRef(last.value)
      if (!a || !b) return null
      at += 3
      return block(
        Math.min(a.row, b.row),
        Math.min(a.col, b.col),
        Math.max(a.row, b.row),
        Math.max(a.col, b.col),
      )
    }
    const cell = cellRef(first.value)
    if (!cell || !endsArgument(colon)) return null
    const outside = cell.row >= from.rows || cell.col >= from.cols
    if (outside && from === grid && (context?.bookmark(first.value) ?? null) !== null) return null
    at++
    return block(cell.row, cell.col, cell.row, cell.col)
  }

  /** one function argument: a direction, a range or a cell lists its numbers; anything else is a value */
  const argument = (): number[] => {
    const token = peek()
    if (token?.kind === 'name') {
      const direction = DIRECTIONS.find((name) => name === token.value)
      if (direction) {
        if (grid.rows === 0) throw new Unsupported()
        at++
        return directionValues(grid, self, direction)
      }
      // `Sales A1:B2`: cells of the table the bookmark Sales sits in
      const other = tokens[at + 1]?.kind !== 'op' ? context?.table(token.value) : null
      if (other) {
        at++
        const values = reference(other, null)
        if (!values) throw new Unsupported()
        return values
      }
    }
    return reference(grid, self) ?? [expression()]
  }

  /** DEFINED's argument: whether it evaluates, leaving the cursor on the closing parenthesis */
  const defined = (): boolean => {
    const start = at
    let ok: boolean
    try {
      ok = !Number.isNaN(expression()) && isOp(')')
    } catch (error) {
      if (!(error instanceof Unsupported)) throw error
      ok = false
    }
    if (ok) return true
    let depth = 0
    for (at = start; at < tokens.length; at++) {
      if (isOp('(')) depth++
      else if (isOp(')') && depth-- === 0) return false
    }
    throw new Unsupported()
  }

  const call = (name: string): number => {
    expect('(')
    if (name === 'DEFINED') {
      const ok = defined()
      expect(')')
      return ok ? 1 : 0
    }
    const groups: number[][] = []
    if (!isOp(')')) {
      groups.push(argument())
      while (isOp(',') || isOp(';')) {
        at++
        groups.push(argument())
      }
    }
    expect(')')
    const all = groups.flat()
    const scalar = (index: number) => {
      const group = groups[index]
      if (!group) throw new Unsupported()
      return group.length === 1 ? group[0] : (group[0] ?? 0)
    }
    const arity = (count: number) => {
      if (groups.length !== count) throw new Unsupported()
    }
    const truth = (value: boolean, ...of: number[]) =>
      of.some(Number.isNaN) ? Number.NaN : value ? 1 : 0
    switch (name) {
      case 'SUM':
        return all.reduce((sum, value) => sum + value, 0)
      case 'PRODUCT':
        return all.length === 0 ? 0 : all.reduce((product, value) => product * value, 1)
      case 'COUNT':
        return all.length
      case 'AVERAGE':
        return all.reduce((sum, value) => sum + value, 0) / all.length
      case 'MIN':
        return all.length === 0 ? 0 : Math.min(...all)
      case 'MAX':
        return all.length === 0 ? 0 : Math.max(...all)
      case 'ABS':
        return Math.abs(scalar(0))
      case 'INT':
        return Math.trunc(scalar(0))
      case 'SIGN':
        return Math.sign(scalar(0))
      case 'MOD':
        return scalar(0) % scalar(1)
      case 'ROUND': {
        const scale = 10 ** Math.trunc(scalar(1))
        return (Math.sign(scalar(0)) * Math.round(Math.abs(scalar(0)) * scale)) / scale
      }
      case 'IF':
        arity(3)
        return Number.isNaN(scalar(0)) ? Number.NaN : scalar(0) !== 0 ? scalar(1) : scalar(2)
      case 'AND':
        arity(2)
        return truth(scalar(0) !== 0 && scalar(1) !== 0, scalar(0), scalar(1))
      case 'OR':
        arity(2)
        return truth(scalar(0) !== 0 || scalar(1) !== 0, scalar(0), scalar(1))
      case 'NOT':
        arity(1)
        return truth(scalar(0) === 0, scalar(0))
      case 'TRUE':
        arity(0)
        return 1
      case 'FALSE':
        arity(0)
        return 0
      default:
        throw new Unsupported()
    }
  }

  const primary = (): number => {
    const token = peek()
    if (!token) throw new Unsupported()
    if (token.kind === 'num') {
      at++
      return token.value
    }
    if (token.kind === 'name') {
      at++
      if (isOp('(')) return call(token.value)
      if (token.value === 'TRUE') return 1
      if (token.value === 'FALSE') return 0
      const cell = cellRef(token.value)
      // a name shaped like a cell the table does not have can only be a bookmark
      if (cell && cell.row < grid.rows && cell.col < grid.cols) {
        if (cell.row === self.row && cell.col === self.col) return 0
        const text = grid.text(cell.row, cell.col)
        return (text === null ? null : cellNumber(text)) ?? 0
      }
      const marked = context?.bookmark(token.value)
      if (marked === null || marked === undefined) throw new Unsupported()
      return cellNumber(marked) ?? 0
    }
    if (token.value === '(') {
      at++
      const value = expression()
      expect(')')
      return value
    }
    throw new Unsupported()
  }
  const unary = (): number => {
    if (isOp('-')) {
      at++
      return -unary()
    }
    if (isOp('+')) {
      at++
      return unary()
    }
    return primary()
  }
  const power = (): number => {
    let value = unary()
    while (isOp('^')) {
      at++
      value = value ** unary()
    }
    return value
  }
  const term = (): number => {
    let value = power()
    while (isOp('*') || isOp('/')) {
      const divide = isOp('/')
      at++
      const right = power()
      value = divide ? (right === 0 ? Number.NaN : value / right) : value * right
    }
    return value
  }
  const sum = (): number => {
    let value = term()
    while (isOp('+') || isOp('-')) {
      const minus = isOp('-')
      at++
      const right = term()
      value = minus ? value - right : value + right
    }
    return value
  }
  function expression(): number {
    const left = sum()
    const token = peek()
    const compare = token?.kind === 'op' ? COMPARISONS[token.value] : undefined
    if (!compare) return left
    at++
    const right = sum()
    return Number.isNaN(left) || Number.isNaN(right) ? Number.NaN : compare(left, right) ? 1 : 0
  }

  const result = expression()
  if (at !== tokens.length || !Number.isFinite(result)) throw new Unsupported()
  return result
}

/** the formula's value in the cell `self`; null when it is outside what is evaluated here */
export function evaluateFormula(
  expr: string,
  grid: FormulaGrid,
  self: FormulaCell,
  context?: FormulaContext,
): number | null {
  const tokens = tokenize(expr)
  if (!tokens || tokens.length === 0) return null
  try {
    return evaluate(tokens, grid, self, context)
  } catch (error) {
    if (error instanceof Unsupported) return null
    throw error
  }
}

function groupThousands(digits: string): string {
  return digits.replace(/\B(?=(\d{3})+(?!\d))/g, ',')
}

function formatSection(value: number, section: string): string {
  const first = section.search(/[0#]/)
  if (first < 0) return section
  let last = first
  for (let i = first; i < section.length; i++) if (/[0#,.]/.test(section[i])) last = i
  while (last > first && /[,.]/.test(section[last])) last--
  const prefix = section.slice(0, first)
  const suffix = section.slice(last + 1)
  const [whole, fraction = ''] = section.slice(first, last + 1).split('.')
  const required = (fraction.match(/0/g) ?? []).length
  const places = fraction.replace(/[^0#]/g, '').length
  let [intPart, fracPart = ''] = Math.abs(value).toFixed(places).split('.')
  while (fracPart.length > required && fracPart.endsWith('0')) fracPart = fracPart.slice(0, -1)
  intPart = intPart.padStart((whole.match(/0/g) ?? []).length, '0')
  if (!whole.includes('0') && intPart === '0' && fracPart) intPart = ''
  if (whole.includes(',')) intPart = groupThousands(intPart)
  return prefix + intPart + (fracPart ? `.${fracPart}` : '') + suffix
}

/**
 * Word's numeric picture (`\#`): `0` and `#` digit places, `,` grouping, a
 * decimal point, literal text around them, and `positive;negative;zero`
 * sections. A `%` is a literal, as in Word.
 */
export function formatNumberPicture(value: number, picture: string): string {
  const sections = picture.split(';')
  if (value === 0 && sections[2] !== undefined) return formatSection(0, sections[2])
  if (value < 0 && sections[1] !== undefined) return formatSection(value, sections[1])
  const text = formatSection(value, sections[0])
  return value < 0 && /[1-9]/.test(text) ? `-${text}` : text
}

/** without a picture Word shows up to two decimals */
export function formatFormulaResult(value: number, format: string | null): string {
  if (format) return formatNumberPicture(value, format)
  const text = String(Math.round((Math.abs(value) + Number.EPSILON) * 100) / 100)
  return value < 0 && /[1-9]/.test(text) ? `-${text}` : text
}

/** the text a formula field shows; null when the instruction is not one evaluated here */
export function formulaResult(
  instr: string,
  grid: FormulaGrid,
  self: FormulaCell,
  context?: FormulaContext,
): string | null {
  const parsed = parseFormulaInstr(instr)
  if (!parsed) return null
  const value = evaluateFormula(parsed.expr, grid, self, context)
  return value === null ? null : formatFormulaResult(value, parsed.format)
}

const isCell = (node: PmNode) =>
  node.type.name === 'docTableCell' || node.type.name === 'docTableHeader'

function gridOf(table: PmNode): { grid: FormulaGrid; map: TableMap } {
  const map = TableMap.get(table)
  const grid: FormulaGrid = {
    rows: map.height,
    cols: map.width,
    text(row, col) {
      const index = row * map.width + col
      const offset = map.map[index]
      const continued =
        (col > 0 && map.map[index - 1] === offset) ||
        (row > 0 && map.map[index - map.width] === offset)
      const cell = continued ? null : table.nodeAt(offset)
      if (!cell || cell.attrs.gridGap || cell.attrs.mergeGap) return null
      return cell.textBetween(0, cell.content.size, ' ')
    },
  }
  return { grid, map }
}

let sourceBlocks: () => Block[] = () => []

/** the open document's parsed blocks: a bookmark read from the file has its exact range there */
export function setFormulaSourceBlocks(source: () => Block[]): void {
  sourceBlocks = source
}

const escapeRegExp = (text: string): string => text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')

/**
 * The text between a bookmark's start and end in a block's XML, when both sit
 * in one paragraph that still reads `current`; null otherwise.
 */
function markedText(xml: string, name: string, current: string): string | null {
  const start = new RegExp(
    `<w:bookmarkStart\\b(?=[^>]*\\bw:name="${escapeRegExp(name)}")[^>]*\\bw:id="([^"]*)"[^>]*>`,
    'i',
  ).exec(xml)
  if (!start) return null
  const from = start.index + start[0].length
  const end = new RegExp(`<w:bookmarkEnd\\b[^>]*\\bw:id="${escapeRegExp(start[1])}"[^>]*>`).exec(
    xml.slice(from),
  )
  if (!end) return null
  const inner = xml.slice(from, from + end.index)
  if (inner.includes('</w:p>')) return null
  let open = -1
  for (const match of xml.slice(0, start.index).matchAll(/<w:p[ >]/g)) open = match.index
  const close = xml.indexOf('</w:p>', from + end.index)
  if (open < 0 || close < 0 || textOfXml(xml.slice(open, close)) !== current) return null
  // a bookmark set at a point covers nothing: its paragraph stands for it
  return textOfXml(inner) || null
}

/**
 * The document's bookmarks as formula operands. A name stands for the text
 * the bookmark covers when the file still tells (the paragraph reads as it
 * did when it was opened); otherwise for its paragraph.
 */
export function bookmarkContext(doc: PmNode, blocks: Block[] = sourceBlocks()): FormulaContext {
  let hosts: Map<string, number> | null = null
  const hostOf = (name: string): number | undefined => {
    if (!hosts) {
      hosts = new Map()
      doc.descendants((node, pos) => {
        const names = node.attrs?.bookmarks as string[] | null | undefined
        if (Array.isArray(names))
          for (const each of names)
            if (!hosts!.has(each.toUpperCase())) hosts!.set(each.toUpperCase(), pos)
        return !node.isLeaf
      })
    }
    return hosts.get(name)
  }
  const sourceXml = (pos: number): string | null => {
    const $pos = doc.resolve(pos)
    const top = $pos.depth === 0 ? doc.nodeAt(pos) : $pos.node(1)
    if (!top) return null
    if (top.attrs.genXml) return String(top.attrs.genXml)
    const index = top.attrs.docxIndex as number | null
    if (typeof index !== 'number') return null
    return blocks.find((block) => block.docxIndex === index)?.originalXml ?? null
  }
  return {
    bookmark(name) {
      // a bookmark made here from a selection: the text it was made from
      const range = bookmarkRange(doc, name)
      if (range) return doc.textBetween(range.from, range.to, ' ')
      const pos = hostOf(name)
      const host = pos === undefined ? null : doc.nodeAt(pos)
      if (pos === undefined || !host) return null
      const xml = sourceXml(pos)
      return (xml ? markedText(xml, name, host.textContent) : null) ?? host.textContent
    },
    table(name) {
      const pos = hostOf(name)
      if (pos === undefined) return null
      const $pos = doc.resolve(pos)
      for (let depth = $pos.depth; depth > 0; depth--)
        if ($pos.node(depth).type.name === 'docTable') return gridOf($pos.node(depth)).grid
      return null
    },
  }
}

interface FormulaSite {
  table: PmNode
  tablePos: number
  self: FormulaCell
}

/** the table cell holding `pos`, when it sits in a table of the body */
function formulaSite(doc: PmNode, pos: number): FormulaSite | null {
  const $pos = doc.resolve(pos)
  for (let depth = $pos.depth; depth > 1; depth--) {
    if (!isCell($pos.node(depth))) continue
    const table = $pos.node(depth - 2)
    if (table.type.name !== 'docTable') return null
    const tablePos = $pos.before(depth - 2)
    const rect = TableMap.get(table).findCell($pos.before(depth) - tablePos - 1)
    return { table, tablePos, self: { row: rect.top, col: rect.left } }
  }
  return null
}

/** a formula may read another formula's cell, so results settle over a few passes */
const MAX_PASSES = 4

/**
 * Recomputes the formula fields of the tables `within` selects (all of them
 * by default; null stands for the formulas outside a table), on `tr`. Returns
 * how many results changed.
 */
export function refreshTableFormulas(
  tr: Transaction,
  within: (tablePos: number | null) => boolean = () => true,
): number {
  const changed = new Set<number>()
  for (let pass = 0; pass < MAX_PASSES; pass++) {
    const doc = tr.doc
    const context = bookmarkContext(doc)
    const jobs: Array<{ field: FieldRange; text: string; index: number; differs: boolean }> = []
    collectFields(doc).forEach((field, index) => {
      if (!isFormulaInstr(field.instr) || isFieldLocked(field.mark)) return
      const site = formulaSite(doc, field.from)
      if (!within(site ? site.tablePos : null)) return
      const text = site
        ? formulaResult(field.instr, gridOf(site.table).grid, site.self, context)
        : formulaResult(field.instr, NO_TABLE, NO_CELL, context)
      if (text === null) return
      const differs = text !== doc.textBetween(field.from, field.to, '')
      // a result computed here is what the file carries: the field is not left for Word to update
      if (differs || field.mark.attrs.dirty === true) jobs.push({ field, text, index, differs })
    })
    if (jobs.length === 0) break
    for (const { field, text, index, differs } of jobs.reverse()) {
      const marks = doc.nodeAt(field.from)?.marks ?? []
      const mark = field.mark.type.create({ ...field.mark.attrs, dirty: false })
      tr.replaceWith(
        field.from,
        field.to,
        doc.type.schema.text(text, [...field.mark.type.removeFromSet(marks), mark]),
      )
      if (differs) changed.add(index)
    }
  }
  return changed.size
}

export type FormulaRefusal = 'notInCell' | 'unsupported' | 'outsideTable'

/** what the Formula dialog opens with: Word proposes a sum of the numbers above, else of those to the left */
export function suggestedFormula(state: EditorState): string {
  const site = formulaSite(state.doc, state.selection.from)
  if (!site) return '='
  const { grid } = gridOf(site.table)
  if (directionValues(grid, site.self, 'ABOVE').length > 0) return '=SUM(ABOVE)'
  if (directionValues(grid, site.self, 'LEFT').length > 0) return '=SUM(LEFT)'
  return '='
}

/** the formula field at the caret, for the dialog to edit instead of adding a second one */
export function formulaFieldAt(state: EditorState): FieldRange | null {
  const { from, to } = state.selection
  return (
    collectFields(state.doc).find(
      (field) => isFormulaInstr(field.instr) && field.from <= from && to <= field.to,
    ) ?? null
  )
}

/** whether Formula has a place for its field: a table cell, or text of the body */
export function formulaPlace(state: EditorState): boolean {
  const { $from, $to } = state.selection
  return (
    formulaSite(state.doc, $from.pos) !== null ||
    ($from.parent.isTextblock && $to.parent.isTextblock)
  )
}

/**
 * Puts a formula field at the caret, in a table cell or in body text
 * (replacing the formula field the caret is in), showing its result. Null
 * when it did.
 */
export function insertTableFormula(editor: Editor, instr: string): FormulaRefusal | null {
  const { state } = editor
  const markType = state.schema.marks.instrField
  const existing = formulaFieldAt(state)
  const from = existing?.from ?? state.selection.from
  const to = existing?.to ?? state.selection.to
  const site = formulaSite(state.doc, from)
  const $from = state.doc.resolve(from)
  const $to = state.doc.resolve(to)
  if (!markType || !(site ? $from.sameParent($to) : formulaPlace(state))) return 'notInCell'
  const context = bookmarkContext(state.doc)
  const text = site
    ? formulaResult(instr, gridOf(site.table).grid, site.self, context)
    : formulaResult(instr, NO_TABLE, NO_CELL, context)
  if (text === null) return site ? 'unsupported' : 'outsideTable'
  const marks = existing
    ? markType.removeFromSet(state.doc.nodeAt(from)?.marks ?? [])
    : markType.removeFromSet(state.storedMarks ?? state.doc.resolve(from).marks())
  const mark = markType.create({ instr: instr.trim(), dirty: false })
  const tr = state.tr.replaceWith(from, to, state.schema.text(text, [...marks, mark]))
  if (site) refreshTableFormulas(tr, (tablePos) => tablePos === site.tablePos)
  editor.view.dispatch(tr.scrollIntoView())
  return null
}

const tableFormulaKey = new PluginKey('tableFormula')

/** the tables an edit touched from the inside or replaced one at a time: loading a document is not one */
function editedTables(transactions: readonly Transaction[], doc: PmNode): Set<number> {
  const tables = new Set<number>()
  transactions.forEach((tr, index) => {
    if (!tr.docChanged || tr.getMeta('addToHistory') === false) return
    const later = transactions.slice(index + 1)
    tr.mapping.maps.forEach((map, step) => {
      map.forEach((_oldFrom, _oldTo, newFrom, newTo) => {
        let from = tr.mapping.slice(step + 1).map(newFrom, -1)
        let to = tr.mapping.slice(step + 1).map(newTo, 1)
        for (const next of later) {
          from = next.mapping.map(from, -1)
          to = next.mapping.map(to, 1)
        }
        if (from > doc.content.size || to > doc.content.size) return
        const $from = doc.resolve(from)
        const $to = doc.resolve(to)
        const whole = $from.nodeAfter
        if (whole?.type.name === 'docTable' && from + whole.nodeSize === to) tables.add(from)
        for (let depth = $from.depth; depth > 0; depth--) {
          if ($from.node(depth).type.name !== 'docTable') continue
          const pos = $from.before(depth)
          if ($to.depth >= depth && $to.before(depth) === pos) tables.add(pos)
          break
        }
      })
    })
  })
  return tables
}

/** a table's formulas follow its cells: an edit inside a table recomputes that table's fields */
export const TableFormulaExtension = Extension.create({
  name: 'tableFormula',
  addProseMirrorPlugins() {
    return [
      new Plugin({
        key: tableFormulaKey,
        appendTransaction(transactions, _oldState, newState) {
          // undo and redo restore what the user had
          if (transactions.some((tr) => tr.getMeta(tableFormulaKey) || tr.getMeta('history$')))
            return null
          const tables = editedTables(transactions, newState.doc)
          if (tables.size === 0) return null
          const tr = newState.tr
          refreshTableFormulas(tr, (pos) => pos !== null && tables.has(pos))
          if (!tr.docChanged) return null
          return tr.setMeta(tableFormulaKey, true).setMeta(TRACK_IGNORE, true)
        },
      }),
    ]
  },
})
