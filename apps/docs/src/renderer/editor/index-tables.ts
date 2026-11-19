import {
  fieldSwitchArg,
  formatFieldInstruction,
  parseFieldInstruction,
  withFieldSwitch,
  type FieldDisplay,
  type FieldInstruction,
  type FieldTableLeader,
} from '@genoffice/docx-engine'

export type IndexTablesKind = 'index' | 'toc' | 'figures' | 'authorities'

/** What a generated table carries besides its entries: the field instruction and its tab leader. */
export interface TableSpec {
  instr: string
  leader: FieldTableLeader
}

export const TABLE_LEADERS: readonly FieldTableLeader[] = ['none', 'dot', 'hyphen', 'underscore']

/** the dialog's four leaders for whatever a tab stop carries; an unknown one draws as dots */
export function tableLeaderOf(leader: FieldDisplay['leader']): FieldTableLeader {
  if (leader === 'heavy') return 'underscore'
  if (leader === 'none' || leader === 'hyphen' || leader === 'underscore') return leader
  return 'dot'
}

/** one row of Options' style table: paragraphs in that style are entries of that level (TOC \t) */
export interface TocStyleLevel {
  name: string
  level: number
}

export interface TocSources {
  /** the styles listed besides the built-in headings (\t) */
  styles: TocStyleLevel[]
  /** paragraphs with an outline level of their own (\u) */
  outline: boolean
  /** TC fields (\f) */
  entryFields: boolean
}

export interface TocOptions extends TocSources {
  /** deepest heading level listed; 0 when the built-in headings are not a source */
  levels: number
  pageNumbers: boolean
  rightAlign: boolean
  leader: FieldTableLeader
  hyperlinks: boolean
}

export interface FiguresOptions {
  label: string
  /** the paragraph style listed in place of the label's captions (\t); '' for the captions */
  style: string
  /** the TC fields listed too, by their identifier letter (\f); '' for none */
  entryId: string
  includeLabel: boolean
  pageNumbers: boolean
  rightAlign: boolean
  leader: FieldTableLeader
  hyperlinks: boolean
}

export interface IndexOptions {
  runIn: boolean
  /** entries that begin with an accented letter get a heading of their own (\a) */
  accented: boolean
  rightAlign: boolean
  leader: FieldTableLeader
  /** the column count the field carries (\\c), one column when it has none */
  columns: number
}

export interface AuthoritiesOptions {
  /** 0: every category */
  category: number
  passim: boolean
  /** a citation is listed with the bold, italic and underline it was marked with (off: \\f) */
  keepFormatting: boolean
  leader: FieldTableLeader
}

export const DEFAULT_SPECS: Record<IndexTablesKind, TableSpec> = {
  toc: { instr: ' TOC \\o "1-3" \\h \\z \\u ', leader: 'dot' },
  figures: { instr: ' TOC \\h \\z \\c "Figure" ', leader: 'dot' },
  index: { instr: ' INDEX ', leader: 'dot' },
  authorities: { instr: ' TOA \\h \\p ', leader: 'dot' },
}

const has = (instr: FieldInstruction, name: string): boolean =>
  fieldSwitchArg(instr, name) !== undefined

/** "2-4" as [2, 4]; a single number as both ends */
function rangeOf(arg: string | undefined): [number, number] | null {
  const m = /^\s*(\d+)\s*(?:-\s*(\d+))?\s*$/.exec(arg ?? '')
  if (!m) return null
  const from = Number(m[1])
  return [from, m[2] === undefined ? from : Number(m[2])]
}

/** How a TOC instruction's switches shape the entries. */
export interface TocPlan extends TocSources {
  /** the built-in heading styles are a source by themselves (\o, or no source named at all) */
  headings: boolean
  from: number
  to: number
  /** only the TC fields with this identifier (\f F); undefined for those without one */
  entryId?: string
  /** the TC field levels listed (\l); every level when absent */
  entryLevels: [number, number] | null
  /** levels that show no page number (\n, or \n with a range) */
  noPage: (level: number) => boolean
  /** set when the page follows the title instead of sitting behind the right tab (\p) */
  pageSep?: string
  hyperlinks: boolean
}

/** `\t "Style A,1,Style B,2"` as rows; Word writes the list separator of the system, a comma or a semicolon */
export function tocStyleLevels(arg: string | undefined): TocStyleLevel[] {
  const parts = (arg ?? '').split(arg?.includes(';') && !arg.includes(',') ? ';' : ',')
  const out: TocStyleLevel[] = []
  for (let i = 0; i < parts.length; i++) {
    const name = parts[i].trim()
    if (!name) continue
    const level = /^\s*[1-9]\s*$/.test(parts[i + 1] ?? '') ? Number(parts[++i]) : 1
    out.push({ name, level })
  }
  return out
}

const styleLevelsArg = (styles: TocStyleLevel[]): string =>
  styles.map((s) => `${s.name.replace(/[",;]/g, '')},${s.level}`).join(',')

export function tocPlan(instrText: string): TocPlan {
  const instr = parseFieldInstruction(instrText)
  const styles = tocStyleLevels(fieldSwitchArg(instr, 't'))
  const entryId = fieldSwitchArg(instr, 'f')
  const [from, to] = rangeOf(fieldSwitchArg(instr, 'o')) ?? [1, 9]
  const n = fieldSwitchArg(instr, 'n')
  const hidden = n === undefined ? null : (rangeOf(n) ?? [1, 9])
  const p = fieldSwitchArg(instr, 'p')
  return {
    headings: has(instr, 'o') || (styles.length === 0 && entryId === undefined && !has(instr, 'u')),
    styles,
    outline: has(instr, 'u'),
    entryFields: entryId !== undefined,
    ...(entryId ? { entryId } : {}),
    entryLevels: rangeOf(fieldSwitchArg(instr, 'l')),
    from,
    to,
    noPage: (level) => hidden !== null && level >= hidden[0] && level <= hidden[1],
    ...(p !== undefined ? { pageSep: p } : {}),
    hyperlinks: has(instr, 'h'),
  }
}

export function tocOptionsOf(spec: TableSpec): TocOptions {
  const plan = tocPlan(spec.instr)
  return {
    levels: plan.headings ? Math.min(Math.max(plan.to, 1), 9) : 0,
    styles: plan.styles,
    outline: plan.outline,
    entryFields: plan.entryFields,
    pageNumbers: !plan.noPage(plan.from),
    rightAlign: plan.pageSep === undefined,
    leader: spec.leader,
    hyperlinks: plan.hyperlinks,
  }
}

/** the page-number switches shared by a table of contents and a table of figures */
function withPageSwitches(
  instr: FieldInstruction,
  options: { pageNumbers: boolean; rightAlign: boolean; hyperlinks: boolean },
): FieldInstruction {
  let next = withFieldSwitch(instr, 'n', !options.pageNumbers)
  // a separator already there stays; Word writes a space for a new one
  if (options.rightAlign) next = withFieldSwitch(next, 'p', false)
  else if (!has(next, 'p')) next = withFieldSwitch(next, 'p', ' ')
  return withFieldSwitch(next, 'h', options.hyperlinks)
}

/** the table's instruction with the dialog's options written into it; its other switches stay */
export function tocSpec(options: TocOptions, base: TableSpec = DEFAULT_SPECS.toc): TableSpec {
  let instr = parseFieldInstruction(base.instr)
  const from = Math.max(1, Math.min(tocPlan(base.instr).from, options.levels))
  instr = withFieldSwitch(instr, 'o', options.levels > 0 && `${from}-${options.levels}`)
  instr = withFieldSwitch(instr, 't', options.styles.length > 0 && styleLevelsArg(options.styles))
  instr = withFieldSwitch(instr, 'u', options.outline)
  // an identifier already there stays; without one the field lists the TC fields that have none
  if (!options.entryFields) instr = withFieldSwitch(withFieldSwitch(instr, 'f', false), 'l', false)
  else if (!has(instr, 'f')) instr = withFieldSwitch(instr, 'f', true)
  return {
    instr: formatFieldInstruction(withPageSwitches(instr, options)),
    leader: options.leader,
  }
}

export function figuresOptionsOf(spec: TableSpec): FiguresOptions {
  const instr = parseFieldInstruction(spec.instr)
  const plan = tocPlan(spec.instr)
  return {
    label: fieldSwitchArg(instr, 'c') ?? fieldSwitchArg(instr, 'a') ?? '',
    style: plan.styles[0]?.name ?? '',
    entryId: plan.entryId ?? '',
    includeLabel: !has(instr, 'a'),
    pageNumbers: !has(instr, 'n'),
    rightAlign: plan.pageSep === undefined,
    leader: spec.leader,
    hyperlinks: plan.hyperlinks,
  }
}

export function figuresSpec(
  options: FiguresOptions,
  base: TableSpec = DEFAULT_SPECS.figures,
): TableSpec {
  // the label is quoted inside the instruction: a quote of its own would end it early
  const label = options.label.replace(/"/g, '')
  let instr = parseFieldInstruction(base.instr)
  // \c lists whole captions, \a the same captions without label and number
  const [on, off] = options.includeLabel ? ['c', 'a'] : ['a', 'c']
  if (has(instr, off)) {
    instr = {
      ...instr,
      switches: instr.switches.map((s) => (s.name === off ? { name: on, arg: label } : s)),
    }
  } else instr = withFieldSwitch(instr, on, label)
  instr = withFieldSwitch(
    instr,
    't',
    !!options.style && styleLevelsArg([{ name: options.style, level: 1 }]),
  )
  instr = withFieldSwitch(instr, 'f', /^[A-Za-z]$/.test(options.entryId) && options.entryId)
  return {
    instr: formatFieldInstruction(withPageSwitches(instr, options)),
    leader: options.leader,
  }
}

export function indexOptionsOf(spec: TableSpec): IndexOptions {
  const instr = parseFieldInstruction(spec.instr)
  return {
    runIn: has(instr, 'r'),
    accented: has(instr, 'a'),
    rightAlign: (fieldSwitchArg(instr, 'e') ?? '').includes('\t'),
    leader: spec.leader,
    columns: Math.max(1, Math.round(Number(fieldSwitchArg(instr, 'c'))) || 1),
  }
}

/** the text between an index entry and its pages: \e, else Word's ", " */
export const indexPageSep = (instrText: string): string =>
  fieldSwitchArg(parseFieldInstruction(instrText), 'e') ?? ', '

export function indexSpec(options: IndexOptions, base: TableSpec = DEFAULT_SPECS.index): TableSpec {
  let instr = withFieldSwitch(parseFieldInstruction(base.instr), 'r', options.runIn)
  instr = withFieldSwitch(instr, 'a', options.accented)
  const aligned = (fieldSwitchArg(instr, 'e') ?? '').includes('\t')
  if (options.rightAlign !== aligned)
    instr = withFieldSwitch(instr, 'e', options.rightAlign && '\t')
  instr = withFieldSwitch(instr, 'c', options.columns > 1 && String(options.columns))
  return { instr: formatFieldInstruction(instr), leader: options.leader }
}

export function authoritiesOptionsOf(spec: TableSpec): AuthoritiesOptions {
  const instr = parseFieldInstruction(spec.instr)
  const category = Number(fieldSwitchArg(instr, 'c'))
  return {
    category: category > 0 ? Math.round(category) : 0,
    passim: has(instr, 'p'),
    keepFormatting: !has(instr, 'f'),
    leader: spec.leader,
  }
}

export function authoritiesSpec(
  options: AuthoritiesOptions,
  base: TableSpec = DEFAULT_SPECS.authorities,
): TableSpec {
  let instr = parseFieldInstruction(base.instr)
  instr = withFieldSwitch(instr, 'c', options.category > 0 && String(options.category))
  instr = withFieldSwitch(instr, 'p', options.passim)
  instr = withFieldSwitch(instr, 'f', !options.keepFormatting)
  return { instr: formatFieldInstruction(instr), leader: options.leader }
}

/** a caption's own text: what follows its label and number ("Figure 2: Sales" → "Sales") */
export function captionBody(text: string, label: string): string {
  let rest = text.trim()
  if (label && rest.toLowerCase().startsWith(label.toLowerCase()))
    rest = rest.slice(label.length).trimStart()
  const number =
    /^(?:\d+[-.\u2013:])*(?:\d+|[ivxlcdm]+|[IVXLCDM]+|[a-z]|[A-Z])(?=$|[\s.:\u2013\u2014-])/.exec(
      rest,
    )
  if (!number) return rest
  return rest.slice(number[0].length).replace(/^[\s.:\u2013\u2014-]+/, '')
}

/** the text an INDEX field puts above each letter's entries (\h): "A" stands for the letter; undefined without the switch */
export const indexHeadingPattern = (instrText: string): string | undefined =>
  fieldSwitchArg(parseFieldInstruction(instrText), 'h')

/** the index instruction with letter headings in that pattern, or without any */
export function withIndexHeadings(instrText: string, pattern: string | null): string {
  const instr = parseFieldInstruction(instrText)
  return formatFieldInstruction(withFieldSwitch(instr, 'h', pattern ?? false))
}

/** what a TC field says: its text, its identifier letter (\f) and its level (\l, 1 without) */
export function tocEntryField(tc: string): { text: string; id?: string; level: number } {
  const m = /^\s*(?:"((?:[^"\\]|\\.)*)"|(\S+))([\s\S]*)$/.exec(tc)
  if (!m) return { text: '', level: 1 }
  const rest = parseFieldInstruction(`TC ${m[3]}`)
  const id = fieldSwitchArg(rest, 'f')
  const level = Number(fieldSwitchArg(rest, 'l'))
  return {
    text: (m[1] ?? m[2]).replace(/\\(["\\])/g, '$1'),
    ...(id ? { id } : {}),
    level: level >= 1 && level <= 9 ? Math.round(level) : 1,
  }
}
