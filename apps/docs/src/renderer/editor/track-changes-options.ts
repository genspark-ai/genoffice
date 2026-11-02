/**
 * Word's Track Changes Options → Advanced Options: how markup is drawn on
 * screen. An application preference (like the Balloons mode); the saved
 * docx, exports and print output never see it. The options become one
 * generated stylesheet (screen media only) that overrides the defaults in
 * styles.css through `--rev-*` custom properties and a few mark rules, plus
 * the recorder switches (moves / formatting) and the margin layer's geometry.
 */

export type RevMark =
  'none' | 'color' | 'bold' | 'italic' | 'underline' | 'dunderline' | 'strike' | 'dstrike'

/** Word's 16-color list as it appears in the Advanced dialog (plus its two classic entries) */
export const REV_COLOR_NAMES = [
  'black',
  'blue',
  'turquoise',
  'brightGreen',
  'pink',
  'red',
  'yellow',
  'white',
  'darkBlue',
  'teal',
  'green',
  'violet',
  'darkRed',
  'darkYellow',
  'gray50',
  'gray25',
  'classicRed',
  'classicBlue',
] as const
export type RevColorName = (typeof REV_COLOR_NAMES)[number]

/** the pale fills of Word's table cell highlighting list */
export const CELL_COLOR_NAMES = [
  'lightBlue',
  'palePink',
  'lightYellow',
  'lightOrange',
  ...REV_COLOR_NAMES,
] as const
export type CellColorName = (typeof CELL_COLOR_NAMES)[number]

export type RevColor = 'author' | RevColorName
export type LineColor = 'auto' | RevColorName
export type ChangedLines = 'outside' | 'left' | 'right' | 'none'

export interface TrackChangesOptions {
  insMark: RevMark
  insColor: RevColor
  delMark: RevMark
  delColor: RevColor
  changedLines: ChangedLines
  changedLinesColor: LineColor
  commentColor: RevColor
  trackMoves: boolean
  moveFromMark: RevMark
  moveFromColor: RevColor
  moveToMark: RevMark
  moveToColor: RevColor
  cellIns: CellColorName
  cellDel: CellColorName
  cellMerge: CellColorName
  cellSplit: CellColorName
  trackFormatting: boolean
  fmtMark: RevMark
  fmtColor: RevColor
  /** markup column width, CSS px at 100% zoom */
  balloonWidth: number
  balloonMargin: 'right' | 'left'
  balloonLines: boolean
}

export const DEFAULT_TRACK_CHANGES_OPTIONS: Readonly<TrackChangesOptions> = Object.freeze({
  insMark: 'underline',
  insColor: 'author',
  delMark: 'strike',
  delColor: 'author',
  changedLines: 'outside',
  changedLinesColor: 'auto',
  commentColor: 'author',
  trackMoves: true,
  moveFromMark: 'dstrike',
  moveFromColor: 'green',
  moveToMark: 'dunderline',
  moveToColor: 'green',
  cellIns: 'lightBlue',
  cellDel: 'palePink',
  cellMerge: 'lightYellow',
  cellSplit: 'lightOrange',
  trackFormatting: true,
  fmtMark: 'none',
  fmtColor: 'author',
  balloonWidth: 200,
  balloonMargin: 'right',
  balloonLines: true,
} satisfies TrackChangesOptions)

export const TRACK_CHANGES_OPTIONS_KEY = 'aidocs.trackChangesOptions'
export const BALLOON_WIDTH_MIN = 120
export const BALLOON_WIDTH_MAX = 480

const MARKS = new Set<string>([
  'none',
  'color',
  'bold',
  'italic',
  'underline',
  'dunderline',
  'strike',
  'dstrike',
])
const REV_COLORS = new Set<string>(REV_COLOR_NAMES)
const CELL_COLORS = new Set<string>(CELL_COLOR_NAMES)

const isMark = (v: unknown): v is RevMark => typeof v === 'string' && MARKS.has(v)
const isRevColor = (v: unknown): v is RevColor =>
  v === 'author' || (typeof v === 'string' && REV_COLORS.has(v))
const isLineColor = (v: unknown): v is LineColor =>
  v === 'auto' || (typeof v === 'string' && REV_COLORS.has(v))
const isCellColor = (v: unknown): v is CellColorName => typeof v === 'string' && CELL_COLORS.has(v)

export function clampBalloonWidth(n: unknown): number {
  const v =
    typeof n === 'number' && Number.isFinite(n)
      ? Math.round(n)
      : DEFAULT_TRACK_CHANGES_OPTIONS.balloonWidth
  return Math.max(BALLOON_WIDTH_MIN, Math.min(BALLOON_WIDTH_MAX, v))
}

/** every field validated; unknown or missing ones fall back to the default */
export function normalizeTrackChangesOptions(raw: unknown): TrackChangesOptions {
  const d = DEFAULT_TRACK_CHANGES_OPTIONS
  const r = (raw && typeof raw === 'object' ? raw : {}) as Partial<
    Record<keyof TrackChangesOptions, unknown>
  >
  const mark = (k: keyof TrackChangesOptions) =>
    isMark(r[k]) ? (r[k] as RevMark) : (d[k] as RevMark)
  const color = (k: keyof TrackChangesOptions) =>
    isRevColor(r[k]) ? (r[k] as RevColor) : (d[k] as RevColor)
  const cell = (k: keyof TrackChangesOptions) =>
    isCellColor(r[k]) ? (r[k] as CellColorName) : (d[k] as CellColorName)
  const bool = (k: keyof TrackChangesOptions) =>
    typeof r[k] === 'boolean' ? (r[k] as boolean) : (d[k] as boolean)
  return {
    insMark: mark('insMark'),
    insColor: color('insColor'),
    delMark: mark('delMark'),
    delColor: color('delColor'),
    changedLines: ['outside', 'left', 'right', 'none'].includes(r.changedLines as string)
      ? (r.changedLines as ChangedLines)
      : d.changedLines,
    changedLinesColor: isLineColor(r.changedLinesColor) ? r.changedLinesColor : d.changedLinesColor,
    commentColor: color('commentColor'),
    trackMoves: bool('trackMoves'),
    moveFromMark: mark('moveFromMark'),
    moveFromColor: color('moveFromColor'),
    moveToMark: mark('moveToMark'),
    moveToColor: color('moveToColor'),
    cellIns: cell('cellIns'),
    cellDel: cell('cellDel'),
    cellMerge: cell('cellMerge'),
    cellSplit: cell('cellSplit'),
    trackFormatting: bool('trackFormatting'),
    fmtMark: mark('fmtMark'),
    fmtColor: color('fmtColor'),
    balloonWidth: clampBalloonWidth(r.balloonWidth),
    balloonMargin: r.balloonMargin === 'left' ? 'left' : 'right',
    balloonLines: bool('balloonLines'),
  }
}

export function loadTrackChangesOptions(): TrackChangesOptions {
  if (typeof localStorage === 'undefined') return { ...DEFAULT_TRACK_CHANGES_OPTIONS }
  try {
    const raw = localStorage.getItem(TRACK_CHANGES_OPTIONS_KEY)
    return normalizeTrackChangesOptions(raw ? JSON.parse(raw) : null)
  } catch {
    return { ...DEFAULT_TRACK_CHANGES_OPTIONS }
  }
}

export function saveTrackChangesOptions(o: TrackChangesOptions): void {
  if (typeof localStorage === 'undefined') return
  localStorage.setItem(TRACK_CHANGES_OPTIONS_KEY, JSON.stringify(o))
}

/** the options the margin layer / preview read while building (App syncs it) */
export const trackChangesOptionsState: { current: TrackChangesOptions } = {
  current: { ...DEFAULT_TRACK_CHANGES_OPTIONS },
}

/** the paper token carrying a named color (light + dark-page values live in styles.css) */
export const revColorVar = (name: RevColorName | CellColorName) =>
  `var(--docs-paper-rev-color-${name})`

/** the ink a fixed choice resolves to; by-author keeps the per-author cycle (null = emit nothing) */
function fixedInk(c: RevColor | LineColor): string | null {
  if (c === 'author') return null
  if (c === 'auto') return 'var(--docs-paper-ink)'
  return revColorVar(c)
}

const RESET_MARK = 'text-decoration:none;font-weight:inherit;font-style:inherit'

function markDecl(mark: RevMark, ink: string): string {
  switch (mark) {
    case 'none':
      return `color:inherit;${RESET_MARK}`
    case 'color':
      return RESET_MARK
    case 'bold':
      return `${RESET_MARK};font-weight:700`
    case 'italic':
      return `${RESET_MARK};font-style:italic`
    case 'underline':
      return `text-decoration:underline;text-decoration-color:${ink};text-decoration-style:solid`
    case 'dunderline':
      return `text-decoration:underline;text-decoration-color:${ink};text-decoration-style:double`
    case 'strike':
      return `text-decoration:line-through;text-decoration-color:${ink};text-decoration-style:solid`
    case 'dstrike':
      return `text-decoration:line-through;text-decoration-color:${ink};text-decoration-style:double`
  }
}

/** Word's named defaults render as before: the cell fill tokens in styles.css already carry them */
const CELL_DEFAULT: Record<'cellIns' | 'cellDel' | 'cellMerge', CellColorName> = {
  cellIns: DEFAULT_TRACK_CHANGES_OPTIONS.cellIns,
  cellDel: DEFAULT_TRACK_CHANGES_OPTIONS.cellDel,
  cellMerge: DEFAULT_TRACK_CHANGES_OPTIONS.cellMerge,
}

/**
 * Screen-only stylesheet for the non-default choices. Every rule is scoped
 * under .app and follows styles.css in the cascade, so it wins ties; the
 * final-text views (No Markup / Simple / Original / hidden filters) keep
 * their own resets, which is why the mark rules also reset weight and style.
 */
export function trackChangesOptionsCss(o: TrackChangesOptions): string {
  const d = DEFAULT_TRACK_CHANGES_OPTIONS
  const vars: string[] = []
  const rules: string[] = []
  const inkVar = (name: string, c: RevColor | LineColor) => {
    const ink = fixedInk(c)
    if (ink) vars.push(`${name}:${ink}`)
  }
  inkVar('--rev-ins-ink', o.insColor)
  inkVar('--rev-del-ink', o.delColor)
  inkVar('--rev-fmt-ink', o.fmtColor)
  if (o.changedLinesColor !== 'auto') inkVar('--rev-bar-ink', o.changedLinesColor)
  if (o.commentColor !== 'author') {
    const ink = revColorVar(o.commentColor)
    vars.push(
      `--rev-comment-ink:${ink}`,
      `--rev-comment-bg:color-mix(in srgb, ${ink} 18%, transparent)`,
    )
  }
  for (const k of ['cellIns', 'cellDel', 'cellMerge'] as const) {
    if (o[k] !== CELL_DEFAULT[k])
      vars.push(`--rev-${k.slice(4).toLowerCase()}-bg:${revColorVar(o[k])}`)
  }
  // :where() keeps the mark rules at one class of specificity: later than styles.css
  // (so they beat the defaults) but still under the final-text views' resets
  const insSel = ':where(.app) .doc-ins:where(:not(.doc-move-to))'
  const delSel = ':where(.app) .doc-del:where(:not(.doc-move-from))'
  const insInk = 'var(--rev-ins-ink, var(--docs-paper-rev-ink))'
  const delInk = 'var(--rev-del-ink, var(--docs-paper-rev-ink))'
  if (o.insMark !== d.insMark) rules.push(`${insSel}{${markDecl(o.insMark, insInk)}}`)
  if (o.delMark !== d.delMark) rules.push(`${delSel}{${markDecl(o.delMark, delInk)}}`)
  if (o.fmtMark !== d.fmtMark) {
    const ink = 'var(--rev-fmt-ink, var(--docs-paper-rev-ink))'
    rules.push(
      `:where(.app) .has-rpr-change{${o.fmtMark === 'none' ? RESET_MARK : `color:${ink};${markDecl(o.fmtMark, ink)}`}}`,
    )
  }
  if (!o.trackMoves) {
    // moves read as a plain deletion + insertion (the author stylesheet colors them like one)
    rules.push(
      '.doc-move-from,.doc-move-to{--docs-paper-rev-ink:inherit;text-decoration-style:solid}',
      '.app .doc-del-marker.doc-move-from{border-left-style:solid;border-left-width:2px}',
    )
  } else {
    const move = (cls: string, mark: RevMark, color: RevColor, defMark: RevMark) => {
      // by-author: drop the fixed move green so the author stylesheet (or the root ink) applies
      if (color === 'author') rules.push(`.${cls}{--docs-paper-rev-ink:inherit}`)
      else if (color !== d.moveFromColor)
        rules.push(`.app .${cls}{--docs-paper-rev-ink:${revColorVar(color)}}`)
      if (mark !== defMark)
        rules.push(`:where(.app) .${cls}{${markDecl(mark, 'var(--docs-paper-rev-ink)')}}`)
    }
    move('doc-move-from', o.moveFromMark, o.moveFromColor, d.moveFromMark)
    move('doc-move-to', o.moveToMark, o.moveToColor, d.moveToMark)
  }
  if (o.changedLines === 'none') rules.push('.app .change-bar{display:none}')
  const out: string[] = []
  if (vars.length) out.push(`.app{${vars.join(';')}}`)
  out.push(...rules)
  return out.length ? `@media screen{${out.join('\n')}}` : ''
}

/** which move ends take the author's ink instead of Word's fixed move green */
export const movesByAuthor = (o: TrackChangesOptions): { from: boolean; to: boolean } => ({
  from: !o.trackMoves || o.moveFromColor === 'author',
  to: !o.trackMoves || o.moveToColor === 'author',
})
