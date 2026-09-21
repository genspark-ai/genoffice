/**
 * The author's markdown conventions, read off the source and applied when a
 * changed block is re-serialized, so an edit inside a `*` list or a `_em_`
 * document does not come back with `-` and `*`.
 */

export interface MarkdownStyle {
  bullet: '-' | '*' | '+'
  orderedDelimiter: '.' | ')'
  /** every item carries the first item's number (`1. 1. 1.`) instead of counting up */
  orderedRepeat: boolean
  em: '*' | '_'
  strong: '**' | '__'
  /** level 1 and 2 headings underlined with `===` / `---` */
  setext: boolean
  fence: '`' | '~'
  rule: string
  hardBreak: 'spaces' | 'backslash' | 'html'
  /** cells padded so the pipes line up across rows */
  tableAligned: boolean
}

export type StyleHints = Partial<MarkdownStyle>

export const DEFAULT_MARKDOWN_STYLE: MarkdownStyle = {
  bullet: '-',
  orderedDelimiter: '.',
  orderedRepeat: false,
  em: '*',
  strong: '**',
  setext: false,
  fence: '`',
  rule: '---',
  hardBreak: 'spaces',
  tableAligned: true,
}

/** what the renderers consult; the serializer's defaults outside a scope */
let active: MarkdownStyle = DEFAULT_MARKDOWN_STYLE

export function activeMarkdownStyle(): MarkdownStyle {
  return active
}

/** run `fn` with `style` in force for every renderer it triggers */
export function withMarkdownStyle<T>(style: MarkdownStyle, fn: () => T): T {
  const previous = active
  active = style
  try {
    return fn()
  } finally {
    active = previous
  }
}

/** first hint that has an answer wins; the defaults fill the rest */
export function resolveMarkdownStyle(...hints: StyleHints[]): MarkdownStyle {
  const style: Record<string, unknown> = { ...DEFAULT_MARKDOWN_STYLE }
  for (const key of Object.keys(style) as (keyof MarkdownStyle)[]) {
    const hint = hints.find((h) => h[key] !== undefined)
    if (hint) style[key] = hint[key]
  }
  return style as unknown as MarkdownStyle
}

/** a marked token; only the fields the style walk reads are typed */
export interface MarkedToken {
  type: string
  raw: string
  depth?: number
  ordered?: boolean
  tokens?: MarkedToken[]
  nestedTokens?: MarkedToken[]
  items?: MarkedToken[]
  header?: Array<{ tokens?: MarkedToken[] }>
  rows?: Array<Array<{ tokens?: MarkedToken[] }>>
}

type Tally = { [K in keyof MarkdownStyle]?: Map<MarkdownStyle[K], number> }

function vote<K extends keyof MarkdownStyle>(tally: Tally, key: K, value: MarkdownStyle[K]): void {
  const counts = (tally[key] ??= new Map()) as Map<MarkdownStyle[K], number>
  counts.set(value, (counts.get(value) ?? 0) + 1)
}

function walk(tokens: MarkedToken[], tally: Tally): void {
  for (const token of tokens) {
    switch (token.type) {
      case 'list': {
        const items = token.items ?? []
        const first = items[0]?.raw ?? ''
        if (token.ordered) {
          const delimiter = /^\s*\d+([.)])/.exec(first)?.[1]
          if (delimiter === '.' || delimiter === ')') vote(tally, 'orderedDelimiter', delimiter)
          if (items.length > 1) {
            const numbers = items.map((item) => /^\s*(\d+)/.exec(item.raw)?.[1])
            if (numbers.every((n) => n !== undefined)) {
              vote(
                tally,
                'orderedRepeat',
                numbers.every((n) => n === numbers[0]),
              )
            }
          }
        } else {
          const bullet = /^\s*([-*+])/.exec(first)?.[1]
          if (bullet === '-' || bullet === '*' || bullet === '+') vote(tally, 'bullet', bullet)
        }
        for (const item of items) walk(item.tokens ?? [], tally)
        continue
      }
      case 'taskList': {
        // the task list tokenizer leaves item raws empty; the marker is at the start of the list
        const bullet = /^\s*([-*+])/.exec(token.raw)?.[1]
        if (bullet === '-' || bullet === '*' || bullet === '+') vote(tally, 'bullet', bullet)
        for (const item of token.items ?? []) {
          walk(item.tokens ?? [], tally)
          walk(item.nestedTokens ?? [], tally)
        }
        continue
      }
      case 'heading':
        // only levels 1 and 2 can be underlined, so only they express a choice
        if ((token.depth ?? 1) <= 2) vote(tally, 'setext', !/^ {0,3}#/.test(token.raw))
        break
      case 'code': {
        const fence = /^ {0,3}([`~])\1{2,}/.exec(token.raw)?.[1]
        if (fence === '`' || fence === '~') vote(tally, 'fence', fence)
        break
      }
      case 'hr':
        vote(tally, 'rule', token.raw.trim())
        break
      case 'em':
        vote(tally, 'em', token.raw.startsWith('_') ? '_' : '*')
        break
      case 'strong':
        vote(tally, 'strong', token.raw.startsWith('__') ? '__' : '**')
        break
      case 'br':
        vote(tally, 'hardBreak', token.raw.startsWith('\\') ? 'backslash' : 'spaces')
        break
      case 'html':
        if (/^<br\s*\/?>$/i.test(token.raw)) vote(tally, 'hardBreak', 'html')
        break
      case 'table': {
        const lines = token.raw.trimEnd().split('\n')
        vote(
          tally,
          'tableAligned',
          lines.every((line) => line.trimEnd().length === lines[0]!.trimEnd().length),
        )
        for (const cell of token.header ?? []) walk(cell.tokens ?? [], tally)
        for (const row of token.rows ?? []) for (const cell of row) walk(cell.tokens ?? [], tally)
        continue
      }
      default:
        break
    }
    if (token.tokens) walk(token.tokens, tally)
  }
}

/** the majority convention for each property the tokens exhibit */
export function styleHintsFromTokens(tokens: MarkedToken[]): StyleHints {
  const tally: Tally = {}
  walk(tokens, tally)
  const hints: StyleHints = {}
  for (const key of Object.keys(tally) as (keyof MarkdownStyle)[]) {
    let best: unknown
    let bestCount = 0
    for (const [value, count] of tally[key]!) {
      if (count > bestCount) {
        best = value
        bestCount = count
      }
    }
    ;(hints as Record<string, unknown>)[key] = best
  }
  return hints
}
