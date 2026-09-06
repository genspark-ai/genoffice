import { Fragment, type ReactNode } from 'react'

/**
 * Minimal dependency-free markdown for chat bubbles: paragraphs, ul/ol,
 * headings, pipe tables, fenced code, **bold**, *italic*, `inline code`.
 * Tolerates partial (streaming) input — anything unrecognized renders as
 * plain text.
 *
 * Markdown links stay literal text unless the host passes `nav` and the href
 * carries its scheme — then they become in-app navigation links. External
 * URLs never turn into clickable links here.
 */

export interface MarkdownNav {
  /** href prefix that renders as an in-app navigation link (e.g. 'docnav://') */
  scheme: string
  onNavigate: (href: string) => void
}

// Hrefs may carry one level of balanced parens (sheet names like `Data (2)`
// arrive as sheetnav://Data%20(2)!B2), so the href cannot simply stop at ')'.
const HREF = /(?:[^\s()]|\([^\s()]*\))+/.source
const INLINE_RE = new RegExp(
  `(\`[^\`\\n]+\`|\\*\\*[^*\\n]+?\\*\\*|\\*[^*\\n]+?\\*|\\[[^\\]\\n]+\\]\\(${HREF}\\))`,
  'g',
)
const LINK_RE = new RegExp(`^\\[([^\\]]+)\\]\\((${HREF})\\)$`)

function renderInline(text: string, nav?: MarkdownNav): ReactNode[] {
  const out: ReactNode[] = []
  let last = 0
  let key = 0
  for (const m of text.matchAll(INLINE_RE)) {
    const i = m.index ?? 0
    if (i > last) out.push(text.slice(last, i))
    const tok = m[0] ?? ''
    if (tok.startsWith('`')) out.push(<code key={key++}>{tok.slice(1, -1)}</code>)
    else if (tok.startsWith('**')) out.push(<strong key={key++}>{tok.slice(2, -2)}</strong>)
    else if (tok.startsWith('[')) {
      const link = LINK_RE.exec(tok)
      const href = link?.[2] ?? ''
      if (link && nav && href.startsWith(nav.scheme)) {
        out.push(
          <a
            key={key++}
            className="ai-md-nav"
            href={href}
            onClick={(e) => {
              e.preventDefault()
              nav.onNavigate(href)
            }}
          >
            {link[1]}
          </a>,
        )
      } else {
        out.push(tok) // non-nav links keep today's literal rendering
      }
    } else out.push(<em key={key++}>{tok.slice(1, -1)}</em>)
    last = i + tok.length
  }
  if (last < text.length) out.push(text.slice(last))
  return out
}

type CellAlign = 'left' | 'center' | 'right' | undefined

type MdBlock =
  | { kind: 'p'; lines: string[] }
  | { kind: 'ul'; items: string[] }
  | { kind: 'ol'; items: string[] }
  | { kind: 'h'; text: string }
  | { kind: 'table'; align: CellAlign[]; head: string[]; rows: string[][] }
  | { kind: 'code'; lines: string[] }

const FENCE_RE = /^\s*(`{3,}|~{3,})/
const DELIM_CELL_RE = /^\s*:?-+:?\s*$/

/** Splits a table row on pipes that are neither escaped nor inside inline code. */
function splitCells(line: string): string[] {
  const cells: string[] = []
  let cur = ''
  let inCode = false
  for (let i = 0; i < line.length; i++) {
    const ch = line[i]
    if (ch === '\\' && line[i + 1] === '|') {
      cur += '|'
      i++
    } else if (ch === '`') {
      inCode = !inCode
      cur += ch
    } else if (ch === '|' && !inCode) {
      cells.push(cur)
      cur = ''
    } else cur += ch
  }
  cells.push(cur)
  if (cells.length && !cells[0]?.trim()) cells.shift()
  if (cells.length && !cells[cells.length - 1]?.trim()) cells.pop()
  return cells.map((c) => c.trim())
}

function parseDelimiterRow(line: string): CellAlign[] | null {
  if (!line.includes('-')) return null
  const cells = splitCells(line)
  if (!cells.length || !cells.every((c) => DELIM_CELL_RE.test(c))) return null
  return cells.map((c) => {
    const l = c.startsWith(':')
    const r = c.endsWith(':')
    if (l && r) return 'center'
    if (r) return 'right'
    if (l) return 'left'
    return undefined
  })
}

function parseBlocks(text: string): MdBlock[] {
  const blocks: MdBlock[] = []
  let cur: MdBlock | null = null
  const flush = (): void => {
    if (cur) {
      blocks.push(cur)
      cur = null
    }
  }
  const lines = text.split('\n')
  for (let idx = 0; idx < lines.length; idx++) {
    const raw = lines[idx] ?? ''
    const line = raw.trimEnd()
    if (cur?.kind === 'code') {
      if (FENCE_RE.test(line)) flush()
      else cur.lines.push(raw)
      continue
    }
    if (!line.trim()) {
      flush()
      continue
    }
    if (FENCE_RE.test(line)) {
      flush()
      cur = { kind: 'code', lines: [] }
      continue
    }
    if (cur?.kind === 'table') {
      if (line.includes('|')) {
        const width = cur.head.length
        const cells = splitCells(line).slice(0, width)
        while (cells.length < width) cells.push('')
        cur.rows.push(cells)
        continue
      }
      flush()
    }
    if (line.includes('|')) {
      const align = parseDelimiterRow(lines[idx + 1] ?? '')
      if (align) {
        const head = splitCells(line)
        if (head.length === align.length) {
          flush()
          cur = { kind: 'table', align, head, rows: [] }
          idx++
          continue
        }
      }
    }
    const h = /^#{1,6}\s+(.*)$/.exec(line)
    if (h) {
      flush()
      blocks.push({ kind: 'h', text: h[1] ?? '' })
      continue
    }
    const ul = /^\s*[-*•]\s+(.*)$/.exec(line)
    if (ul) {
      if (cur?.kind !== 'ul') {
        flush()
        cur = { kind: 'ul', items: [] }
      }
      cur.items.push(ul[1] ?? '')
      continue
    }
    const ol = /^\s*\d+[.、)]\s+(.*)$/.exec(line)
    if (ol) {
      if (cur?.kind !== 'ol') {
        flush()
        cur = { kind: 'ol', items: [] }
      }
      cur.items.push(ol[1] ?? '')
      continue
    }
    if (cur?.kind !== 'p') {
      flush()
      cur = { kind: 'p', lines: [] }
    }
    cur.lines.push(line)
  }
  flush()
  return blocks
}

export function Markdown({ text, nav }: { text: string; nav?: MarkdownNav }): React.JSX.Element {
  return (
    <div className="ai-md">
      {parseBlocks(text).map((b, i) => {
        if (b.kind === 'h') {
          return (
            <p key={i} className="ai-md-h">
              {renderInline(b.text, nav)}
            </p>
          )
        }
        if (b.kind === 'ul' || b.kind === 'ol') {
          const items = b.items.map((it, j) => <li key={j}>{renderInline(it, nav)}</li>)
          return b.kind === 'ul' ? <ul key={i}>{items}</ul> : <ol key={i}>{items}</ol>
        }
        if (b.kind === 'code') {
          return (
            <pre key={i} className="ai-md-pre">
              <code>{b.lines.join('\n')}</code>
            </pre>
          )
        }
        if (b.kind === 'table') {
          const cellStyle = (j: number): React.CSSProperties | undefined =>
            b.align[j] ? { textAlign: b.align[j] } : undefined
          return (
            <div key={i} className="ai-md-table-wrap">
              <table className="ai-md-table">
                <thead>
                  <tr>
                    {b.head.map((c, j) => (
                      <th key={j} style={cellStyle(j)}>
                        {renderInline(c, nav)}
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {b.rows.map((row, r) => (
                    <tr key={r}>
                      {row.map((c, j) => (
                        <td key={j} style={cellStyle(j)}>
                          {renderInline(c, nav)}
                        </td>
                      ))}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )
        }
        return (
          <p key={i}>
            {b.lines.map((ln, j) => (
              <Fragment key={j}>
                {j > 0 && <br />}
                {renderInline(ln, nav)}
              </Fragment>
            ))}
          </p>
        )
      })}
    </div>
  )
}
