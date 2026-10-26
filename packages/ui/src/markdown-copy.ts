import { parseBlocks, type MdBlock } from './Markdown'

/**
 * Copy-side companions of the Markdown chat renderer: walk the same block
 * tree so what lands on the clipboard matches what the bubble showed —
 * tables become tab-separated text (or a real <table> in the HTML flavor),
 * fences and emphasis markers never survive into plain text.
 */

// the inline grammar mirrors Markdown.tsx's INLINE_RE: `code`, **bold**,
// *italic*, ![img](href), [text](href). Kept parallel deliberately: the
// renderer tolerates partial input and so do we.
const HREF = /(?:[^\s()]|\([^\s()]*\))+/.source
const INLINE_RE = new RegExp(
  `(\`[^\`\\n]+\`|\\*\\*[^*\\n]+?\\*\\*|\\*[^*\\n]+?\\*|!\\[[^\\]\\n]*\\]\\(${HREF}\\)|\\[[^\\]\\n]+\\]\\(${HREF}\\))`,
  'g',
)
const LINK_RE = new RegExp(`^\\[([^\\]]+)\\]\\((${HREF})\\)$`)
const IMG_RE = new RegExp(`^!\\[([^\\]\\n]*)\\]\\((${HREF})\\)$`)

/** inline markdown → visible text: markers dropped, links keep their label, images keep their alt */
export function inlinePlainText(text: string): string {
  let out = ''
  let last = 0
  for (const m of text.matchAll(INLINE_RE)) {
    const i = m.index ?? 0
    if (i > last) out += text.slice(last, i)
    const tok = m[0] ?? ''
    if (tok.startsWith('`')) out += tok.slice(1, -1)
    else if (tok.startsWith('**')) out += inlinePlainText(tok.slice(2, -2))
    else if (tok.startsWith('![')) {
      const img = IMG_RE.exec(tok)
      out += img?.[1] ?? ''
    } else if (tok.startsWith('[')) {
      const link = LINK_RE.exec(tok)
      out += link?.[1] ?? tok
    } else out += inlinePlainText(tok.slice(1, -1))
    last = i + tok.length
  }
  if (last < text.length) out += text.slice(last)
  return out
}

function escapeHtml(s: string): string {
  return s
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
}

/** inline markdown → HTML fragment */
function inlineHtml(text: string): string {
  let out = ''
  let last = 0
  for (const m of text.matchAll(INLINE_RE)) {
    const i = m.index ?? 0
    if (i > last) out += escapeHtml(text.slice(last, i))
    const tok = m[0] ?? ''
    if (tok.startsWith('`')) out += `<code>${escapeHtml(tok.slice(1, -1))}</code>`
    else if (tok.startsWith('**')) out += `<strong>${inlineHtml(tok.slice(2, -2))}</strong>`
    else if (tok.startsWith('![')) {
      const img = IMG_RE.exec(tok)
      out += escapeHtml(img?.[1] ?? '')
    } else if (tok.startsWith('[')) {
      const link = LINK_RE.exec(tok)
      if (link) out += `<a href="${escapeHtml(link[2] ?? '')}">${escapeHtml(link[1] ?? '')}</a>`
      else out += escapeHtml(tok)
    } else out += `<em>${inlineHtml(tok.slice(1, -1))}</em>`
    last = i + tok.length
  }
  if (last < text.length) out += escapeHtml(text.slice(last))
  return out
}

function blockPlainText(b: MdBlock): string {
  switch (b.kind) {
    case 'p':
    case 'h':
      return b.kind === 'p'
        ? b.lines.map(inlinePlainText).join('\n')
        : inlinePlainText(b.text)
    case 'ul':
      return b.items.map((it) => `- ${inlinePlainText(it)}`).join('\n')
    case 'ol':
      return b.items.map((it, i) => `${i + 1}. ${inlinePlainText(it)}`).join('\n')
    case 'code':
      return b.lines.join('\n')
    case 'table': {
      const row = (cells: string[]): string => cells.map(inlinePlainText).join('\t')
      return [row(b.head), ...b.rows.map(row)].join('\n')
    }
    case 'img':
      return b.alt
  }
}

function blockHtml(b: MdBlock): string {
  switch (b.kind) {
    case 'p':
      return `<p>${b.lines.map(inlineHtml).join('<br>')}</p>`
    case 'h':
      return `<h${b.level}>${inlineHtml(b.text)}</h${b.level}>`
    case 'ul':
      return `<ul>${b.items.map((it) => `<li>${inlineHtml(it)}</li>`).join('')}</ul>`
    case 'ol':
      return `<ol>${b.items.map((it) => `<li>${inlineHtml(it)}</li>`).join('')}</ol>`
    case 'code':
      return `<pre><code>${escapeHtml(b.lines.join('\n'))}</code></pre>`
    case 'table': {
      const th = b.head.map((c) => `<th>${inlineHtml(c)}</th>`).join('')
      const rows = b.rows
        .map((r) => `<tr>${r.map((c) => `<td>${inlineHtml(c)}</td>`).join('')}</tr>`)
        .join('')
      return `<table><thead><tr>${th}</tr></thead><tbody>${rows}</tbody></table>`
    }
    case 'img':
      return b.alt ? `<p>${escapeHtml(b.alt)}</p>` : ''
  }
}

/** markdown chat reply → clean plain text (no fences, no ** markers, tables as TSV) */
export function markdownToPlainText(text: string): string {
  return parseBlocks(text)
    .map(blockPlainText)
    .filter((s) => s !== '')
    .join('\n\n')
}

/** markdown chat reply → standalone HTML fragment for the text/html clipboard flavor */
export function markdownToHtml(text: string): string {
  return parseBlocks(text)
    .map(blockHtml)
    .filter((s) => s !== '')
    .join('\n')
}

/**
 * Copy a chat reply with both flavors (text/html + text/plain) so pasting into
 * a rich editor keeps tables/bold while pasting into a plain field stays clean.
 * Falls back to plain-text-only when the async clipboard or ClipboardItem is
 * unavailable (older contexts, denied permission).
 */
export async function copyMarkdownRich(text: string): Promise<void> {
  const plain = markdownToPlainText(text)
  const html = markdownToHtml(text)
  try {
    const item = new ClipboardItem({
      'text/html': new Blob([html], { type: 'text/html' }),
      'text/plain': new Blob([plain], { type: 'text/plain' }),
    })
    await navigator.clipboard.write([item])
  } catch {
    await navigator.clipboard.writeText(plain)
  }
}
