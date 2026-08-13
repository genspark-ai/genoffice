import type { Extension, JSONContent } from '@tiptap/core'
import MarkdownIt from 'markdown-it'
import yaml from 'js-yaml'

const md = new MarkdownIt({
  html: false,
  linkify: true,
  typographer: true,
})

// ── Frontmatter ──────────────────────────────────────────────────────────────

const FRONTMATTER_RE = /^---\r?\n([\s\S]*?)\r?\n---\r?\n?/

export interface ParsedMarkdown {
  frontmatter: Record<string, unknown> | null
  body: string
}

export function parseFrontmatter(raw: string): ParsedMarkdown {
  const match = raw.match(FRONTMATTER_RE)
  if (!match) return { frontmatter: null, body: raw }
  try {
    const data = yaml.load(match[1]) as Record<string, unknown>
    return {
      frontmatter: data && typeof data === 'object' ? data : null,
      body: raw.slice(match[0].length),
    }
  } catch {
    return { frontmatter: null, body: raw }
  }
}

export function serializeFrontmatter(data: Record<string, unknown> | null): string {
  if (!data || Object.keys(data).length === 0) return ''
  return `---\n${yaml.dump(data, { lineWidth: 80 }).trimEnd()}\n---\n\n`
}

// ── Parse .md → ProseMirror JSON ─────────────────────────────────────────────

export function markdownToProseMirror(raw: string, extensions: Extension[]): JSONContent {
  const { frontmatter, body } = parseFrontmatter(raw)
  const html = md.render(body)

  // Use TipTap's generateJSON if available at runtime, otherwise manual parse
  // We import dynamically to avoid circular deps in tests
  let json: JSONContent
  try {
    // eslint-disable-next-line @typescript-eslint/no-require-imports
    const { generateJSON } = require('@tiptap/core')
    json = generateJSON(html, extensions)
  } catch {
    json = htmlToSimpleJson(html)
  }

  // Inject frontmatter node at the top if present
  if (frontmatter && Object.keys(frontmatter).length > 0) {
    const fmNode: JSONContent = {
      type: 'frontmatter',
      attrs: { data: frontmatter },
    }
    if (json.content) {
      json.content = [fmNode, ...json.content]
    } else {
      json.content = [fmNode]
    }
  }

  return json
}

// ── Serialize ProseMirror JSON → .md ─────────────────────────────────────────

export function prosemirrorToMarkdown(json: JSONContent): string {
  const parts: string[] = []
  const content = json.content ?? []

  for (const node of content) {
    if (node.type === 'frontmatter') {
      parts.push(serializeFrontmatter(node.attrs?.data as Record<string, unknown> | null))
      continue
    }
    parts.push(serializeNode(node))
  }

  return parts.join('')
}

function serializeNode(node: JSONContent): string {
  const text = serializeInline(node.content ?? [])
  switch (node.type) {
    case 'heading': {
      const level = (node.attrs?.level as number) ?? 1
      return `${'#'.repeat(level)} ${text}\n\n`
    }
    case 'paragraph':
      return `${text}\n\n`
    case 'bulletList':
      return (node.content ?? []).map((li) => serializeListItem(li, '- ')).join('') + '\n'
    case 'orderedList':
      return (
        (node.content ?? []).map((li, i) => serializeListItem(li, `${i + 1}. `)).join('') + '\n'
      )
    case 'listItem':
      return serializeListItem(node, '- ')
    case 'taskList':
      return (node.content ?? []).map((li) => serializeTaskItem(li)).join('') + '\n'
    case 'taskItem': {
      const checked = node.attrs?.checked as boolean
      const checkbox = checked ? '[x] ' : '[ ] '
      return `- ${checkbox}${serializeInline(node.content ?? [])}\n`
    }
    case 'blockquote':
      return (
        (node.content ?? [])
          .map((n) => {
            const s = serializeNode(n)
            return s.replace(/\n/g, '\n> ').replace(/^> /, '> ')
          })
          .join('') + '\n'
      )
    case 'codeBlock': {
      const lang = (node.attrs?.language as string) ?? ''
      const code = serializeInline(node.content ?? [])
      return `\`\`\`${lang}\n${code}\n\`\`\`\n\n`
    }
    case 'horizontalRule':
      return '---\n\n'
    case 'table':
      return serializeTable(node)
    case 'frontmatter':
      return serializeFrontmatter(node.attrs?.data as Record<string, unknown> | null)
    default:
      return text + '\n\n'
  }
}

function serializeListItem(node: JSONContent, prefix: string): string {
  const inner = (node.content ?? []).map((n) => serializeNode(n).trimEnd()).join('\n')
  const lines = inner.split('\n')
  return lines.map((line, i) => `${i === 0 ? prefix : '  '}${line}`).join('\n') + '\n'
}

function serializeTaskItem(node: JSONContent): string {
  const checked = node.attrs?.checked as boolean
  const checkbox = checked ? '[x] ' : '[ ] '
  const inner = (node.content ?? []).map((n) => serializeNode(n).trimEnd()).join(' ')
  return `- ${checkbox}${inner}\n`
}

function serializeTable(node: JSONContent): string {
  const rows = node.content ?? []
  if (rows.length === 0) return ''

  const tableData: string[][] = []
  for (const row of rows) {
    const cells = row.content ?? []
    const cellTexts: string[] = []
    for (const cell of cells) {
      cellTexts.push(serializeInline(cell.content ?? []).replace(/\|/g, '\\|'))
    }
    tableData.push(cellTexts)
  }

  if (tableData.length === 0) return ''

  const lines: string[] = []
  const header = tableData[0]
  lines.push(`| ${header.join(' | ')} |`)
  lines.push(`| ${header.map(() => '---').join(' | ')} |`)
  for (let i = 1; i < tableData.length; i++) {
    lines.push(`| ${tableData[i].join(' | ')} |`)
  }
  return lines.join('\n') + '\n\n'
}

function serializeInline(content: JSONContent[]): string {
  const parts: string[] = []
  for (const node of content) {
    if (node.type === 'text') {
      let text = (node.text as string) ?? ''
      const marks = node.marks ?? []
      // Sort marks for consistent output: bold → italic → strike → code → highlight → link
      const MARK_ORDER = {
        bold: 0,
        strong: 0,
        italic: 1,
        em: 1,
        strike: 2,
        code: 3,
        highlight: 4,
        link: 5,
      }
      const sorted = [...marks].sort(
        (a, b) =>
          (MARK_ORDER[a.type as keyof typeof MARK_ORDER] ?? 99) -
          (MARK_ORDER[b.type as keyof typeof MARK_ORDER] ?? 99),
      )
      // Apply inline marks (bold, italic, strike, code, highlight) first
      for (const mark of sorted) {
        switch (mark.type) {
          case 'bold':
          case 'strong':
            text = `**${text}**`
            break
          case 'italic':
          case 'em':
            text = `*${text}*`
            break
          case 'strike':
            text = `~~${text}~~`
            break
          case 'code':
            text = `\`${text}\``
            break
          case 'highlight':
            text = `==${text}==`
            break
        }
      }
      // Apply link last (wraps differently)
      const linkMark = sorted.find((m) => m.type === 'link')
      if (linkMark) {
        text = `[${text}](${linkMark.attrs?.href ?? ''})`
      }
      parts.push(text)
    } else if (node.type === 'hardBreak') {
      parts.push('  \n')
    } else if (node.type === 'image') {
      const alt = (node.attrs?.alt as string) ?? ''
      const src = (node.attrs?.src as string) ?? ''
      parts.push(`![${alt}](${src})`)
    } else if (node.type === 'mention') {
      parts.push((node.attrs?.label as string) ?? '')
    } else {
      // Recurse into inline nodes
      parts.push(serializeInline(node.content ?? []))
    }
  }
  return parts.join('')
}

// ── Fallback HTML → simple JSON (when generateJSON unavailable) ──────────────

function htmlToSimpleJson(html: string): JSONContent {
  const content: JSONContent[] = []
  const div = typeof document !== 'undefined' ? document.createElement('div') : null
  if (!div)
    return {
      type: 'doc',
      content: [{ type: 'paragraph', content: [{ type: 'text', text: html }] }],
    }
  div.innerHTML = html

  for (const child of Array.from(div.children)) {
    const node = domElementToNode(child)
    if (node) content.push(node)
  }

  return { type: 'doc', content }
}

function domElementToNode(el: Element): JSONContent | null {
  const tag = el.tagName.toLowerCase()
  const childContent = () => {
    const items: JSONContent[] = []
    for (const c of Array.from(el.childNodes)) {
      if (c.nodeType === 3) {
        const text = c.textContent ?? ''
        if (text) items.push({ type: 'text', text })
      } else if (c.nodeType === 1) {
        const n = domElementToNode(c as Element)
        if (n) items.push(n)
      }
    }
    return items
  }

  switch (tag) {
    case 'h1':
    case 'h2':
    case 'h3':
    case 'h4':
    case 'h5':
    case 'h6':
      return { type: 'heading', attrs: { level: parseInt(tag[1]) }, content: childContent() }
    case 'p':
      return { type: 'paragraph', content: childContent() }
    case 'strong':
    case 'b':
      return { type: 'text', text: el.textContent ?? '', marks: [{ type: 'bold' }] }
    case 'em':
    case 'i':
      return { type: 'text', text: el.textContent ?? '', marks: [{ type: 'italic' }] }
    case 'code':
      return { type: 'text', text: el.textContent ?? '', marks: [{ type: 'code' }] }
    case 'a':
      return {
        type: 'text',
        text: el.textContent ?? '',
        marks: [{ type: 'link', attrs: { href: el.getAttribute('href') ?? '' } }],
      }
    case 'ul':
      return {
        type: 'bulletList',
        content: Array.from(el.children)
          .map((li) => domElementToNode(li))
          .filter(Boolean) as JSONContent[],
      }
    case 'ol':
      return {
        type: 'orderedList',
        content: Array.from(el.children)
          .map((li) => domElementToNode(li))
          .filter(Boolean) as JSONContent[],
      }
    case 'li':
      return { type: 'listItem', content: childContent() }
    case 'blockquote':
      return { type: 'blockquote', content: childContent() }
    case 'pre':
      return { type: 'codeBlock', content: [{ type: 'text', text: el.textContent ?? '' }] }
    case 'hr':
      return { type: 'horizontalRule' }
    case 'img':
      return {
        type: 'image',
        attrs: { src: el.getAttribute('src') ?? '', alt: el.getAttribute('alt') ?? '' },
      }
    default:
      return { type: 'paragraph', content: childContent() }
  }
}
