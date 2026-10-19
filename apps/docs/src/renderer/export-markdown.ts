import type { Editor } from '@tiptap/core'
import { isTrackedDeleted } from './ai/doc-utils'
import { serializeRangeToHtml } from './ai/protocol'
import { MARKDOWN_MEDIA_DIR_TOKEN, type MarkdownImage } from '../shared/ipc'

export interface MarkdownExport {
  markdown: string
  images: MarkdownImage[]
}

const IMAGE_EXT: Record<string, string> = {
  'image/png': 'png',
  'image/jpeg': 'jpg',
  'image/gif': 'gif',
  'image/bmp': 'bmp',
  'image/webp': 'webp',
  'image/svg+xml': 'svg',
  'image/tiff': 'tif',
}

function imageExt(src: string): string | null {
  const data = /^data:(image\/[\w.+-]+);base64,/.exec(src)
  if (data) return IMAGE_EXT[data[1]] ?? null
  if (src.startsWith('genoffice-docx-media://')) {
    const ext = /\.(\w+)$/.exec(src)?.[1]?.toLowerCase()
    if (!ext) return null
    return Object.values(IMAGE_EXT).includes(ext === 'jpeg' ? 'jpg' : ext) ? ext : null
  }
  return null
}

/**
 * The whole document as GFM. Rich blocks go through the AI protocol's
 * restricted-HTML serializer (same source the CLI's docx→md uses); protected
 * blocks become a display formula, a field's visible text, a picture, or
 * nothing. Pictures are collected for the writer and referenced by name.
 */
export function documentToMarkdown(editor: Editor): MarkdownExport {
  const top = editor.state.doc
  const parts: string[] = []
  const images: MarkdownImage[] = []
  const named = new Map<string, string>()
  const image = (src: string): string | null => {
    if (!src) return null
    let name = named.get(src)
    if (!name) {
      const ext = imageExt(src)
      if (!ext) return null
      name = `image${images.length + 1}.${ext}`
      named.set(src, name)
      images.push({ name, src })
    }
    return `${MARKDOWN_MEDIA_DIR_TOKEN}/${name}`
  }
  let runStart = -1
  const flush = (end: number) => {
    if (runStart >= 0) parts.push(serializeRangeToHtml(editor, runStart, end, null, image))
    runStart = -1
  }
  top.forEach((node, _offset, index) => {
    if (node.type.name !== 'docProtected') {
      if (runStart < 0) runStart = index
      return
    }
    flush(index - 1)
    if (isTrackedDeleted(node)) return
    const formula = node.attrs.formulaDisplay as { tokens?: string[]; latex?: string } | null
    if (formula?.latex) {
      parts.push(`<formula>${escapeHtml(formula.latex)}</formula>`)
    } else if (formula?.tokens?.length) {
      parts.push(`<p>${escapeHtml(formula.tokens.join(' '))}</p>`)
    } else if (node.attrs.imageDataUrl) {
      const src = image(String(node.attrs.imageDataUrl))
      if (src) parts.push(`<p><img src="${escapeHtml(src)}"></p>`)
    } else {
      const preview = String(node.attrs.previewText ?? '')
        .replace(/\s+/g, ' ')
        .trim()
      if (preview) parts.push(`<p>${escapeHtml(preview)}</p>`)
    }
  })
  flush(top.childCount - 1)
  return { markdown: restrictedHtmlToMarkdown(parts.join('\n')), images }
}

/** Restricted HTML (paragraphs, headings, lists, pre, blockquote, table, formula + inline marks) → GFM. */
export function restrictedHtmlToMarkdown(html: string): string {
  const body = new DOMParser().parseFromString(`<body>${html}</body>`, 'text/html').body
  const blocks: string[] = []
  for (const el of Array.from(body.children)) {
    const md = blockToMarkdown(el)
    if (md) blocks.push(md)
  }
  return blocks.join('\n\n').trimEnd() + '\n'
}

function blockToMarkdown(el: Element): string {
  const tag = el.tagName.toLowerCase()
  const heading = /^h([1-6])$/.exec(tag)
  if (heading) return `${'#'.repeat(Number(heading[1]))} ${inline(el)}`
  switch (tag) {
    case 'p':
      return escapeBlockStart(inline(el))
    case 'ul':
    case 'ol':
      return Array.from(el.children)
        .map(
          (li, i) =>
            `${tag === 'ol' ? `${i + 1}.` : '-'} ${escapeBlockStart(inline(li)).replace(/\n/g, '\n  ')}`,
        )
        .join('\n')
    case 'pre':
      return fence(el.textContent ?? '')
    case 'blockquote':
      return escapeBlockStart(inline(el))
        .split('\n')
        .map((line) => `> ${line}`)
        .join('\n')
    case 'formula':
      return `$$\n${el.textContent ?? ''}\n$$`
    case 'table':
      return table(el)
    default:
      return escapeBlockStart(inline(el))
  }
}

function fence(code: string): string {
  const longest = Math.max(2, ...Array.from(code.matchAll(/`+/g), (m) => m[0].length))
  const ticks = '`'.repeat(longest + 1)
  return `${ticks}\n${code.replace(/\n$/, '')}\n${ticks}`
}

function table(el: Element): string {
  const rows = Array.from(el.querySelectorAll('tr')).map((tr) =>
    Array.from(tr.children).map((cell) =>
      inline(cell).replace(/\n/g, '<br>').replace(/\|/g, '\\|'),
    ),
  )
  if (rows.length === 0) return ''
  const width = Math.max(...rows.map((r) => r.length))
  const line = (cells: string[]) =>
    `| ${Array.from({ length: width }, (_, i) => cells[i] ?? '').join(' | ')} |`
  const header = rows[0]
  const sep = `| ${Array.from({ length: width }, () => '---').join(' | ')} |`
  return [line(header), sep, ...rows.slice(1).map(line)].join('\n')
}

function inline(node: Node): string {
  let out = ''
  node.childNodes.forEach((child) => {
    if (child.nodeType === Node.TEXT_NODE) {
      out += escapeText(child.textContent ?? '')
      return
    }
    if (child.nodeType !== Node.ELEMENT_NODE) return
    const el = child as Element
    const inner = () => inline(el)
    switch (el.tagName.toLowerCase()) {
      case 'br':
        out += '\n'
        break
      case 'strong':
        out += wrap(inner(), '**')
        break
      case 'em':
        out += wrap(inner(), '*')
        break
      case 's':
        out += wrap(inner(), '~~')
        break
      case 'u':
        out += wrap(inner(), '<u>', '</u>')
        break
      case 'a': {
        const href = el.getAttribute('href') ?? ''
        const text = inner()
        out += href ? `[${text}](${href.replace(/[()\s]/g, encodeURIComponent)})` : text
        break
      }
      case 'formula':
        out += `$${el.textContent ?? ''}$`
        break
      case 'img': {
        const src = el.getAttribute('src') ?? ''
        if (src) out += `![](${src.replace(/[()\s]/g, encodeURIComponent)})`
        break
      }
      default:
        out += inner()
    }
  })
  return out
}

/** Marks hug the text: leading/trailing spaces move outside the delimiters. */
function wrap(text: string, open: string, close = open): string {
  const m = /^(\s*)([\s\S]*?)(\s*)$/.exec(text)!
  return m[2] ? `${m[1]}${open}${m[2]}${close}${m[3]}` : text
}

function escapeText(s: string): string {
  return s.replace(/[\\`*_[\]<~$]/g, '\\$&')
}

/** Every line: hard breaks and quote lines start new Markdown lines too. */
function escapeBlockStart(s: string): string {
  return s.replace(/^(\s*)(#{1,6}\s|>|[-+]\s|\d+[.)]\s|(?:[-*_]\s*){3,}$)/gm, '$1\\$2')
}

function escapeHtml(s: string): string {
  return s.replace(
    /[&<>"]/g,
    (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c]!,
  )
}
