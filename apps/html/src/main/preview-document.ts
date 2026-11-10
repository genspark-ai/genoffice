export const PREVIEW_SCHEME = 'html-preview'
export const ASSET_SCHEME = 'html-asset'

/** Encode an absolute directory as an html-asset base URL with a trailing slash.
 * The scheme is registered as standard, so it needs a host; `local` is a fixed placeholder. */
export function assetBaseHref(documentDir: string): string {
  const normalized = documentDir.replace(/\\/g, '/').replace(/\/+$/, '')
  const path = normalized.startsWith('/') ? normalized : `/${normalized}`
  const encoded = path
    .split('/')
    .map((segment) => encodeURIComponent(segment).replace(/%3A/gi, ':'))
    .join('/')
  return `${ASSET_SCHEME}://local${encoded}/`
}

/**
 * The preview copy of the document: the buffer text with a <base> pointing at
 * the document's directory (so relative assets resolve through html-asset://)
 * unless the author already declared one. The saved file never contains it.
 */
interface TagRange {
  start: number
  end: number
}

function tagEnd(text: string, start: number): number {
  let quote = ''
  for (let index = start + 1; index < text.length; index += 1) {
    const character = text[index]
    if (quote) {
      if (character === quote) quote = ''
      continue
    }
    if (character === '"' || character === "'") {
      quote = character
      continue
    }
    if (character === '>') return index + 1
  }
  return -1
}

const RAW_TEXT_TAGS = new Set(['script', 'style', 'textarea', 'title'])

function skipRawText(text: string, index: number, tagName: string): number {
  const closing = new RegExp(`</${tagName}\\s*>`, 'i').exec(text.slice(index))
  return closing ? index + closing.index + closing[0].length : text.length
}

function skipTemplate(text: string, index: number): number {
  let depth = 1
  while (index < text.length) {
    const start = text.indexOf('<', index)
    if (start < 0) return text.length
    if (text.startsWith('<!--', start)) {
      const end = text.indexOf('-->', start + 4)
      index = end < 0 ? text.length : end + 3
      continue
    }
    if (text.startsWith('<!', start) || text.startsWith('<?', start)) {
      const end = tagEnd(text, start)
      index = end < 0 ? text.length : end
      continue
    }
    const match = /^<(\/?)([A-Za-z][A-Za-z0-9:-]*)/.exec(text.slice(start))
    if (!match) {
      index = start + 1
      continue
    }
    const end = tagEnd(text, start)
    if (end < 0) return text.length
    const closing = match[1] === '/'
    const tagName = match[2].toLowerCase()
    if (tagName === 'template') {
      if (closing) {
        depth -= 1
        if (depth === 0) return end
      } else {
        depth += 1
      }
    }
    index = end
    if (!closing && RAW_TEXT_TAGS.has(tagName)) index = skipRawText(text, index, tagName)
  }
  return text.length
}

function openingTag(text: string, name: string): TagRange | null {
  let index = 0
  while (index < text.length) {
    const start = text.indexOf('<', index)
    if (start < 0) return null
    if (text.startsWith('<!--', start)) {
      const end = text.indexOf('-->', start + 4)
      index = end < 0 ? text.length : end + 3
      continue
    }
    if (text.startsWith('<!', start) || text.startsWith('<?', start)) {
      const end = tagEnd(text, start)
      index = end < 0 ? text.length : end
      continue
    }
    const match = /^<([A-Za-z][A-Za-z0-9:-]*)/.exec(text.slice(start))
    if (!match) {
      index = start + 1
      continue
    }
    const end = tagEnd(text, start)
    if (end < 0) return null
    const tagName = match[1].toLowerCase()
    if (tagName === name.toLowerCase()) return { start, end }
    if (tagName === 'template') {
      index = skipTemplate(text, end)
      continue
    }
    index = end
    if (RAW_TEXT_TAGS.has(tagName)) index = skipRawText(text, index, tagName)
  }
  return null
}

export function buildPreviewDocument(text: string, baseHref: string | null): string {
  if (!baseHref || openingTag(text, 'base')) return text
  const tag = `<base href="${baseHref.replace(/"/g, '%22')}">`
  const head = openingTag(text, 'head')
  if (head) {
    const at = head.end
    return text.slice(0, at) + tag + text.slice(at)
  }
  const html = openingTag(text, 'html')
  if (html) {
    const at = html.end
    return text.slice(0, at) + tag + text.slice(at)
  }
  return tag + text
}

export function previewUrlFor(webContentsId: number): string {
  return `${PREVIEW_SCHEME}://view-${webContentsId}/`
}
