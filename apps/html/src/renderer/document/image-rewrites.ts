import { parse, type DefaultTreeAdapterTypes as T } from 'parse5'

export interface ImageRewrite {
  from: string
  to: string
}

export interface ImageRewriteAdoption {
  savedText: string
  liveText: string
}

interface ImageSourceRange {
  start: number
  end: number
  source: string
  quote: '"' | "'" | null
}

function isElement(node: T.Node): node is T.Element {
  return 'tagName' in node && typeof node.tagName === 'string'
}

function imageSourceRange(html: string, element: T.Element): ImageSourceRange | null {
  const attributes = element.attrs.filter((attribute) => attribute.name.toLowerCase() === 'src')
  const location = element.sourceCodeLocation?.attrs?.src
  if (attributes.length !== 1 || !location) return null
  const raw = html.slice(location.startOffset, location.endOffset)
  const equals = raw.indexOf('=')
  if (equals < 0) return null
  let cursor = equals + 1
  while (/\s/.test(raw[cursor] ?? '')) cursor += 1
  if (cursor >= raw.length) return null
  const quote = raw[cursor] === '"' ? '"' : raw[cursor] === "'" ? "'" : null
  let start: number
  let end: number
  if (quote) {
    start = cursor + 1
    end = raw.indexOf(quote, start)
    if (end < 0) return null
  } else {
    start = cursor
    end = cursor
    while (end < raw.length) {
      const character = raw[end]!
      if (
        /\s/.test(character) ||
        character === '>' ||
        (character === '/' && raw[end + 1] === '>')
      ) {
        break
      }
      end += 1
    }
    if (start === end) return null
  }
  return {
    start: location.startOffset + start,
    end: location.startOffset + end,
    source: attributes[0]!.value,
    quote,
  }
}

function imageSourceRanges(html: string): ImageSourceRange[] {
  try {
    const ranges: ImageSourceRange[] = []
    const document = parse(html, { sourceCodeLocationInfo: true })
    const visit = (node: T.Node) => {
      if (isElement(node) && node.tagName === 'img') {
        const range = imageSourceRange(html, node)
        if (range) ranges.push(range)
      }
      if (isElement(node) && node.tagName === 'template') {
        visit((node as T.Template).content)
        return
      }
      if ('childNodes' in node) for (const child of node.childNodes) visit(child)
    }
    visit(document)
    return ranges
  } catch {
    return []
  }
}

function encodeAttribute(value: string, quote: '"' | "'" | null): string {
  let encoded = ''
  for (const character of value) {
    if (character === '&') encoded += '&amp;'
    else if (character === '<') encoded += '&lt;'
    else if (quote === '"' && character === '"') encoded += '&quot;'
    else if (quote === "'" && character === "'") encoded += '&#39;'
    else if (quote === null && /[\s"'`=<>]/.test(character)) {
      encoded += `&#${character.codePointAt(0)!};`
    } else encoded += character
  }
  return encoded
}

export function rewriteDocumentImageSources(
  html: string,
  rewrites: ReadonlyMap<string, string>,
): string {
  if (rewrites.size === 0) return html
  const ranges = imageSourceRanges(html).sort((left, right) => right.start - left.start)
  let output = html
  let boundary = html.length
  for (const range of ranges) {
    if (range.end > boundary) continue
    const replacement = rewrites.get(range.source)
    if (replacement === undefined || replacement === range.source) continue
    output = `${output.slice(0, range.start)}${encodeAttribute(replacement, range.quote)}${output.slice(range.end)}`
    boundary = range.start
  }
  return output
}

export function adoptImageRewrites(
  textAtSave: string,
  liveText: string,
  rewrites: readonly ImageRewrite[] | undefined,
): ImageRewriteAdoption {
  const bySource = new Map((rewrites ?? []).map(({ from, to }) => [from, to]))
  return {
    savedText: rewriteDocumentImageSources(textAtSave, bySource),
    liveText: rewriteDocumentImageSources(liveText, bySource),
  }
}
