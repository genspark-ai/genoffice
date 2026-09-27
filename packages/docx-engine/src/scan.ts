/**
 * Byte-exact scanner for word/document.xml.
 *
 * We never re-serialize the whole XML (parse->serialize would silently change
 * untouched bytes: attribute order, self-closing forms, entity forms...).
 * Instead we locate the exact character range of every top-level element in
 * <w:body> so that patch-save can splice new fragments while copying untouched
 * elements as raw original substrings.
 */

export interface BodyElement {
  /** tag name, e.g. "w:p", "w:tbl", "w:sectPr" */
  name: string
  /** range [start, end) in the document.xml string */
  start: number
  end: number
}

export interface OpaqueRegion {
  start: number
  end: number
}

export interface BodyScan {
  elements: BodyElement[]
  opaqueRegions: OpaqueRegion[]
  /** range [innerStart, innerEnd) spanning from the first top-level element start to the last element end */
  innerStart: number
  innerEnd: number
  bodyContentStart: number
  bodyContentEnd: number
}

interface Tag {
  start: number
  end: number
  name: string
  closing: boolean
  selfClosing: boolean
}

const NAME_RE = /^<\/?\s*([A-Za-z_][\w:.-]*)/

/**
 * Elements whose character content is document text. A CDATA section inside one
 * is text the parser already carries into the runs, not position-free
 * decoration, so it must not become an opaque region (the save path would
 * re-append it to the regenerated element and duplicate the text).
 */
const TEXT_ELEMENTS = new Set(['w:t', 'w:delText', 'w:instrText', 'w:delInstrText', 'a:t', 'm:t'])

function opaqueEnd(documentXml: string, start: number): number | null {
  if (documentXml.startsWith('<!--', start)) {
    const at = documentXml.indexOf('-->', start + 4)
    return at === -1 ? documentXml.length : at + 3
  }
  if (documentXml.startsWith('<![CDATA[', start)) {
    const at = documentXml.indexOf(']]>', start + 9)
    return at === -1 ? documentXml.length : at + 3
  }
  if (documentXml.startsWith('<?', start)) {
    const at = documentXml.indexOf('?>', start + 2)
    return at === -1 ? documentXml.length : at + 2
  }
  if (!documentXml.startsWith('<!', start)) return null
  let quote = ''
  let subsetDepth = 0
  for (let index = start + 2; index < documentXml.length; index += 1) {
    const character = documentXml[index]
    if (quote) {
      if (character === quote) quote = ''
      continue
    }
    if (character === '"' || character === "'") {
      quote = character
    } else if (character === '[') {
      subsetDepth += 1
    } else if (character === ']') {
      subsetDepth = Math.max(0, subsetDepth - 1)
    } else if (character === '>' && subsetDepth === 0) {
      return index + 1
    }
  }
  return documentXml.length
}

function tagAt(documentXml: string, start: number): Tag {
  let quote = ''
  let end = documentXml.length
  for (let index = start + 1; index < documentXml.length; index += 1) {
    const character = documentXml[index]
    if (quote) {
      if (character === quote) quote = ''
    } else if (character === '"' || character === "'") {
      quote = character
    } else if (character === '>') {
      end = index + 1
      break
    }
  }
  const text = documentXml.slice(start, end)
  const closing = text.startsWith('</')
  return {
    start,
    end,
    name: NAME_RE.exec(text)?.[1] ?? '',
    closing,
    selfClosing: !closing && text.endsWith('/>'),
  }
}

export function scanBody(documentXml: string): BodyScan {
  const elements: BodyElement[] = []
  const opaqueRegions: OpaqueRegion[] = []
  let cursor = 0
  let bodyContentStart = -1
  let bodyContentEnd = -1
  let inBody = false
  const stack: Array<{ name: string; start: number }> = []

  while (cursor < documentXml.length) {
    const start = documentXml.indexOf('<', cursor)
    if (start === -1) break
    const opaque = opaqueEnd(documentXml, start)
    if (opaque !== null) {
      const textCdata =
        documentXml.startsWith('<![CDATA[', start) &&
        stack.length > 0 &&
        TEXT_ELEMENTS.has(stack[stack.length - 1]!.name)
      if (bodyContentStart !== -1 && !textCdata) opaqueRegions.push({ start, end: opaque })
      cursor = opaque
      continue
    }
    const tag = tagAt(documentXml, start)
    cursor = tag.end

    if (!inBody) {
      if (!tag.closing && tag.name === 'w:body') {
        if (bodyContentStart === -1) bodyContentStart = tag.end
        if (!tag.selfClosing) inBody = true
        else if (bodyContentEnd === -1) bodyContentEnd = tag.end
      }
      continue
    }

    if (stack.length === 0 && tag.closing && tag.name === 'w:body') {
      bodyContentEnd = start
      inBody = false
      continue
    }

    if (tag.closing) {
      if (stack.length === 0) {
        throw new Error(`unexpected closing tag </${tag.name}> at body level`)
      }
      const current = stack.pop()
      if (stack.length === 0 && current) {
        elements.push({ name: current.name, start: current.start, end: tag.end })
      }
      continue
    }

    if (tag.selfClosing) {
      if (stack.length === 0) elements.push({ name: tag.name, start: tag.start, end: tag.end })
      continue
    }

    stack.push({ name: tag.name, start: tag.start })
  }

  if (bodyContentStart === -1) {
    throw new Error('document.xml has no <w:body> element')
  }
  if (bodyContentEnd === -1) bodyContentEnd = documentXml.length
  const bodyOpaqueRegions = opaqueRegions.filter(
    (region) => region.start >= bodyContentStart && region.end <= bodyContentEnd,
  )
  if (elements.length === 0) {
    return {
      elements,
      opaqueRegions: bodyOpaqueRegions,
      innerStart: bodyContentStart,
      innerEnd: bodyContentEnd,
      bodyContentStart,
      bodyContentEnd,
    }
  }
  return {
    elements,
    opaqueRegions: bodyOpaqueRegions,
    innerStart: elements[0].start,
    innerEnd: elements[elements.length - 1].end,
    bodyContentStart,
    bodyContentEnd,
  }
}
