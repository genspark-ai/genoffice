/**
 * Byte-level scanner for slideN.xml.
 *
 * Same idea as docx-engine.scanBody: never parse->serialize (that silently alters
 * untouched bytes: attribute order, self-closing form, entity encoding). Only locate
 * the exact character range of each top-level shape inside <p:spTree>
 * (p:sp / p:pic / p:graphicFrame / p:grpSp / p:cxnSp); on save, splice in new
 * fragments and copy unmodified shapes verbatim from the original substring.
 *
 * The leading <p:nvGrpSpPr> / <p:grpSpPr> inside spTree are not shapes; they go
 * into bodyPrefix.
 */

export interface SpElement {
  /** tag: p:sp / p:pic / p:graphicFrame / p:grpSp / p:cxnSp */
  name: string
  /** range [start, end) in slideN.xml string */
  start: number
  end: number
  /** Raw bytes between this shape and the next non-adjacent shape (mc:AlternateContent / p:contentPart etc.), replayed verbatim on rebuild */
  gapAfter?: string
}

export interface SlideScan {
  elements: SpElement[]
  /** From the start of <p:spTree> up to the first shape (spTree open tag + nvGrpSpPr/grpSpPr) */
  bodyPrefix: string
  /** From the end of the last shape to the end of the file (</p:spTree> and trailing content) */
  bodySuffix: string
}

// Matches one XML tag, tolerating '>' inside attribute values.
const TAG_RE = /<\/?(?:[^<>"']|"[^"]*"|'[^']*')*>/g
const NAME_RE = /^<\/?\s*([A-Za-z_][\w:.-]*)/

// mc:AlternateContent: scanned as an element so its Choice (e.g. a chartEx frame)
// can render; save replays the whole block's original bytes
const SHAPE_TAGS = new Set([
  'p:sp',
  'p:pic',
  'p:graphicFrame',
  'p:grpSp',
  'p:cxnSp',
  'mc:AlternateContent',
])

export function scanSlide(slideXml: string): SlideScan {
  const spTreeOpen = /<p:spTree(?:\s(?:[^<>"']|"[^"]*"|'[^']*')*)?>/.exec(slideXml)
  if (!spTreeOpen) {
    throw new Error('slide xml has no <p:spTree> element')
  }
  const scanFrom = spTreeOpen.index + spTreeOpen[0].length

  const elements: SpElement[] = []
  TAG_RE.lastIndex = scanFrom
  let depth = 0
  let currentStart = -1
  let currentName = ''
  let match: RegExpExecArray | null
  let spTreeCloseAt = -1

  while ((match = TAG_RE.exec(slideXml)) !== null) {
    const tag = match[0]
    if (tag.startsWith('<!--') || tag.startsWith('<![') || tag.startsWith('<?')) continue
    const isClosing = tag.startsWith('</')
    const isSelfClosing = !isClosing && tag.endsWith('/>')
    const name = NAME_RE.exec(tag)?.[1] ?? ''

    if (isClosing) {
      if (depth === 0) {
        if (name === 'p:spTree') {
          spTreeCloseAt = match.index
          break
        }
        throw new Error(`unexpected closing tag </${name}> at spTree level`)
      }
      depth--
      if (depth === 0 && SHAPE_TAGS.has(currentName)) {
        elements.push({ name: currentName, start: currentStart, end: match.index + tag.length })
      }
    } else if (isSelfClosing) {
      if (depth === 0 && SHAPE_TAGS.has(name)) {
        elements.push({ name, start: match.index, end: match.index + tag.length })
      }
      // Self-closing top-level nvGrpSpPr / grpSpPr: not a shape, ignore (handled by prefix/suffix logic)
    } else {
      if (depth === 0) {
        currentStart = match.index
        currentName = name
      }
      depth++
    }
  }

  if (spTreeCloseAt < 0) throw new Error('slide xml: unterminated <p:spTree>')

  for (let i = 0; i + 1 < elements.length; i++) {
    const gap = slideXml.slice(elements[i].end, elements[i + 1].start)
    if (gap) elements[i].gapAfter = gap
  }

  const firstShapeStart = elements.length ? elements[0].start : spTreeCloseAt
  const lastShapeEnd = elements.length ? elements[elements.length - 1].end : spTreeCloseAt

  return {
    elements,
    bodyPrefix: slideXml.slice(0, firstShapeStart),
    bodySuffix: slideXml.slice(lastShapeEnd),
  }
}
