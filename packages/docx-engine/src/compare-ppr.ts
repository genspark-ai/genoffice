/**
 * Paragraph properties in a comparison: what two paragraphs' own w:pPr are
 * compared by, and the w:pPrChange record that keeps the original side's.
 */
import { childElements } from './compare-story'
import { stripPPrChange } from './generate'

/** children of w:pPr that are not the paragraph's formatting: the mark's run properties, a section, the history */
const NOT_COMPARED = new Set(['w:rPr', 'w:sectPr', 'w:pPrChange'])

const TOGGLES = new Set([
  'w:keepNext',
  'w:keepLines',
  'w:pageBreakBefore',
  'w:widowControl',
  'w:suppressLineNumbers',
  'w:suppressAutoHyphens',
  'w:kinsoku',
  'w:wordWrap',
  'w:overflowPunct',
  'w:topLinePunct',
  'w:autoSpaceDE',
  'w:autoSpaceDN',
  'w:bidi',
  'w:adjustRightInd',
  'w:snapToGrid',
  'w:contextualSpacing',
  'w:mirrorIndents',
  'w:suppressOverlap',
  'w15:collapsed',
])

interface PPrChild {
  name: string
  xml: string
  inner: string
}

function formattingChildren(rawPPr: string | undefined): PPrChild[] {
  if (!rawPPr) return []
  const pPr = childElements(rawPPr)[0]
  if (!pPr || pPr.name !== 'w:pPr') return []
  return childElements(pPr.inner).filter((child) => !NOT_COMPARED.has(child.name))
}

/** one tag with its attributes in a fixed order, so two writers' spellings compare equal */
const canonicalTags = (xml: string): string =>
  xml.replace(
    /<([\w:]+)((?:\s+[\w:]+\s*=\s*(?:"[^"]*"|'[^']*'))*)\s*(\/?)>/g,
    (_all, name: string, attrs: string, self: string) => {
      const sorted = [...attrs.matchAll(/([\w:]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g)]
        .filter((m) => !m[1]!.startsWith('w:rsid'))
        .map((m) => `${m[1]}="${m[2] ?? m[3] ?? ''}"`)
        .sort()
      return `<${name}${sorted.map((a) => ` ${a}`).join('')}${self}>`
    },
  )

function comparable(child: PPrChild, listKey: (numId: string, ilvl: number) => string): string {
  if (child.name === 'w:numPr') {
    const value = (tag: string) =>
      new RegExp(`<w:${tag}\\b[^>]*\\bw:val\\s*=\\s*["']([^"']*)["']`).exec(child.inner)?.[1]
    const numId = value('numId') ?? ''
    const ilvl = parseInt(value('ilvl') ?? '0', 10) || 0
    return `numPr ${ilvl} ${numId === '0' || numId === '' ? 'none' : listKey(numId, ilvl)}`
  }
  const xml = canonicalTags(child.xml)
  if (TOGGLES.has(child.name)) return xml.replace(/^(<[\w:]+) w:val="(?:1|true|on)"\/>$/, '$1/>')
  return xml
}

/**
 * What a paragraph's own formatting is compared by. A list is its level's
 * definition (`listKey`), not its numId: the id belongs to one package.
 */
export function paragraphFormatKey(
  rawPPr: string | undefined,
  listKey: (numId: string, ilvl: number) => string,
): string {
  return formattingChildren(rawPPr)
    .map((child) => comparable(child, listKey))
    .sort()
    .join('')
}

/**
 * The formatting of a paragraph of the other package as the record of a
 * change: `keep` says which style and list references mean the same here.
 */
export function recordedPPr(
  rawPPr: string | undefined,
  keep: { style: (styleId: string) => boolean; list: (numId: string, ilvl: number) => boolean },
): string {
  return formattingChildren(rawPPr)
    .filter((child) => {
      if (child.name === 'w:pStyle') {
        const id = /\bw:val\s*=\s*["']([^"']*)["']/.exec(child.xml)?.[1]
        return !!id && keep.style(id)
      }
      if (child.name === 'w:numPr') {
        const value = (tag: string) =>
          new RegExp(`<w:${tag}\\b[^>]*\\bw:val\\s*=\\s*["']([^"']*)["']`).exec(child.inner)?.[1]
        const numId = value('numId') ?? ''
        return numId === '0' || keep.list(numId, parseInt(value('ilvl') ?? '0', 10) || 0)
      }
      return true
    })
    .map((child) =>
      child.name === 'w:numPr'
        ? child.xml.replace(/<w:(?:numberingChange|ins)\b[^>]*(?:\/>|>[\s\S]*?<\/w:\w+>)/g, '')
        : child.xml,
    )
    .join('')
}

/** the w:pPrChange a paragraph already carries: its opening attributes and the properties it recorded */
export function existingPPrChange(
  rawPPr: string | undefined,
): { author: string; date?: string; recorded: string } | null {
  if (!rawPPr || !rawPPr.includes('<w:pPrChange')) return null
  const pPr = childElements(rawPPr)[0]
  const change = pPr ? childElements(pPr.inner).find((c) => c.name === 'w:pPrChange') : undefined
  if (!change) return null
  const attr = (name: string) => new RegExp(`\\b${name}\\s*=\\s*"([^"]*)"`).exec(change.open)?.[1]
  const date = attr('w:date')
  return {
    author: attr('w:author') ?? '',
    ...(date ? { date } : {}),
    recorded: `<w:pPr>${childElements(change.inner).find((c) => c.name === 'w:pPr')?.inner ?? ''}</w:pPr>`,
  }
}

/** `rawPPr` with a w:pPrChange recording `recorded` (inner children of the old w:pPr); it replaces one already there */
export function withPPrChange(rawPPr: string | undefined, recorded: string, attrs: string): string {
  const change = `<w:pPrChange${attrs}><w:pPr>${recorded}</w:pPr></w:pPrChange>`
  const clean = rawPPr ? stripPPrChange(rawPPr) : ''
  const close = clean.lastIndexOf('</w:pPr>')
  if (close === -1) return `<w:pPr>${change}</w:pPr>`
  return clean.slice(0, close) + change + clean.slice(close)
}

/** a paragraph's own XML with its w:pPr replaced; null when the paragraph is not a bare w:p */
export function paragraphWithPPr(
  paragraphXml: string | null | undefined,
  rawPPr: string | undefined,
  nextPPr: string,
): string | null {
  if (!paragraphXml || !/^<w:p[\s/>]/.test(paragraphXml)) return null
  if (/^<w:p(?:\s[^>]*)?\/>$/.test(paragraphXml))
    return `${paragraphXml.slice(0, -2)}>${nextPPr}</w:p>`
  const openEnd = paragraphXml.indexOf('>') + 1
  if (rawPPr === undefined)
    return paragraphXml.slice(0, openEnd) + nextPPr + paragraphXml.slice(openEnd)
  if (!paragraphXml.startsWith(rawPPr, openEnd)) return null
  return paragraphXml.slice(0, openEnd) + nextPPr + paragraphXml.slice(openEnd + rawPPr.length)
}
