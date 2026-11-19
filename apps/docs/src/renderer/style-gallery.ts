import type { Node as PmNode } from '@tiptap/pm/model'
import type { CSSProperties } from 'react'
import type { DocDefaults, StyleInfo } from '@genoffice/docx-engine'
import type { StringKey, TFunc } from './i18n/locale'
import { cssFontFamily } from './text-layout/font-rules'

export type StyleMap = ReadonlyMap<string, StyleInfo>

/** What the gallery lists when no document styles are known (tests, no open document) */
export const FALLBACK_STYLES: StyleMap = new Map<string, StyleInfo>(
  [
    { styleId: 'Normal', name: 'Normal', type: 'paragraph', isDefault: true, qFormat: true },
    {
      styleId: 'Heading1',
      name: 'heading 1',
      type: 'paragraph',
      headingLevel: 1,
      qFormat: true,
      uiPriority: 9,
    },
    {
      styleId: 'Heading2',
      name: 'heading 2',
      type: 'paragraph',
      headingLevel: 2,
      qFormat: true,
      uiPriority: 9,
    },
    {
      styleId: 'Heading3',
      name: 'heading 3',
      type: 'paragraph',
      headingLevel: 3,
      qFormat: true,
      uiPriority: 9,
    },
  ].map((s) => [s.styleId, s as StyleInfo]),
)

const byPriorityThenName = (a: StyleInfo, b: StyleInfo): number =>
  (a.uiPriority ?? 0) - (b.uiPriority ?? 0) ||
  a.name.localeCompare(b.name, undefined, { numeric: true, sensitivity: 'base' })

const listable = (info: StyleInfo): boolean =>
  (info.type === 'paragraph' || info.type === 'character') && !info.linkedCharShell

/**
 * Word's Quick Style gallery: every w:qFormat style in w:uiPriority order (name
 * breaks ties); a semiHidden one joins once the document uses it (Heading 3+).
 * The default paragraph style is always listed so body text stays reachable.
 */
export function quickStyleEntries(styles: StyleMap, used: ReadonlySet<string>): StyleInfo[] {
  const out: StyleInfo[] = []
  for (const info of styles.values()) {
    if (!listable(info)) continue
    const isDefault = info.type === 'paragraph' && info.isDefault
    if (!info.qFormat && !isDefault) continue
    if (info.semiHidden && !isDefault && !(info.unhideWhenUsed && used.has(info.styleId))) continue
    out.push(info)
  }
  // documents without any quick-style flags (converters, old generators) fall back to the visible styles
  if (out.every((s) => !s.qFormat)) {
    for (const info of styles.values()) {
      if (listable(info) && !info.semiHidden && !out.includes(info)) out.push(info)
    }
  }
  return out.sort(byPriorityThenName)
}

/** Styles pane "All styles": every paragraph and character style, same order */
export function allStyleEntries(styles: StyleMap): StyleInfo[] {
  return [...styles.values()].filter(listable).sort(byPriorityThenName)
}

/** The paragraph style Word applies to unstyled paragraphs (w:default), by id */
export function defaultParagraphStyleId(styles: StyleMap | undefined): string | undefined {
  if (!styles) return undefined
  for (const info of styles.values())
    if (info.type === 'paragraph' && info.isDefault) return info.styleId
  return undefined
}

/** The document's own "heading N" style (its id varies: Heading1, "1", localized ids) */
export function headingStyleId(styles: StyleMap | undefined, level: number): string | undefined {
  if (!styles) return undefined
  let fallback: string | undefined
  for (const info of styles.values()) {
    if (info.type !== 'paragraph' || info.headingLevel !== level || info.headingLevelInherited)
      continue
    if (/^heading\s*[1-9]$/i.test(info.name) || info.styleId === `Heading${level}`)
      return info.styleId
    fallback ??= info.styleId
  }
  return fallback
}

/**
 * Style ids the document body uses, for the gallery's unhide-when-used rule.
 * Only the candidates that can change the gallery are looked for (semiHidden
 * quick styles), and runs are scanned only when one of them is a character style.
 */
export function collectUsedStyleIds(doc: PmNode | null, styles: StyleMap | undefined): Set<string> {
  const used = new Set<string>()
  if (!doc || !styles) return used
  const candidates = new Set<string>()
  let needRuns = false
  for (const info of styles.values()) {
    if (info.qFormat && info.semiHidden && info.unhideWhenUsed && listable(info)) {
      candidates.add(info.styleId)
      if (info.type === 'character') needRuns = true
    }
  }
  if (candidates.size === 0) return used
  const headingIds = new Map<number, string | undefined>()
  doc.descendants((node) => {
    if (!node.isTextblock) return true
    let id = typeof node.attrs.styleId === 'string' ? node.attrs.styleId : null
    if (!id && node.type.name === 'docHeading') {
      const level = Number(node.attrs.level) || 1
      if (!headingIds.has(level)) headingIds.set(level, headingStyleId(styles, level))
      id = headingIds.get(level) ?? null
    }
    if (id && candidates.has(id)) used.add(id)
    if (needRuns) {
      node.forEach((child) => {
        for (const m of child.marks) {
          const sid = m.type.name === 'docTextStyle' ? m.attrs.styleId : null
          if (typeof sid === 'string' && candidates.has(sid)) used.add(sid)
        }
      })
    }
    return false
  })
  return used
}

/** the paragraph style a text block resolves to: its own id, else its heading level's style, else the default */
function blockStyleId(node: PmNode, styles: StyleMap): string | undefined {
  if (typeof node.attrs.styleId === 'string') return node.attrs.styleId
  if (node.type.name === 'docHeading') {
    const level = Number(node.attrs.level) || 1
    return headingStyleId(styles, level) ?? `Heading${level}`
  }
  return defaultParagraphStyleId(styles)
}

/** Styles pane "In use": every paragraph / character style some text block or run resolves to */
export function collectAllUsedStyleIds(doc: PmNode | null, styles: StyleMap): Set<string> {
  const used = new Set<string>()
  if (!doc) return used
  doc.descendants((node) => {
    if (!node.isTextblock) return true
    const id = blockStyleId(node, styles)
    if (id) used.add(id)
    node.forEach((child) => {
      for (const m of child.marks) {
        if (m.type.name === 'docTextStyle' && typeof m.attrs.styleId === 'string')
          used.add(m.attrs.styleId)
      }
    })
    return false
  })
  return used
}

export interface DocRange {
  readonly from: number
  readonly to: number
}

/**
 * Word's "Select All N Instance(s)": the text ranges carrying a style — each
 * text block resolving to a paragraph style, or each run (adjacent runs merged)
 * marked with a character style. Empty blocks are skipped: there is nothing to
 * select in them.
 */
export function styleInstanceRanges(doc: PmNode, info: StyleInfo, styles: StyleMap): DocRange[] {
  const out: DocRange[] = []
  const push = (from: number, to: number) => {
    const last = out[out.length - 1]
    if (last && last.to === from) out[out.length - 1] = { from: last.from, to }
    else out.push({ from, to })
  }
  doc.descendants((node, pos) => {
    if (!node.isTextblock) return true
    if (info.type === 'character') {
      node.forEach((child, offset) => {
        const hit = child.marks.some(
          (m) => m.type.name === 'docTextStyle' && m.attrs.styleId === info.styleId,
        )
        if (hit) push(pos + 1 + offset, pos + 1 + offset + child.nodeSize)
      })
    } else if (node.content.size > 0 && blockStyleId(node, styles) === info.styleId) {
      push(pos + 1, pos + 1 + node.content.size)
    }
    return false
  })
  return out
}

/** gallery / pane highlight: the character style at the caret wins, else the paragraph's style */
export function activeStyleKey(
  fs: { charStyleId: string | null; paraStyleId: string | null; headingLevel: number | null },
  styles: StyleMap | undefined,
): string {
  if (fs.charStyleId) return `char:${fs.charStyleId}`
  const paraId =
    fs.paraStyleId ??
    (fs.headingLevel !== null
      ? (headingStyleId(styles, fs.headingLevel) ?? `Heading${fs.headingLevel}`)
      : (defaultParagraphStyleId(styles) ?? 'Normal'))
  return `para:${paraId}`
}

export const styleKeyOf = (info: StyleInfo): string =>
  `${info.type === 'character' ? 'char' : 'para'}:${info.styleId}`

/** Word's built-in style names, localized like Word does in its gallery */
const BUILT_IN_LABELS: Record<string, StringKey> = {
  normal: 'ribbonStyleNormal',
  'no spacing': 'ribbonStyleNoSpacing',
  'heading 1': 'ribbonStyleHeading1',
  'heading 2': 'ribbonStyleHeading2',
  'heading 3': 'ribbonStyleHeading3',
  title: 'ribbonStyleTitle',
  subtitle: 'ribbonStyleSubtitle',
  'subtle emphasis': 'ribbonStyleSubtleEmphasis',
  emphasis: 'ribbonStyleEmphasis',
  'intense emphasis': 'ribbonStyleIntenseEmphasis',
  strong: 'ribbonStyleStrong',
  quote: 'ribbonStyleQuote',
  'intense quote': 'ribbonStyleIntenseQuote',
  'subtle reference': 'ribbonStyleSubtleReference',
  'intense reference': 'ribbonStyleIntenseReference',
  'book title': 'ribbonStyleBookTitle',
  'list paragraph': 'ribbonStyleListParagraph',
}

export function styleLabel(info: Pick<StyleInfo, 'name'>, t: TFunc): string {
  const name = info.name.trim()
  const key = BUILT_IN_LABELS[name.toLowerCase()]
  if (key) return t(key)
  const heading = /^heading\s*([1-9])$/i.exec(name)
  if (heading) return t('ribbonStyleHeadingN', { n: heading[1] })
  return name.charAt(0).toUpperCase() + name.slice(1)
}

/** card preview: the style's resolved face, weight, slant, color and caps at gallery scale */
export function stylePreviewCss(info: StyleInfo, docDefaults?: DocDefaults): CSSProperties {
  const d = info.display ?? {}
  const css: CSSProperties = {}
  const face = d.fontAscii ?? d.font
  if (face) css.fontFamily = cssFontFamily(face)
  if (info.type === 'paragraph') {
    const pt = (d.sizeHalfPoints ?? docDefaults?.sizeHalfPoints ?? 22) / 2
    css.fontSize = `${Math.min(20, Math.max(11, Math.round((pt * 15) / 11)))}px`
  }
  if (d.bold ?? docDefaults?.bold) css.fontWeight = 700
  if (d.italic ?? docDefaults?.italic) css.fontStyle = 'italic'
  const deco = [d.underline ? 'underline' : '', d.strike ? 'line-through' : ''].filter(Boolean)
  if (deco.length) css.textDecoration = deco.join(' ')
  if (d.caps === 'small') css.fontVariant = 'small-caps'
  else if (d.caps === 'all') css.textTransform = 'uppercase'
  const color = d.color ?? docDefaults?.color
  if (color && color !== 'auto') css.color = `#${color}`
  return css
}
