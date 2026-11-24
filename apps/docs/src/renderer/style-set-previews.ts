import type { CSSProperties } from 'react'
import {
  STYLE_SET_IDS,
  customStyleSetUpserts,
  previewFontSettings,
  readDocxStylesXml,
  styleSetUpserts,
  type DefaultFonts,
  type DocDefaults,
  type ParaBorderLine,
  type ParsedDocFull,
  type StyleInfo,
  type StyleSetId,
  type StyleUpsert,
  type ThemeEdit,
} from '@genoffice/docx-engine'
import { cssFontFamily } from './text-layout/font-rules'
import type { CustomStyleSet } from './custom-style-sets'
import { textOutlineCssValue } from './editor/text-outline'

/** what the previews are resolved against: the open document and its unsaved style definitions */
export interface StyleSetSource {
  parsed: ParsedDocFull
  pending: StyleUpsert[]
  defaultFonts?: DefaultFonts
  theme?: ThemeEdit
  /** the style sets saved on this computer */
  custom?: CustomStyleSet[]
}

/** the three styles a gallery card shows; `id` null with no `custom` is the document as it stands */
export interface StyleSetPreview {
  id: StyleSetId | null
  custom?: CustomStyleSet
  title?: StyleInfo
  heading?: StyleInfo
  body?: StyleInfo
  docDefaults?: DocDefaults
}

export async function styleSetPreviews(source: StyleSetSource): Promise<StyleSetPreview[]> {
  const { parsed, pending, defaultFonts, theme } = source
  const xml = await readDocxStylesXml(parsed.internal.originalBytes, pending)
  const out: StyleSetPreview[] = []
  for (const set of [null, ...(source.custom ?? []), ...STYLE_SET_IDS]) {
    const custom = set !== null && typeof set === 'object' ? set : undefined
    const id = typeof set === 'string' ? set : null
    const look = custom
      ? customStyleSetUpserts(custom.xml, xml)
      : id
        ? styleSetUpserts(id, xml)
        : []
    const upserts = [...pending, ...look]
    const { styles, docDefaults } = await previewFontSettings(parsed, upserts, defaultFonts, theme)
    const paragraphs = [...styles.values()].filter((s) => s.type === 'paragraph')
    const named = (name: string) => paragraphs.find((s) => s.name.toLowerCase() === name)
    out.push({
      id,
      ...(custom ? { custom } : {}),
      title: named('title'),
      heading: named('heading 1'),
      body: paragraphs.find((s) => s.isDefault),
      docDefaults,
    })
  }
  return out
}

const pt = (twips: number | undefined, max: number): string | undefined =>
  twips ? `${Math.min(max, twips / 20)}pt` : undefined

const borderCss = (line: ParaBorderLine | null | undefined): string | undefined =>
  line && !line.none
    ? `${Math.max(0.75, line.szPt ?? 0.5)}pt ${line.val === 'dotted' ? 'dotted' : 'solid'} ${line.color ? `#${line.color}` : 'currentColor'}`
    : undefined

/** a line of the card's miniature page, in the style's own look; the colours are the document's */
export function styleSetLineCss(
  info: StyleInfo | undefined,
  docDefaults?: DocDefaults,
): CSSProperties {
  const d = info?.display ?? {}
  const css: CSSProperties = {}
  const face = d.fontAscii ?? d.font ?? docDefaults?.asciiFont
  if (face) css.fontFamily = cssFontFamily(face)
  css.fontSize = `${(d.sizeHalfPoints ?? docDefaults?.sizeHalfPoints ?? 20) / 2}pt`
  if (d.bold ?? docDefaults?.bold) css.fontWeight = 700
  if (d.italic ?? docDefaults?.italic) css.fontStyle = 'italic'
  if (d.caps === 'small') css.fontVariant = 'small-caps'
  else if (d.caps === 'all') css.textTransform = 'uppercase'
  const color = d.color ?? docDefaults?.color
  if (color && color !== 'auto') css.color = `#${color}`
  if (d.align && d.align !== 'distribute') css.textAlign = d.align
  if (d.charSpacingTwips) css.letterSpacing = `${d.charSpacingTwips / 20}pt`
  if (d.shadingFill && d.shadingFill !== 'auto') css.background = `#${d.shadingFill}`
  if (d.textOutline) css.WebkitTextStroke = textOutlineCssValue(d.textOutline)
  css.marginTop = pt(d.spaceBeforeTwips, 12)
  css.marginBottom = pt(d.spaceAfterTwips, 12)
  css.marginInlineStart = pt(d.indentLeftTwips, 120)
  css.textIndent = pt(d.indentFirstLineTwips, 36)
  css.lineHeight =
    d.lineRule && d.lineRule !== 'auto' ? undefined : Math.min(2, d.lineSpacing ?? 1) * 1.2
  css.borderTop = borderCss(d.borderSides?.t)
  css.borderBottom = borderCss(d.borderSides?.b)
  css.borderLeft = borderCss(d.borderSides?.l)
  css.borderRight = borderCss(d.borderSides?.r)
  return css
}
