/**
 * ODF style index: every `style:style` from content.xml's automatic styles and
 * styles.xml's named styles, resolved along `style:parent-style-name` chains
 * onto the flat text/paragraph property sets the two converters consume.
 * Anything the chain cannot prove (a missing parent, a cycle) degrades to the
 * properties gathered so far — an import filter never guesses.
 */

import { XNode, attr, firstChild, nodeAttrs, nodeChildren, nodeName } from './xml'
import { emuToPt, emuToTwips, odfColorToHex, odfLengthToEmu } from './units'

export interface OdfTextProps {
  bold?: boolean
  italic?: boolean
  underline?: boolean
  strike?: boolean
  sizePt?: number
  /** hex without '#' */
  color?: string
  font?: string
  vertAlign?: 'superscript' | 'subscript'
  caps?: 'all' | 'small'
  /** run shading fill, hex without '#' */
  shading?: string
}

export interface OdfParaProps {
  align?: 'left' | 'center' | 'right' | 'justify'
  spaceBeforeTw?: number
  spaceAfterTw?: number
  /** multiple of single spacing (lineRule auto) */
  lineSpacing?: number
  /** fixed/at-least line height in twips */
  lineAtLeastTw?: number
  lineExactTw?: number
  indentLeftTw?: number
  indentRightTw?: number
  indentFirstLineTw?: number
  pageBreakBefore?: boolean
}

export interface OdfListLevel {
  kind: 'bullet' | 'ordered'
  bulletChar?: string
  numFormat?: string
  numPrefix?: string
  numSuffix?: string
  indentLeftTw?: number
  hangingTw?: number
}

interface StyleDef {
  family: string
  parent?: string
  text?: OdfTextProps
  para?: OdfParaProps
}

export class OdfStyles {
  private defs = new Map<string, StyleDef>()
  private fonts = new Map<string, string>()
  private lists = new Map<string, Map<number, OdfListLevel>>()
  private columnWidths = new Map<string, number>()
  private cellFills = new Map<string, string>()
  private defaults: { text?: OdfTextProps; para?: OdfParaProps } = {}

  constructor(contentRoots: XNode[] | null, stylesRoots: XNode[] | null) {
    // named styles first, then automatic ones (later index wins on name clash —
    // automatic styles are the ones the body actually references)
    if (stylesRoots) this.collect(stylesRoots)
    if (contentRoots) this.collect(contentRoots)
  }

  private collect(roots: XNode[]): void {
    for (const root of roots) this.walkCollect(root)
  }

  private walkCollect(node: XNode): void {
    const name = nodeName(node)
    if (name === 'style:style') {
      this.addStyle(node)
    } else if (name === 'style:default-style') {
      const family = attr(node, 'style:family')
      if (family === 'paragraph' || family === 'text') {
        const props = this.readProps(node)
        if (family === 'text' && props.text)
          this.defaults.text = { ...this.defaults.text, ...props.text }
        if (family === 'paragraph' && props.para)
          this.defaults.para = { ...this.defaults.para, ...props.para }
        if (family === 'paragraph' && props.text)
          this.defaults.text = { ...this.defaults.text, ...props.text }
      }
    } else if (name === 'text:list-style') {
      this.addListStyle(node)
    } else if (name === 'style:font-face') {
      const declName = attr(node, 'style:name')
      const family = attr(node, 'svg:font-family')
      if (declName && family) {
        this.fonts.set(declName, family.replace(/^['"]|['"]$/g, ''))
      }
    }
    for (const child of nodeChildren(node)) this.walkCollect(child)
  }

  private addStyle(node: XNode): void {
    const name = attr(node, 'style:name')
    const family = attr(node, 'style:family')
    if (!name || !family) return
    if (family === 'table-column') {
      const properties = firstChild(node, 'style:column-properties')
      const width = properties ? odfLengthToEmu(attr(properties, 'style:column-width')) : null
      if (width !== null && width > 0) this.columnWidths.set(name, width)
      return
    }
    if (family === 'table-cell') {
      const properties = firstChild(node, 'style:table-cell-properties')
      const fill = properties ? odfColorToHex(attr(properties, 'fo:background-color')) : null
      if (fill) this.cellFills.set(name, fill)
      return
    }
    const props = this.readProps(node)
    this.defs.set(`${family}:${name}`, {
      family,
      parent: attr(node, 'style:parent-style-name'),
      text: props.text,
      para: props.para,
    })
  }

  /** width of a table-column style, EMU; null when undeclared */
  columnWidthEmu(styleName: string | undefined): number | null {
    if (!styleName) return null
    return this.columnWidths.get(styleName) ?? null
  }

  /** fill of a table-cell style, hex without '#'; null otherwise */
  cellFill(styleName: string | undefined): string | null {
    if (!styleName) return null
    return this.cellFills.get(styleName) ?? null
  }

  private readProps(node: XNode): { text?: OdfTextProps; para?: OdfParaProps } {
    let text: OdfTextProps | undefined
    let para: OdfParaProps | undefined
    for (const child of nodeChildren(node)) {
      const childName = nodeName(child)
      if (childName === 'style:text-properties') text = this.readTextProps(child)
      else if (childName === 'style:paragraph-properties') para = this.readParaProps(child)
    }
    return { text, para }
  }

  private readTextProps(node: XNode): OdfTextProps | undefined {
    const a = nodeAttrs(node)
    const out: OdfTextProps = {}
    if (a['fo:font-weight'] === 'bold') out.bold = true
    if (a['fo:font-style'] === 'italic') out.italic = true
    if (a['style:text-underline-type'] !== undefined) {
      out.underline = a['style:text-underline-type'] !== 'none'
    } else if (a['style:text-underline-style'] !== undefined) {
      out.underline = a['style:text-underline-style'] !== 'none'
    }
    const strike =
      a['style:text-line-through-type'] ??
      a['fo:text-line-through-type'] ??
      a['style:text-line-through-style']
    if (strike !== undefined) out.strike = strike !== 'none'
    const size = odfLengthToEmu(a['fo:font-size'])
    if (size !== null && size > 0) out.sizePt = Math.round(emuToPt(size) * 10) / 10
    const color = odfColorToHex(a['fo:color'])
    if (color) out.color = color
    const shading = odfColorToHex(a['fo:background-color'])
    if (shading) out.shading = shading
    const fontName = a['style:font-name']
    if (fontName) out.font = this.fonts.get(fontName) ?? fontName
    const position = a['style:text-position']
    if (position) {
      if (/^super/i.test(position.trim())) out.vertAlign = 'superscript'
      else if (/^sub/i.test(position.trim())) out.vertAlign = 'subscript'
    }
    const transform = a['style:text-transform'] ?? a['fo:text-transform']
    if (transform === 'uppercase') out.caps = 'all'
    else if (transform === 'lowercase') out.caps = undefined
    else if (transform === 'capitalize') out.caps = undefined
    if (a['fo:font-variant'] === 'small-caps') out.caps = 'small'
    return Object.keys(out).length > 0 ? out : undefined
  }

  private readParaProps(node: XNode): OdfParaProps | undefined {
    const a = nodeAttrs(node)
    const out: OdfParaProps = {}
    const align = a['fo:text-align']
    if (align === 'start' || align === 'left') out.align = 'left'
    else if (align === 'center') out.align = 'center'
    else if (align === 'end' || align === 'right') out.align = 'right'
    else if (align === 'justify') out.align = 'justify'
    const before = odfLengthToEmu(a['fo:margin-top'])
    if (before !== null) out.spaceBeforeTw = emuToTwips(before)
    const after = odfLengthToEmu(a['fo:margin-bottom'])
    if (after !== null) out.spaceAfterTw = emuToTwips(after)
    const lineHeight = a['fo:line-height']
    if (lineHeight) {
      const percent = /^(\d+(?:\.\d+)?)%$/.exec(lineHeight.trim())
      if (percent) {
        out.lineSpacing = Number(percent[1]) / 100
      } else {
        const fixed = odfLengthToEmu(lineHeight)
        if (fixed !== null) out.lineAtLeastTw = emuToTwips(fixed)
      }
    }
    const atLeast = odfLengthToEmu(a['style:line-height-at-least'])
    if (atLeast !== null) out.lineAtLeastTw = emuToTwips(atLeast)
    const exact = odfLengthToEmu(a['style:line-spacing'])
    if (exact !== null) out.lineExactTw = emuToTwips(exact)
    const left = odfLengthToEmu(a['fo:margin-left'])
    if (left !== null) out.indentLeftTw = emuToTwips(left)
    const right = odfLengthToEmu(a['fo:margin-right'])
    if (right !== null) out.indentRightTw = emuToTwips(right)
    const first = odfLengthToEmu(a['fo:text-indent'])
    if (first !== null) out.indentFirstLineTw = emuToTwips(first)
    if (a['fo:break-before'] === 'page') out.pageBreakBefore = true
    return Object.keys(out).length > 0 ? out : undefined
  }

  private addListStyle(node: XNode): void {
    const name = attr(node, 'style:name')
    if (!name) return
    const levels = new Map<number, OdfListLevel>()
    for (const child of nodeChildren(node)) {
      const childName = nodeName(child)
      const level = Number(attr(child, 'text:level'))
      if (!Number.isInteger(level) || level < 1 || level > 10) continue
      if (childName === 'text:list-level-style-bullet') {
        levels.set(level, {
          kind: 'bullet',
          bulletChar: attr(child, 'style:bullet-char') ?? '•',
          ...listLevelIndent(child),
        })
      } else if (
        childName === 'text:list-level-style-number' ||
        childName === 'text:list-level-style-roman'
      ) {
        levels.set(level, {
          kind: 'ordered',
          numFormat: attr(child, 'style:num-format') ?? '1',
          numPrefix: attr(child, 'style:num-prefix'),
          numSuffix: attr(child, 'style:num-suffix'),
          ...listLevelIndent(child),
        })
      }
    }
    // lists defined without level entries still name a style; give them the
    // default bullet shape so the structure survives
    this.lists.set(name, levels)
  }

  /** text/paragraph props for a style name, resolved along the parent chain. */
  textProps(styleName: string | undefined, families: readonly string[]): OdfTextProps | null {
    return this.resolve(styleName, families).text
  }

  paraProps(styleName: string | undefined, families: readonly string[]): OdfParaProps | null {
    return this.resolve(styleName, families).para
  }

  private resolve(
    styleName: string | undefined,
    families: readonly string[],
  ): { text: OdfTextProps | null; para: OdfParaProps | null } {
    const chain: StyleDef[] = []
    const seen = new Set<string>()
    let current = styleName
    while (current && !seen.has(current)) {
      seen.add(current)
      const def = families.map((family) => this.defs.get(`${family}:${current}`)).find(Boolean)
      if (!def) break
      chain.push(def)
      current = def.parent
    }
    let text: OdfTextProps = { ...(this.defaults.text ?? {}) }
    let para: OdfParaProps = { ...(this.defaults.para ?? {}) }
    // parent first so the style itself wins
    for (let index = chain.length - 1; index >= 0; index--) {
      if (chain[index].text) text = { ...text, ...chain[index].text }
      if (chain[index].para) para = { ...para, ...chain[index].para }
    }
    return {
      text: Object.keys(text).length > 0 ? text : null,
      para: Object.keys(para).length > 0 ? para : null,
    }
  }

  /** Levels of a `text:list-style`; null when the style is unknown. */
  listLevels(styleName: string | undefined): Map<number, OdfListLevel> | null {
    if (!styleName) return null
    return this.lists.get(styleName) ?? null
  }

  /** Page geometry of the first master page: width/height/margins in twips. */
  pageSetup(stylesRoots: XNode[] | null): {
    widthTw: number
    heightTw: number
    marginTopTw: number
    marginRightTw: number
    marginBottomTw: number
    marginLeftTw: number
  } | null {
    const properties = this.pageLayoutProperties(stylesRoots)
    if (!properties) return null
    const a = nodeAttrs(properties)
    const width = odfLengthToEmu(a['fo:page-width'])
    const height = odfLengthToEmu(a['fo:page-height'])
    if (width === null || height === null || width <= 0 || height <= 0) return null
    const margin = (key: string, fallback: string | undefined): number => {
      const value = odfLengthToEmu(a[key] ?? fallback)
      return value === null ? 0 : emuToTwips(Math.max(0, value))
    }
    return {
      widthTw: emuToTwips(width),
      heightTw: emuToTwips(height),
      marginTopTw: margin('fo:margin-top', a['fo:margin']),
      marginRightTw: margin('fo:margin-right', a['fo:margin']),
      marginBottomTw: margin('fo:margin-bottom', a['fo:margin']),
      marginLeftTw: margin('fo:margin-left', a['fo:margin']),
    }
  }

  /** Page size of the first master page in EMU — the deck's slide size. */
  pageSizeEmu(stylesRoots: XNode[] | null): { widthEmu: number; heightEmu: number } | null {
    const properties = this.pageLayoutProperties(stylesRoots)
    if (!properties) return null
    const a = nodeAttrs(properties)
    const width = odfLengthToEmu(a['fo:page-width'])
    const height = odfLengthToEmu(a['fo:page-height'])
    if (width === null || height === null || width <= 0 || height <= 0) return null
    return { widthEmu: width, heightEmu: height }
  }

  private pageLayoutProperties(stylesRoots: XNode[] | null): XNode | null {
    if (!stylesRoots) return null
    let layoutName: string | null = null
    const walk = (node: XNode): void => {
      if (layoutName) return
      if (nodeName(node) === 'style:master-page') {
        layoutName = attr(node, 'style:page-layout-name') ?? null
        return
      }
      for (const child of nodeChildren(node)) walk(child)
    }
    for (const root of stylesRoots) walk(root)
    if (!layoutName) return null
    let properties: XNode | null = null
    const findLayout = (node: XNode): void => {
      if (properties) return
      if (nodeName(node) === 'style:page-layout' && attr(node, 'style:name') === layoutName) {
        properties = firstChild(node, 'style:page-layout-properties')
        return
      }
      for (const child of nodeChildren(node)) findLayout(child)
    }
    for (const root of stylesRoots) findLayout(root)
    return properties
  }
}

/** label-alignment indents of a list level (space-before + min-label-width). */
function listLevelIndent(node: XNode): { indentLeftTw?: number; hangingTw?: number } {
  const properties = firstChild(node, 'style:list-level-properties')
  if (!properties) return {}
  const a = nodeAttrs(properties)
  const spaceBefore = odfLengthToEmu(a['text:space-before'])
  const minWidth = odfLengthToEmu(a['text:min-label-width'])
  const indent = a['fo:text-indent']
  const out: { indentLeftTw?: number; hangingTw?: number } = {}
  if (spaceBefore !== null || minWidth !== null) {
    const left = (spaceBefore ?? 0) + (minWidth ?? 0)
    out.indentLeftTw = emuToTwips(left)
  }
  if (minWidth !== null) out.hangingTw = emuToTwips(minWidth)
  if (indent !== undefined) {
    const value = odfLengthToEmu(indent)
    // ODF expresses the hanging as a negative first-line indent
    if (value !== null && value < 0) out.hangingTw = emuToTwips(-value)
  }
  return out
}
