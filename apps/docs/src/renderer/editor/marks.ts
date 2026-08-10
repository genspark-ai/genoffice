import { Extension, Mark } from '@tiptap/core'
import { Plugin, PluginKey } from '@tiptap/pm/state'
import { Decoration, DecorationSet } from '@tiptap/pm/view'
import {} from '@tiptap/pm/tables'
import { cssCsFontFamily, cssDualFontFamily, cssFontFamily } from '../line-metrics'
import { isEastAsianFontName } from '../font-list'
import { t } from '../i18n/locale'
import {} from '@genoffice/docx-engine'

/**
 * Custom schema mirroring the docx-engine Block model 1:1.
 * Every top-level node carries `docxIndex` (patch anchor, null = new) and
 * `aiChanged` (diff highlighting for AI edits).
 */

export const BoldMark = Mark.create({
  name: 'bold',
  parseHTML() {
    return [
      { tag: 'strong' },
      // Word-pasted <b style="font-weight:normal"> doesn't count as bold
      {
        tag: 'b',
        getAttrs: (el) => ((el as HTMLElement).style.fontWeight !== 'normal' ? null : false),
      },
      {
        style: 'font-weight',
        getAttrs: (value) =>
          value === 'bold' || value === 'bolder' || parseInt(String(value), 10) >= 600
            ? null
            : false,
      },
    ]
  },
  renderHTML() {
    return ['strong', 0]
  },
  addKeyboardShortcuts() {
    return { 'Mod-b': () => this.editor.commands.toggleMark('bold') }
  },
})

export const ItalicMark = Mark.create({
  name: 'italic',
  parseHTML() {
    return [{ tag: 'em' }, { tag: 'i' }, { style: 'font-style=italic' }]
  },
  renderHTML() {
    return ['em', 0]
  },
  addKeyboardShortcuts() {
    return { 'Mod-i': () => this.editor.commands.toggleMark('italic') }
  },
})

export const UnderlineMark = Mark.create({
  name: 'underline',
  parseHTML() {
    return [
      { tag: 'u' },
      {
        style: 'text-decoration',
        getAttrs: (value) => (String(value).includes('underline') ? null : false),
      },
    ]
  },
  renderHTML() {
    return ['u', 0]
  },
  addKeyboardShortcuts() {
    return { 'Mod-u': () => this.editor.commands.toggleMark('underline') }
  },
})

export const StrikeMark = Mark.create({
  name: 'strike',
  parseHTML() {
    return [
      { tag: 's' },
      { tag: 'del' },
      { tag: 'strike' },
      {
        style: 'text-decoration',
        getAttrs: (value) => (String(value).includes('line-through') ? null : false),
      },
    ]
  },
  renderHTML() {
    return ['s', 0]
  },
})

export const LinkMark = Mark.create({
  name: 'link',
  inclusive: false,
  addAttributes() {
    return {
      href: { default: '' },
      rId: { default: null as string | null },
      tooltip: { default: null as string | null },
    }
  },
  parseHTML() {
    return [
      {
        tag: 'a[href]',
        getAttrs: (el) => ({
          href: (el as HTMLElement).getAttribute('href') ?? '',
          tooltip: (el as HTMLElement).getAttribute('title'),
        }),
      },
    ]
  },
  renderHTML({ mark }) {
    return [
      'a',
      {
        href: mark.attrs.href,
        class: 'doc-link',
        ...(mark.attrs.tooltip ? { title: String(mark.attrs.tooltip) } : {}),
      },
      0,
    ]
  },
})

/**
 * Word comment range: rendered as a highlighted span carrying its comment ids
 * (space-separated), so the comments panel can jump to it. `inclusive: false`
 * keeps typing at the range edges from silently extending the comment.
 */
export const CommentMark = Mark.create({
  name: 'comment',
  inclusive: false,
  addAttributes() {
    return {
      ids: { default: '' },
    }
  },
  parseHTML() {
    return [{ tag: 'span[data-comment-ids]' }]
  },
  renderHTML({ mark }) {
    return ['span', { 'data-comment-ids': mark.attrs.ids, class: 'doc-comment' }, 0]
  },
})

/**
 * Tracked insertion (w:ins). Rendered underlined in the revision color; hover
 * shows author/date. `inclusive: false` so typing at the edge of someone
 * else's insertion doesn't inherit their authorship — the track-changes
 * recorder marks new input itself.
 */
export const InsMark = Mark.create({
  name: 'ins',
  inclusive: false,
  addAttributes() {
    return {
      author: { default: '' },
      date: { default: null as string | null },
      id: { default: null as string | null },
    }
  },
  parseHTML() {
    return [{ tag: 'span[data-ins-author]' }]
  },
  renderHTML({ mark }) {
    return [
      'span',
      {
        'data-ins-author': mark.attrs.author,
        class: 'doc-ins',
        title: `${t('editorInsertedBy', { author: String(mark.attrs.author) })}${mark.attrs.date ? ` · ${String(mark.attrs.date).slice(0, 10)}` : ''}`,
      },
      0,
    ]
  },
})

/** Tracked deletion (w:del). Rendered struck-through in the revision color. */
export const DelMark = Mark.create({
  name: 'del',
  inclusive: false,
  addAttributes() {
    return {
      author: { default: '' },
      date: { default: null as string | null },
      id: { default: null as string | null },
    }
  },
  parseHTML() {
    return [{ tag: 'span[data-del-author]' }]
  },
  renderHTML({ mark }) {
    return [
      'span',
      {
        'data-del-author': mark.attrs.author,
        class: 'doc-del',
        title: `${t('editorDeletedBy', { author: String(mark.attrs.author) })}${mark.attrs.date ? ` · ${String(mark.attrs.date).slice(0, 10)}` : ''}`,
      },
      0,
    ]
  },
})

/** OOXML named highlight color -> CSS color, for on-screen rendering */
export const HIGHLIGHT_CSS: Record<string, string> = {
  yellow: '#FFFF00',
  green: '#00FF00',
  cyan: '#00FFFF',
  magenta: '#FF00FF',
  blue: '#0000FF',
  red: '#FF0000',
  darkBlue: '#00008B',
  darkCyan: '#008B8B',
  darkGreen: '#006400',
  darkMagenta: '#8B008B',
  darkRed: '#8B0000',
  darkYellow: '#808000',
  darkGray: '#808080',
  lightGray: '#C0C0C0',
  black: '#000000',
  white: '#FFFFFF',
}

/** Cross-reference (REF field): gray background, title hints the target bookmark, text is the cached display result */
export const RefFieldMark = Mark.create({
  name: 'refField',
  addAttributes() {
    return { name: { default: '' } }
  },
  parseHTML() {
    return [{ tag: 'span[data-ref-field]' }]
  },
  renderHTML({ mark }) {
    return [
      'span',
      {
        'data-ref-field': String(mark.attrs.name),
        class: 'doc-ref-field',
        title: t('editorCrossReference', { name: String(mark.attrs.name) }),
      },
      0,
    ]
  },
})

/** Revision display mode (synced by App; in original mode the extension below restores old formatting via decorations) */
export const revisionDisplayState = { mode: 'all' as 'all' | 'none' | 'original' }

const revisionOriginalKey = new PluginKey('revisionOriginal')

/**
 * Original view (shown as if rejected): for text with rPrChange, restore the pre-revision
 * modeled formatting via inner decorations (mirroring revisions.ts reject logic). The inner
 * span can override bold/italic, color, font size, and font; undoing underline/strikethrough
 * isn't possible in CSS, so it stays approximate.
 */
export const RevisionOriginalExtension = Extension.create({
  name: 'revisionOriginal',
  addProseMirrorPlugins() {
    return [
      new Plugin({
        key: revisionOriginalKey,
        props: {
          decorations(state) {
            if (revisionDisplayState.mode !== 'original') return DecorationSet.empty
            const decos: Decoration[] = []
            state.doc.descendants((node, pos) => {
              if (!node.isText) return
              const rpr = node.marks.find((m) => m.type.name === 'rprChange')
              if (!rpr) return
              const old = (rpr.attrs.old ?? {}) as Record<string, unknown>
              const styles = [
                `font-weight:${old.bold ? 600 : 400}`,
                `font-style:${old.italic ? 'italic' : 'normal'}`,
              ]
              if (old.color) styles.push(`color:#${old.color}`)
              if (old.sizeHalfPoints) styles.push(`font-size:${Number(old.sizeHalfPoints) / 2}pt`)
              if (old.font || old.fontAscii) {
                const ea = old.font ? String(old.font) : null
                const ascii = old.fontAscii ? String(old.fontAscii) : null
                styles.push(
                  `font-family:${ea && ascii ? cssDualFontFamily(ascii, ea) : cssFontFamily((ea ?? ascii)!)}`,
                )
              }
              decos.push(Decoration.inline(pos, pos + node.nodeSize, { style: styles.join(';') }))
            })
            return decos.length > 0 ? DecorationSet.create(state.doc, decos) : DecorationSet.empty
          },
        },
      }),
    ]
  },
})

/** Run-level format revision (w:rPrChange): shown with an amber squiggle; accept/reject on the Review tab */
export const RprChangeMark = Mark.create({
  name: 'rprChange',
  addAttributes() {
    return {
      author: { default: '' },
      date: { default: null as string | null },
      id: { default: null as string | null },
      old: { default: null as Record<string, unknown> | null, rendered: false },
    }
  },
  parseHTML() {
    return [{ tag: 'span[data-rpr-change]' }]
  },
  renderHTML({ mark }) {
    return [
      'span',
      {
        'data-rpr-change': '1',
        class: 'has-rpr-change',
        title: t('editorFormatChangeBy', {
          author: String(mark.attrs.author || t('editorUnknownAuthor')),
        }),
      },
      0,
    ]
  },
})

/** Generic inline field (DATE/TIME/NUMPAGES/FILENAME…): text is the cached result, recomputed on F9 */
export const InstrFieldMark = Mark.create({
  name: 'instrField',
  addAttributes() {
    return { instr: { default: '' } }
  },
  parseHTML() {
    return [{ tag: 'span[data-instr-field]' }]
  },
  renderHTML({ mark }) {
    return [
      'span',
      {
        'data-instr-field': String(mark.attrs.instr),
        class: 'doc-ref-field',
        title: t('editorFieldHint', { instr: String(mark.attrs.instr) }),
      },
      0,
    ]
  },
})

/**
 * font-family chain → dual-slot font attrs, inverting renderHTML's encoding:
 * the Latin slot is the chain's first Latin family, the eastAsia slot its first
 * East Asian family. Generic families and the bundled CJK tofu-fallbacks that
 * cssFontFamily appends to Latin chains are not user picks and are skipped.
 * A Latin-only chain fills both slots, matching how parse.ts reads a
 * Latin-only w:rFonts. Exported for tests.
 */
export function fontAttrsFromFamilyChain(chain: string | undefined): Record<string, unknown> {
  const families = (chain ?? '')
    .split(',')
    .map((x) => x.trim().replace(/^["']|["']$/g, ''))
    .filter(
      (x) =>
        x &&
        !/^(serif|sans-serif|monospace|cursive|fantasy|system-ui)$/i.test(x) &&
        // internal fonts.css aliases (GenOffice Songti SC etc.) are not user picks
        !/^genoffice /i.test(x),
    )
  const ea = families.find(
    (f, i) => isEastAsianFontName(f) && (i === 0 || !/^noto (sans|serif) cjk sc$/i.test(f)),
  )
  const latin = families.find((f) => !isEastAsianFontName(f))
  if (!ea && !latin) return {}
  return { font: ea ?? latin, ...(latin ? { fontAscii: latin } : {}) }
}

/** Inline styles of foreign HTML → docTextStyle attrs (returns false when no usable style, so no mark is applied) */
function textStyleAttrsFromDom(el: HTMLElement): Record<string, unknown> | false {
  const attrs: Record<string, unknown> = {}
  const st = el.style
  const hex = (css: string): string | null => {
    if (!css) return null
    const rgb = /rgba?\((\d+)[,\s]+(\d+)[,\s]+(\d+)/.exec(css)
    if (rgb) {
      return [rgb[1], rgb[2], rgb[3]]
        .map((n) => Number(n).toString(16).padStart(2, '0'))
        .join('')
        .toUpperCase()
    }
    const m = /^#([0-9a-fA-F]{6})$/.exec(css.trim())
    return m ? m[1].toUpperCase() : null
  }
  const color = hex(st.color)
  if (color && color !== '000000') attrs.color = color
  const size = st.fontSize
  if (size) {
    const v = parseFloat(size)
    if (Number.isFinite(v) && v > 0) {
      if (size.endsWith('pt')) attrs.sizeHalfPoints = Math.round(v * 2)
      else if (size.endsWith('px')) attrs.sizeHalfPoints = Math.round(v * 1.5)
    }
  }
  Object.assign(attrs, fontAttrsFromFamilyChain(st.fontFamily))
  const spacing = parseFloat(st.letterSpacing)
  if (Number.isFinite(spacing) && spacing !== 0) {
    attrs.charSpacingTwips = Math.round(
      st.letterSpacing.endsWith('px') ? spacing * 15 : spacing * 20,
    )
  }
  if (st.verticalAlign === 'super') attrs.vertAlign = 'superscript'
  else if (st.verticalAlign === 'sub') attrs.vertAlign = 'subscript'
  const styleId = el.getAttribute('data-style')
  if (styleId) attrs.styleId = styleId
  return Object.keys(attrs).length > 0 ? attrs : false
}

/** color: hex without '#'; sizeHalfPoints: OOXML half-points; highlight: OOXML named color */
export const TextStyleMark = Mark.create({
  name: 'docTextStyle',
  addAttributes() {
    return {
      color: { default: null as string | null },
      sizeHalfPoints: { default: null as number | null },
      font: { default: null as string | null },
      // the EA face backfills an empty theme slot (line metrics follow the Latin face like LO)
      eaSlotEmpty: { default: null as boolean | null },
      // Latin slot (w:ascii/w:hAnsi) when it differs from the primary/eastAsia font
      fontAscii: { default: null as string | null },
      // complex-script slot (w:cs); convert sets it only when the run text needs it
      csFont: { default: null as string | null },
      charSpacingTwips: { default: null as number | null },
      // letter spacing (em, negative = condensed) converted from w:w scaling; precomputed by convert per run text
      charScaleEm: { default: null as number | null },
      highlight: { default: null as string | null },
      vertAlign: { default: null as 'superscript' | 'subscript' | null },
      // East Asian emphasis mark (w:em val); saving is kept faithful by rawRPr
      em: { default: null as string | null },
      styleId: { default: null as string | null },
      // raw rPr slice pass-through (not rendered; on save mergeRPrModel preserves unmodeled attributes)
      rawRPr: { default: null as string | null, rendered: false },
    }
  },
  parseHTML() {
    return [
      {
        tag: 'span[data-doc-style]',
        getAttrs: (el) => textStyleAttrsFromDom(el as HTMLElement) || {},
      },
      { tag: 'sup', attrs: { vertAlign: 'superscript' } },
      { tag: 'sub', attrs: { vertAlign: 'subscript' } },
      // inline-styled span from foreign paste (color/size/font); no mark applied when no usable style
      { tag: 'span', getAttrs: (el) => textStyleAttrsFromDom(el as HTMLElement) },
    ]
  },
  renderHTML({ mark }) {
    const styles: string[] = []
    if (mark.attrs.color) styles.push(`color:#${mark.attrs.color}`)
    if (mark.attrs.sizeHalfPoints)
      styles.push(`font-size:${Number(mark.attrs.sizeHalfPoints) / 2}pt`)
    if (mark.attrs.font || mark.attrs.fontAscii || mark.attrs.csFont) {
      const ea = mark.attrs.font ? String(mark.attrs.font) : null
      const ascii = mark.attrs.fontAscii ? String(mark.attrs.fontAscii) : null
      const cs = mark.attrs.csFont ? String(mark.attrs.csFont) : null
      styles.push(
        `font-family:${
          cs
            ? cssCsFontFamily(cs, ascii ?? undefined, ea ?? undefined)
            : ea && ascii
              ? cssDualFontFamily(ascii, ea)
              : cssFontFamily((ea ?? ascii)!)
        }`,
      )
    }
    const spacingPt = mark.attrs.charSpacingTwips ? Number(mark.attrs.charSpacingTwips) / 20 : 0
    const scaleEm = mark.attrs.charScaleEm ? Number(mark.attrs.charScaleEm) : 0
    if (spacingPt && scaleEm) styles.push(`letter-spacing:calc(${spacingPt}pt + ${scaleEm}em)`)
    else if (spacingPt) styles.push(`letter-spacing:${spacingPt}pt`)
    else if (scaleEm) styles.push(`letter-spacing:${scaleEm}em`)
    if (mark.attrs.highlight) {
      styles.push(
        `background-color:${HIGHLIGHT_CSS[mark.attrs.highlight as string] ?? mark.attrs.highlight}`,
      )
    }
    if (mark.attrs.vertAlign === 'superscript') styles.push('vertical-align:super;font-size:0.75em')
    if (mark.attrs.vertAlign === 'subscript') styles.push('vertical-align:sub;font-size:0.75em')
    if (mark.attrs.em) {
      const em = String(mark.attrs.em)
      const shape =
        em === 'circle' ? 'open circle' : em === 'comma' ? 'filled sesame' : 'filled dot'
      // Word renders Chinese emphasis marks (dot/underDot) below the text; comma/circle kenten go above
      const pos = em === 'comma' || em === 'circle' ? 'over' : 'under'
      styles.push(`text-emphasis:${shape}`, `text-emphasis-position:${pos} right`)
    }
    const attrs: Record<string, string> = { 'data-doc-style': '1', style: styles.join(';') }
    if (mark.attrs.styleId) attrs['data-style'] = String(mark.attrs.styleId)
    return ['span', attrs, 0]
  },
})

// ---- textbox sub-editor schema ----
