import { Editor, Extension, Node } from '@tiptap/core'
import type { ChainedCommands, RawCommands } from '@tiptap/core'
import { UndoRedo } from '@tiptap/extensions'
import { DOMSerializer } from '@tiptap/pm/model'
import type { Node as PmNode } from '@tiptap/pm/model'
import {
  NodeSelection,
  Plugin,
  PluginKey,
  type EditorState,
  type Transaction,
} from '@tiptap/pm/state'
import { Decoration, DecorationSet } from '@tiptap/pm/view'
import type { EditorView } from '@tiptap/pm/view'
import {
  CellSelection,
  addRowAfter,
  columnResizing,
  columnResizingPluginKey,
  deleteTable,
  goToNextCell,
  isInTable,
  tableEditing,
} from '@tiptap/pm/tables'
import { cssLineHeight, textHasCjk } from '../line-metrics'
import { noteMarkText } from '../note-format'
import { t } from '../i18n/locale'
import {
  ommlToMathML,
  patchMathTokens,
  type ChartDisplay,
  type FieldDisplay,
  type FormulaDisplay,
  type NewChart,
  type NumberingDef,
  type TableCell,
  type TableModel,
  type TextboxDisplay,
} from '@genoffice/docx-engine'
import { computeListMarkers, type ListItemRef } from './numbering'
import { dropActiveSubEditor, notifySubEditorState, setActiveSubEditor } from './active-editor'
import { PaginationGapsExtension } from './pagination-gaps'
import { TrackChangesExtension } from './revisions'
import { inlineToRuns, runsToInline, textboxParaSignature, type PmNode as PmJson } from './convert'
import { constrainTableWidthAtCell } from './table-sizing'

/**
 * Custom schema mirroring the docx-engine Block model 1:1.
 * Every top-level node carries `docxIndex` (patch anchor, null = new) and
 * `aiChanged` (diff highlighting for AI edits).
 */

import {
  CHART_MAX_WIDTH_PX,
  drawChartSvg,
  renderChartSpec,
  renderFieldSpec,
  renderFormulaSpec,
  renderTableSpec,
  renderTextboxSpec,
  textboxBoxStyle,
  wireChartEditing,
} from './protected-render'
import { isStraightLineKind } from './shape-svg'
import {
  BoldMark,
  CommentMark,
  DelMark,
  InsMark,
  InstrFieldMark,
  ItalicMark,
  LinkMark,
  RefFieldMark,
  RevisionOriginalExtension,
  RprChangeMark,
  StrikeMark,
  TextStyleMark,
  UnderlineMark,
} from './marks'
import {
  DropCapExtension,
  MoveRevisionExtension,
  PPrChangeExtension,
  PendingCommentHighlightExtension,
  ResolvedCommentsExtension,
  SdtExtension,
  SearchHighlightExtension,
  TabStopExtension,
} from './decoration-extensions'
import { AutoDirectionExtension } from './direction'
export * from './marks'
export * from './decoration-extensions'

const anchorAttrs = {
  docxIndex: { default: null as number | null },
  styleId: { default: null as string | null },
  aiChanged: { default: false },
  /** user bookmark names starting in this paragraph */
  bookmarks: { default: null as string[] | null },
  /** Word internal bookmarks (_Ref/_Toc…): kept out of the bookmark manager, written back verbatim on paragraph rebuild so cross-references don't break */
  hiddenBookmarks: { default: null as string[] | null },
  /** Endpoints of cross-paragraph comment ranges (only one end in this paragraph); written back on paragraph rebuild to avoid orphan marks */
  commentStarts: { default: null as string[] | null },
  commentEnds: { default: null as string[] | null },
  align: { default: null as string | null },
  lineSpacing: { default: null as number | null },
  /** Word line-spacing rule (auto/atLeast/exact); null = auto */
  lineRule: { default: null as string | null },
  /** Raw w:spacing w:line twips (used by atLeast/exact) */
  lineRawTwips: { default: null as number | null },
  indentLeft: { default: null as number | null },
  indentRight: { default: null as number | null },
  indentFirstLine: { default: null as number | null },
  spaceBefore: { default: null as number | null },
  spaceAfter: { default: null as number | null },
  pageBreakBefore: { default: false },
  /** RTL paragraph (w:bidi); align is already the visual value */
  bidi: { default: false },
  shadingFill: { default: null as string | null },
  /** subset of "tblr": which sides have a single-line border */
  borders: { default: null as string | null },
  /** custom tab stops JSON: Array<{pos:number,val:string,leader?:string}> */
  tabStops: { default: null as string | null },
  /** drop cap: JSON {type:'drop'|'margin',lines:number} */
  dropCap: { default: null as string | null },
  /** SDT shell: JSON SdtShell (alias, tag, controlType, openXml, closeXml) */
  sdtShell: { default: null as string | null },
  /** move revision type: 'from' (content moved away) or 'to' (content moved here) */
  moveRevision: { default: null as string | null },
  /** JSON {author,date?,id?} when the paragraph has a pPrChange tracked format change */
  pPrChange: { default: null as string | null },
  /** top-level insertion/deletion revision ({kind,author,date?}) */
  blockRevision: { default: null as Record<string, string> | null },
}

function blockAttrs(
  node: { attrs: Record<string, unknown>; textContent?: string },
  {
    includeIndent = true,
    listGeometry = false,
  }: { includeIndent?: boolean; listGeometry?: boolean } = {},
): Record<string, string> {
  const attrs: Record<string, string> = {}
  if (node.attrs.docxIndex !== null) attrs['data-idx'] = String(node.attrs.docxIndex)
  // per-document style CSS (generated from styles.xml) targets this attribute
  if (node.attrs.styleId) attrs['data-style'] = String(node.attrs.styleId)
  // bookmark jump targets ([data-bookmarks~="name"]; names cannot contain spaces)
  if (Array.isArray(node.attrs.bookmarks) && node.attrs.bookmarks.length > 0) {
    attrs['data-bookmarks'] = (node.attrs.bookmarks as string[]).join(' ')
  }
  const classes: string[] = []
  if (node.attrs.aiChanged) classes.push('ai-changed')
  if (node.attrs.pageBreakBefore) {
    classes.push('page-break-before')
    attrs['data-page-break-label'] = t('editorPageBreak')
  }
  if (classes.length > 0) attrs['class'] = classes.join(' ')
  const styles: string[] = []
  if (node.attrs.bidi) styles.push('direction:rtl', 'unicode-bidi:isolate')
  if (node.attrs.align) {
    styles.push(`text-align:${node.attrs.align === 'distribute' ? 'justify' : node.attrs.align}`)
  }
  // the line-height factor follows paragraph content (approximating Word's max-of-inline-fonts
  // line height): CJK paragraphs get the CJK factor, pure-Western ones the document's
  // font-aware Latin factor (doc-style-css sets --doc-line-factor-latin per body font)
  if (node.textContent) {
    styles.push(
      `--doc-line-factor:${
        textHasCjk(node.textContent) ? 1.3 : 'var(--doc-line-factor-latin,1.2)'
      }`,
    )
  }
  const lh = cssLineHeight(
    (node.attrs.lineRule as 'auto' | 'atLeast' | 'exact' | null) ?? undefined,
    node.attrs.lineRawTwips ? Number(node.attrs.lineRawTwips) : undefined,
    node.attrs.lineSpacing ? Number(node.attrs.lineSpacing) : undefined,
  )
  if (lh) styles.push(`line-height:${lh}`)
  // list items already indent via padding; a margin would double-shift them.
  // listGeometry: drive list geometry with w:ind (--li-left text indent, --li-hang the hanging
  // area i.e. the number-marker width); negative text-indent is expressed by the marker box, no longer emitted directly
  // Logical (inline-start/end) margins: identical to left/right in LTR, and in bidi
  // paragraphs they mirror, matching Word's quirk that w:ind left/right swap sides.
  if (includeIndent && node.attrs.indentLeft) {
    styles.push(`margin-inline-start:${Number(node.attrs.indentLeft) / 20}pt`)
  } else if (listGeometry && node.attrs.indentLeft) {
    styles.push(`--li-left:${Number(node.attrs.indentLeft) / 20}pt`)
  }
  if (node.attrs.indentRight)
    styles.push(`margin-inline-end:${Number(node.attrs.indentRight) / 20}pt`)
  if (node.attrs.indentFirstLine) {
    const firstLine = Number(node.attrs.indentFirstLine)
    if (listGeometry && firstLine < 0) styles.push(`--li-hang:${-firstLine / 20}pt`)
    else styles.push(`text-indent:${firstLine / 20}pt`)
  }
  // explicit 0 must still emit (w:after="0" overrides the style/docDefaults margin)
  if (node.attrs.spaceBefore != null)
    styles.push(`margin-top:${Number(node.attrs.spaceBefore) / 20}pt`)
  if (node.attrs.spaceAfter != null)
    styles.push(`margin-bottom:${Number(node.attrs.spaceAfter) / 20}pt`)
  if (node.attrs.shadingFill) styles.push(`background-color:#${node.attrs.shadingFill}`)
  if (node.attrs.borders) {
    const borders = String(node.attrs.borders)
    const line = '1px solid #444'
    if (borders.includes('t')) styles.push(`border-top:${line}`)
    if (borders.includes('b')) styles.push(`border-bottom:${line}`)
    if (borders.includes('l')) styles.push(`border-left:${line}`)
    if (borders.includes('r')) styles.push(`border-right:${line}`)
    styles.push('padding:1px 4px')
  }
  if (node.attrs.tabStops) {
    // debugging aid only; actual tab layout is measured by TabStopExtension
    attrs['data-tab-stops'] = String(node.attrs.tabStops)
  }
  if (node.attrs.dropCap) {
    attrs['data-drop-cap'] = String(node.attrs.dropCap)
  }
  if (node.attrs.sdtShell) {
    attrs['data-sdt'] = 'true'
  }
  if (node.attrs.moveRevision) {
    attrs['data-move-revision'] = String(node.attrs.moveRevision)
  }
  if (node.attrs.pPrChange) {
    attrs['data-ppr-change'] = 'true'
  }
  if (styles.length > 0) attrs['style'] = styles.join(';')
  return attrs
}

export const DocDocument = Node.create({
  name: 'doc',
  topNode: true,
  content: 'block+',
})

export const DocText = Node.create({
  name: 'text',
  group: 'inline',
})

/** Footnote / endnote reference marker: an atomic superscript number. */
export const DocNoteRef = Node.create({
  name: 'docNoteRef',
  inline: true,
  group: 'inline',
  atom: true,
  selectable: true,
  addAttributes() {
    return {
      kind: { default: 'footnote' as 'footnote' | 'endnote' },
      id: { default: '' },
      num: { default: 1 },
    }
  },
  parseHTML() {
    return [{ tag: 'sup[data-note-ref]' }]
  },
  renderHTML({ node }) {
    return [
      'sup',
      {
        'data-note-ref': String(node.attrs.id),
        'data-note-kind': String(node.attrs.kind),
        class: 'doc-note-ref',
        title: node.attrs.kind === 'footnote' ? t('editorFootnote') : t('editorEndnote'),
      },
      `[${noteMarkText(node.attrs.kind as 'footnote' | 'endnote', Number(node.attrs.num) || 1)}]`,
    ]
  },
})

/** Index entry (XE field) marker: invisible in Word, a small chip on screen. */
export const DocXeMark = Node.create({
  name: 'docXeMark',
  inline: true,
  group: 'inline',
  atom: true,
  selectable: true,
  addAttributes() {
    return { term: { default: '' } }
  },
  parseHTML() {
    return [{ tag: 'span[data-xe-term]' }]
  },
  renderHTML({ node }) {
    return [
      'span',
      {
        'data-xe-term': String(node.attrs.term),
        class: 'doc-xe-mark',
        title: t('editorIndexEntry', { term: String(node.attrs.term) }),
      },
    ]
  },
})

/**
 * Phonetic guide (w:ruby): atomic, renders as a native <ruby> element.
 * `xml` is the exact <w:ruby> fragment that saves verbatim.
 */
export const DocRuby = Node.create({
  name: 'docRuby',
  inline: true,
  group: 'inline',
  atom: true,
  selectable: true,
  addAttributes() {
    return {
      base: { default: '' },
      rt: { default: '' },
      xml: { default: '' },
    }
  },
  parseHTML() {
    return [{ tag: 'ruby[data-doc-ruby]' }]
  },
  renderText({ node }) {
    return String(node.attrs.base ?? '')
  },
  renderHTML({ node }) {
    return [
      'ruby',
      { 'data-doc-ruby': 'true', class: 'doc-ruby' },
      String(node.attrs.base),
      ['rt', {}, String(node.attrs.rt)],
    ]
  },
})

/**
 * Atomic inline formula flowing with the text. `omml` is the exact <m:oMath>
 * fragment that saves verbatim; `mathml` renders natively in Chromium;
 * `latex` is kept for editor-created formulas so double-click can re-edit.
 */
export const DocInlineMath = Node.create({
  name: 'docInlineMath',
  inline: true,
  group: 'inline',
  atom: true,
  selectable: true,
  addAttributes() {
    return {
      omml: { default: '' },
      mathml: { default: '' },
      latex: { default: null as string | null },
      /** flat token strip (word count / AI read fallback) */
      text: { default: '' },
    }
  },
  parseHTML() {
    return [{ tag: 'span[data-inline-math]' }]
  },
  renderText({ node }) {
    return String(node.attrs.text ?? '')
  },
  renderHTML({ node }) {
    return ['span', inlineMathDomAttrs(node), String(node.attrs.text ?? '')]
  },
  addNodeView() {
    return ({ node, getPos }) => {
      let currentNode = node
      const dom = document.createElement('span')
      const render = () => {
        for (const [key, value] of Object.entries(inlineMathDomAttrs(currentNode))) {
          dom.setAttribute(key, value)
        }
        const mathml = String(currentNode.attrs.mathml ?? '')
        if (mathml) dom.innerHTML = mathml
        else dom.textContent = String(currentNode.attrs.text ?? '')
      }
      render()
      dom.addEventListener('dblclick', () => {
        // only editor-created formulas carry LaTeX and can be re-edited
        const latex = currentNode.attrs.latex
        const pos = (getPos as () => number | undefined)()
        if (latex && typeof pos === 'number') {
          window.dispatchEvent(
            new CustomEvent('ai-docs-edit-inline-math', {
              detail: { pos, latex: String(latex), kind: 'inline' },
            }),
          )
        }
      })
      return {
        dom,
        update: (n: PmNode) => {
          if (n.type.name !== 'docInlineMath') return false
          currentNode = n
          render()
          return true
        },
      }
    }
  },
})

function inlineMathDomAttrs(node: { attrs: Record<string, unknown> }): Record<string, string> {
  return {
    'data-inline-math': '1',
    class: 'doc-inline-math',
    title: node.attrs.latex
      ? t('editorEquationEditHint', { latex: String(node.attrs.latex) })
      : t('editorEquation'),
  }
}

export const DocHardBreak = Node.create({
  name: 'hardBreak',
  inline: true,
  group: 'inline',
  selectable: false,
  linebreakReplacement: true,
  addAttributes() {
    return {
      // in-paragraph page break (w:br w:type="page"): the pagination engine breaks after the containing block
      pageBreak: { default: false },
    }
  },
  parseHTML() {
    return [{ tag: 'br.doc-page-br', attrs: { pageBreak: true } }, { tag: 'br' }]
  },
  renderHTML({ node }) {
    return node.attrs.pageBreak ? ['br', { class: 'doc-page-br' }] : ['br']
  },
  addKeyboardShortcuts() {
    return {
      'Shift-Enter': () => this.editor.commands.insertContent({ type: 'hardBreak' }),
    }
  },
})

export const DocParagraph = Node.create({
  name: 'docParagraph',
  group: 'block',
  content: 'inline*',
  addAttributes() {
    return { ...anchorAttrs }
  },
  parseHTML() {
    return [{ tag: 'p' }]
  },
  renderHTML({ node }) {
    return ['p', blockAttrs(node), 0]
  },
})

export const DocHeading = Node.create({
  name: 'docHeading',
  group: 'block',
  content: 'inline*',
  addAttributes() {
    return { ...anchorAttrs, level: { default: 1 } }
  },
  parseHTML() {
    return [1, 2, 3, 4, 5, 6].map((level) => ({ tag: `h${level}`, attrs: { level } }))
  },
  renderHTML({ node }) {
    const level = Math.min(Math.max(Number(node.attrs.level) || 1, 1), 6)
    return [`h${level}`, blockAttrs(node), 0]
  },
})

export const DocListItem = Node.create({
  name: 'docListItem',
  group: 'block',
  content: 'inline*',
  addAttributes() {
    return {
      ...anchorAttrs,
      kind: { default: 'bullet' as 'bullet' | 'ordered' },
      numId: { default: null as string | null },
      ilvl: { default: 0 },
    }
  },
  parseHTML() {
    const ilvlOf = (el: HTMLElement): number => {
      let depth = -1
      for (let node = el.parentElement; node; node = node.parentElement) {
        if (node.tagName === 'UL' || node.tagName === 'OL') depth++
      }
      return Math.max(0, Math.min(depth, 8))
    }
    return [
      {
        tag: 'li',
        getAttrs: (el) => ({
          kind: (el as HTMLElement).closest('ol') ? 'ordered' : 'bullet',
          ilvl: ilvlOf(el as HTMLElement),
        }),
      },
    ]
  },
  renderHTML({ node }) {
    const base = blockAttrs(node, { includeIndent: false, listGeometry: true })
    const cls = [
      'doc-li',
      `doc-li-${node.attrs.kind}`,
      `ilvl-${Math.min(Number(node.attrs.ilvl) || 0, 4)}`,
      base['class'] ?? '',
    ]
      .filter(Boolean)
      .join(' ')
    return ['div', { ...base, class: cls }, 0]
  },
  addCommands() {
    return {
      /**
       * Word's Enter behavior inside a list: a non-empty item splits
       * into a sibling item (same kind/numId/level, numbering follows), an empty
       * one leaves the list. ProseMirror's default splitBlock produces the
       * schema's default block — a paragraph — which broke continuous entry.
       */
      continueDocList:
        () =>
        ({ state, chain }: { state: EditorState; chain: () => ChainedCommands }) => {
          const { $from, empty } = state.selection
          if (!empty) return false
          const node = $from.parent
          if (node.type.name !== 'docListItem') return false
          if (node.content.size === 0) {
            return chain()
              .setNode('docParagraph', { ...node.attrs, docxIndex: null })
              .run()
          }
          return chain()
            .splitBlock()
            .command(({ tr, dispatch }) => {
              const pos = tr.selection.$from.before(tr.selection.$from.depth)
              const created = tr.doc.nodeAt(pos)
              if (!created) return false
              // The split half is a fresh paragraph: turn it back into a sibling
              // list item, without inheriting the original's docx anchor
              if (dispatch) {
                tr.setNodeMarkup(pos, state.schema.nodes.docListItem, {
                  ...created.attrs,
                  docxIndex: null,
                  kind: node.attrs.kind,
                  numId: node.attrs.numId,
                  ilvl: node.attrs.ilvl,
                })
              }
              return true
            })
            .run()
        },
    } as Partial<RawCommands>
  },
  addKeyboardShortcuts() {
    const changeLevel = (delta: number) => () => {
      if (!this.editor.isActive('docListItem')) return false
      const ilvl = Number(this.editor.getAttributes('docListItem').ilvl) || 0
      const next = Math.min(Math.max(ilvl + delta, 0), 8)
      if (next === ilvl) return true
      return this.editor.commands.updateAttributes('docListItem', { ilvl: next })
    }
    return {
      Tab: changeLevel(1),
      'Shift-Tab': changeLevel(-1),
      Enter: () => (this.editor.commands as unknown as DocListCommands).continueDocList(),
    }
  },
})

interface DocListCommands {
  continueDocList: () => boolean
}

export interface ListNumberingStorage {
  /** numId -> definition, from the open document's numbering.xml */
  defs: Map<string, NumberingDef>
}

declare module '@tiptap/core' {
  interface Storage {
    listNumbering: ListNumberingStorage
  }
}

/**
 * Real multilevel numbering: compute each list item's marker in document order per the
 * numbering.xml definitions (1. / a. / 1.1 / Chinese numerals, …), attach a data-marker
 * attribute via node decoration, and display with CSS.
 * defs are written into storage by App when a document is opened/re-parsed; items without a definition fall back to CSS counters.
 */
// ── Live line-height factor ────────────────────────────────
// blockAttrs bakes --doc-line-factor into toDOM output, but ProseMirror reuses a
// block's DOM while typing (sameMarkup ignores content), so a block typed after
// creation keeps its creation-time factor until save/reopen. These node
// decorations recompute the factor from the live text on every doc change;
// unchanged nodes are structurally shared, so the WeakMap makes a pass cheap.

const LINE_FACTOR_BLOCKS = new Set(['docParagraph', 'docHeading', 'docListItem'])
const lineFactorCjkCache = new WeakMap<PmNode, boolean>()

function lineFactorDecos(doc: PmNode): DecorationSet {
  const decos: Decoration[] = []
  doc.descendants((node, pos) => {
    if (!LINE_FACTOR_BLOCKS.has(node.type.name)) return true
    if (node.textContent) {
      let cjk = lineFactorCjkCache.get(node)
      if (cjk === undefined) {
        cjk = textHasCjk(node.textContent)
        lineFactorCjkCache.set(node, cjk)
      }
      decos.push(
        Decoration.node(pos, pos + node.nodeSize, {
          style: `--doc-line-factor:${cjk ? 1.3 : 'var(--doc-line-factor-latin,1.2)'}`,
        }),
      )
    }
    return false
  })
  return DecorationSet.create(doc, decos)
}

export const LineFactorExtension = Extension.create({
  name: 'lineFactorLive',
  addProseMirrorPlugins() {
    const key = new PluginKey<DecorationSet>('lineFactorLive')
    return [
      new Plugin<DecorationSet>({
        key,
        state: {
          init: (_config, state) => lineFactorDecos(state.doc),
          apply: (tr, old) => (tr.docChanged ? lineFactorDecos(tr.doc) : old),
        },
        props: {
          decorations(state) {
            return key.getState(state)
          },
        },
      }),
    ]
  },
})

export const ListNumberingExtension = Extension.create<object, ListNumberingStorage>({
  name: 'listNumbering',
  addStorage() {
    return { defs: new Map<string, NumberingDef>() }
  },
  addProseMirrorPlugins() {
    const storage = this.storage
    const compute = (doc: PmNode): DecorationSet | null => {
      if (storage.defs.size === 0) return null
      const refs: ListItemRef[] = []
      const nodes: Array<{ pos: number; size: number; attrs: Record<string, unknown> }> = []
      doc.descendants((node, pos) => {
        if (node.type.name === 'docListItem') {
          refs.push({
            numId: (node.attrs.numId as string | null) ?? null,
            ilvl: Number(node.attrs.ilvl) || 0,
          })
          nodes.push({ pos, size: node.nodeSize, attrs: node.attrs })
          return false
        }
        return true
      })
      if (refs.length === 0) return null
      const markers = computeListMarkers(refs, storage.defs)
      const decos: Decoration[] = []
      markers.forEach((marker, i) => {
        if (marker === null) return
        const attrs: Record<string, string> = { 'data-marker': marker }
        // geometry fallback: when the paragraph has no w:ind of its own, use the numbering.xml level's indent;
        // marker font size always comes from the level's rPr (independent of paragraph text size)
        const def = refs[i].numId !== null ? storage.defs.get(refs[i].numId as string) : undefined
        const level = def?.levels[Math.max(0, refs[i].ilvl)]
        if (level) {
          const styles: string[] = []
          const nodeAttrs = nodes[i].attrs
          if (!nodeAttrs.indentLeft && level.indentLeft) {
            styles.push(`--li-left:${level.indentLeft / 20}pt`)
          }
          if (!nodeAttrs.indentFirstLine && level.hanging) {
            styles.push(`--li-hang:${level.hanging / 20}pt`)
          }
          if (level.szHalfPoints) styles.push(`--li-marker-size:${level.szHalfPoints / 2}pt`)
          if (styles.length > 0) attrs.style = styles.join(';')
        }
        decos.push(Decoration.node(nodes[i].pos, nodes[i].pos + nodes[i].size, attrs))
      })
      return DecorationSet.create(doc, decos)
    }

    interface CachedMarkers {
      defs: Map<string, NumberingDef>
      decos: DecorationSet | null
    }
    const key = new PluginKey<CachedMarkers>('listNumbering')
    return [
      new Plugin<CachedMarkers>({
        key,
        state: {
          init: (_config, state) => ({ defs: storage.defs, decos: compute(state.doc) }),
          apply(tr, old) {
            // defs is replaced (never mutated) on open/reparse and marker overlay
            if (tr.docChanged || old.defs !== storage.defs)
              return { defs: storage.defs, decos: compute(tr.doc) }
            return old.decos ? { defs: old.defs, decos: old.decos.map(tr.mapping, tr.doc) } : old
          },
        },
        props: {
          decorations(state) {
            return key.getState(state)?.decos ?? null
          },
        },
      }),
    ]
  },
})

const tableCellAttrs = {
  vAlign: { default: null as string | null },
  borders: { default: null as Record<string, unknown> | null },
  rawTcPr: { default: null as string | null },
  /** tcPr w:cellIns/w:cellDel cell revision ({kind, author, ...} | null) */
  cellRevision: { default: null as Record<string, string> | null },
  cellMar: { default: null as Record<string, number> | null },
  textDirection: { default: null as string | null },
  colspan: { default: 1 },
  rowspan: { default: 1 },
  colwidth: { default: null as number[] | null },
  fill: { default: null as string | null },
  color: { default: null as string | null },
  bold: { default: false },
  align: { default: null as string | null },
}

/** One OOXML border → CSS border value; 'none' means explicitly borderless */
export function borderLineCss(
  b: { style: string; szEighths?: number; color?: string } | undefined | null,
): string | null {
  if (!b) return null
  if (b.style === 'none' || b.style === 'nil') return 'none'
  const px = Math.max(1, Math.round(((b.szEighths ?? 4) / 8 / 72) * 96))
  const dash =
    b.style === 'dashed'
      ? 'dashed'
      : b.style === 'dotted'
        ? 'dotted'
        : b.style === 'double'
          ? 'double'
          : 'solid'
  const color = b.color && b.color !== 'auto' ? `#${b.color}` : '#000'
  return `${px}px ${dash} ${color}`
}

function tableCellHtml(node: PmNode): Record<string, string> {
  const attrs: Record<string, string> = {}
  if (node.attrs.colspan > 1) attrs.colspan = String(node.attrs.colspan)
  if (node.attrs.rowspan > 1) attrs.rowspan = String(node.attrs.rowspan)
  if (node.attrs.colwidth) attrs['data-colwidth'] = (node.attrs.colwidth as number[]).join(',')
  const borderCss = (side: string): string => {
    const v = borderLineCss(
      (
        node.attrs.borders as Record<
          string,
          { style: string; szEighths?: number; color?: string }
        > | null
      )?.[side],
    )
    return v ? `border-${side}:${v}` : ''
  }
  const mar = node.attrs.cellMar as Record<string, number> | null
  const styles = [
    // vertical-text cells: tbRl = vertical right-to-left (vertical-rl), btLr = rotated 90° counterclockwise (sideways-lr)
    node.attrs.textDirection === 'tbRl'
      ? 'writing-mode:vertical-rl'
      : node.attrs.textDirection === 'btLr'
        ? 'writing-mode:sideways-lr'
        : '',
    // font-weight before background: jsdom's CSSOM drops the background getter
    // when font-weight follows it (order is irrelevant to real browsers)
    node.attrs.bold ? 'font-weight:600' : '',
    node.attrs.color ? `color:#${node.attrs.color}` : '',
    node.attrs.fill ? `background:#${node.attrs.fill}` : '',
    node.attrs.align ? `text-align:${node.attrs.align}` : '',
    node.attrs.vAlign && node.attrs.vAlign !== 'top'
      ? `vertical-align:${node.attrs.vAlign === 'center' ? 'middle' : 'bottom'}`
      : '',
    borderCss('top'),
    borderCss('left'),
    borderCss('bottom'),
    borderCss('right'),
    // tcMar only overrides declared sides; the rest inherit the table-level --doc-cell-pad
    ...['top', 'left', 'bottom', 'right'].map((side) =>
      mar?.[side] !== undefined ? `padding-${side}:${(mar[side] / 15).toFixed(1)}px` : '',
    ),
    Array.isArray(node.attrs.colwidth)
      ? `width:${(node.attrs.colwidth as number[]).reduce((sum, width) => sum + width, 0)}px`
      : '',
  ].filter(Boolean)
  if (styles.length > 0) attrs.style = styles.join(';')
  const cellRev = node.attrs.cellRevision as { kind?: string; author?: string } | null
  if (cellRev?.kind) {
    attrs.class = `cell-rev-${cellRev.kind}`
    if (cellRev.author) attrs.title = cellRev.author
  }
  return attrs
}

export type TableBordersAttr = Partial<
  Record<
    'top' | 'left' | 'bottom' | 'right' | 'insideH' | 'insideV',
    { style: string; szEighths?: number; color?: string }
  >
>

/**
 * Table-level w:tblBorders → CSS variables consumed by edge/inside cell rules
 * (--doc-b-t/r/b/l on edge cells beat the inside lines even when the frame is
 * explicitly none). Undeclared = no borders, matching Word's printed output.
 */
export function tableBordersCss(b: TableBordersAttr | null): string[] {
  if (!b) return []
  const styles: string[] = []
  const edge = { top: 't', right: 'r', bottom: 'b', left: 'l' } as const
  for (const side of ['top', 'right', 'bottom', 'left'] as const) {
    styles.push(`--doc-b-${edge[side]}:${borderLineCss(b[side]) ?? 'none'}`)
  }
  styles.push(`--doc-b-h:${borderLineCss(b.insideH) ?? 'none'}`)
  styles.push(`--doc-b-v:${borderLineCss(b.insideV) ?? 'none'}`)
  return styles
}

/** Table-level w:tblCellMar → padding shorthand; undeclared sides use Word defaults (0 top/bottom, 108 twips left/right) */
export function cellPadCss(
  mar: { top?: number; right?: number; bottom?: number; left?: number } | null,
): string | null {
  if (!mar) return null
  const px = (v: number | undefined, dflt: number) => ((v ?? dflt) / 15).toFixed(1)
  return `${px(mar.top, 0)}px ${px(mar.right, 108)}px ${px(mar.bottom, 0)}px ${px(mar.left, 108)}px`
}

export const DocTable = Node.create({
  name: 'docTable',
  group: 'block',
  content: 'docTableRow+',
  isolating: true,
  selectable: true,
  draggable: true,
  addAttributes() {
    return {
      docxIndex: { default: null as number | null },
      colWidthsPct: { default: null as number[] | null },
      widthPx: { default: null as number | null },
      widthPct: { default: null as number | null },
      cellMar: { default: null as Record<string, number> | null },
      borders: { default: null as Record<string, unknown> | null },
      tblAlign: { default: null as string | null },
      indentTwips: { default: null as number | null },
      tblStyleId: { default: null as string | null },
      /** RTL table (tblPr w:bidiVisual): columns right to left */
      bidiVisual: { default: false },
      originalStructure: { default: null as string | null },
      originalFormatting: { default: null as string | null },
      blockRevision: { default: null as Record<string, string> | null },
    }
  },
  parseHTML() {
    return [{ tag: 'table.doc-table' }, { tag: 'table' }]
  },
  renderHTML({ node }) {
    const attrs: Record<string, string> = { class: 'doc-table' }
    if (node.attrs.docxIndex !== null) attrs['data-idx'] = String(node.attrs.docxIndex)
    if (node.attrs.tblStyleId) attrs['data-tbl-style'] = String(node.attrs.tblStyleId)
    if (node.attrs.bidiVisual) attrs.dir = 'rtl'
    const styles: string[] = []
    if (node.attrs.widthPct) styles.push(`width:${Number(node.attrs.widthPct)}%`)
    // min() keeps legacy over-wide grids on the paper instead of the gray canvas
    else if (node.attrs.widthPx) styles.push(`width:min(${Number(node.attrs.widthPx)}px,100%)`)
    const pad = cellPadCss(node.attrs.cellMar as Record<string, number> | null)
    if (pad) styles.push(`--doc-cell-pad:${pad}`)
    styles.push(...tableBordersCss(node.attrs.borders as TableBordersAttr | null))
    if (node.attrs.tblAlign === 'center') styles.push('margin-left:auto', 'margin-right:auto')
    else if (node.attrs.tblAlign === 'right') styles.push('margin-left:auto')
    else if (node.attrs.indentTwips) {
      styles.push(`margin-left:${(Number(node.attrs.indentTwips) / 15).toFixed(1)}px`)
    }
    if (styles.length > 0) attrs.style = styles.join(';')
    // A colgroup with normalized percentages defines the column grid whenever the
    // pct list matches the grid, so a table clamped to the content box compresses
    // its columns proportionally instead of overflowing via fixed td px widths.
    let firstRowCols = 0
    node.firstChild?.forEach((cell) => {
      firstRowCols += Number(cell.attrs.colspan) || 1
    })
    const rawPct = node.attrs.colWidthsPct as number[] | null
    if (rawPct?.length) {
      // zero-width grid slots get a small floor, short grids pad with the average —
      // dropping the whole colgroup falls back to fixed-layout even splitting, which
      // is always worse than an approximate grid
      const pct = rawPct.map((w) => (w > 0 ? w : 0.5))
      const avg = pct.reduce((sum, w) => sum + w, 0) / pct.length
      while (pct.length < firstRowCols) pct.push(avg)
      const total = pct.reduce((sum, w) => sum + w, 0) || 100
      return [
        'table',
        attrs,
        [
          'colgroup',
          {},
          ...pct.map(
            (w) => ['col', { style: `width:${((w / total) * 100).toFixed(2)}%` }] as const,
          ),
        ],
        ['tbody', 0],
      ]
    }
    return ['table', attrs, ['tbody', 0]]
  },
})

export const DocTableRow = Node.create({
  name: 'docTableRow',
  content: '(docTableCell | docTableHeader)+',
  addAttributes() {
    return {
      heightTwips: { default: null as number | null },
      rawTrPr: { default: null as string | null },
      /** trPr w:ins/w:del row-level revision ({kind, author, ...} | null) */
      rowRevision: { default: null as Record<string, string> | null },
    }
  },
  parseHTML() {
    return [{ tag: 'tr' }]
  },
  renderHTML({ node }) {
    const h = node.attrs.heightTwips as number | null
    const rev = node.attrs.rowRevision as { kind?: string; author?: string } | null
    const attrs: Record<string, string> = {}
    if (h) attrs.style = `height:${((h / 1440) * 96).toFixed(1)}px`
    if (rev?.kind) {
      attrs.class = `row-rev-${rev.kind}`
      if (rev.author) attrs.title = rev.author
    }
    return ['tr', attrs, 0]
  },
})

export const DocTableCell = Node.create({
  name: 'docTableCell',
  content: '(docParagraph | docListItem)+ docNestedTable*',
  isolating: true,
  addAttributes() {
    return tableCellAttrs
  },
  parseHTML() {
    return [{ tag: 'td' }]
  },
  renderHTML({ node }) {
    return ['td', tableCellHtml(node), 0]
  },
})

export const DocTableHeader = Node.create({
  name: 'docTableHeader',
  content: '(docParagraph | docListItem)+ docNestedTable*',
  isolating: true,
  addAttributes() {
    return tableCellAttrs
  },
  parseHTML() {
    return [{ tag: 'th' }]
  },
  renderHTML({ node }) {
    return ['th', tableCellHtml(node), 0]
  },
})

/** Nested table inside a cell: read-only atomic child table (editing the outer cell's text
 *  doesn't affect it; saving is byte-faithful via the outer table's originalXml, dropped on structural rebuild — matching old behavior) */
export const DocNestedTable = Node.create({
  name: 'docNestedTable',
  atom: true,
  selectable: false,
  addAttributes() {
    return { model: { default: null as TableModel | null } }
  },
  parseHTML() {
    return []
  },
  renderHTML({ node }) {
    const model = node.attrs.model as TableModel | null
    if (!model?.rows?.length) return ['div', { class: 'doc-nested-table' }]
    return ['div', { class: 'doc-nested-table', contenteditable: 'false' }, renderTableSpec(model)]
  },
  // in-place cell editing: contenteditable island; on blur the text is committed back to the model attribute
  // (saving goes through the outer table's nested-text surgical patch; cells that themselves contain nested tables stay non-editable)
  addNodeView() {
    return ({ node, editor, getPos }) => {
      let currentNode = node
      const dom = document.createElement('div')
      dom.className = 'doc-nested-table'
      dom.setAttribute('contenteditable', 'false')

      /** this table's own td elements (excluding deeper nested-table td), ordered like the model's non-vMerge-continue cells */
      const ownTds = (): HTMLElement[] => {
        const root = dom.querySelector('table')
        if (!root) return []
        return Array.from(dom.querySelectorAll('td')).filter((td) => td.closest('table') === root)
      }

      const flatCells = (): TableCell[] => {
        const model = currentNode.attrs.model as TableModel | null
        const flat: TableCell[] = []
        model?.rows.forEach((row) =>
          row.forEach((cell) => {
            if (cell.vMerge !== 'continue') flat.push(cell)
          }),
        )
        return flat
      }

      const applyEditable = () => {
        const cells = flatCells()
        ownTds().forEach((td, i) => {
          const editable = editor.isEditable && cells[i] && !cells[i].nestedTables?.length
          td.setAttribute('contenteditable', editable ? 'true' : 'false')
        })
      }

      const render = () => {
        dom.innerHTML = ''
        const model = currentNode.attrs.model as TableModel | null
        if (model?.rows?.length) {
          const rendered = DOMSerializer.renderSpec(document, renderTableSpec(model) as never)
          dom.appendChild(rendered.dom)
        }
        applyEditable()
      }
      render()

      const commit = () => {
        const model = currentNode.attrs.model as TableModel | null
        if (!model) return
        const tds = ownTds()
        let k = 0
        let changed = false
        const rows = model.rows.map((row) =>
          row.map((cell) => {
            if (cell.vMerge === 'continue') return cell
            const td = tds[k++]
            if (!td || cell.nestedTables?.length) return cell
            const paras = tdParas(td)
            if (paras.join('\n') === cell.paras.join('\n')) return cell
            changed = true
            const next = { ...cell, paras }
            delete next.richParas
            return next
          }),
        )
        if (!changed) return
        const pos = getPos()
        if (typeof pos !== 'number') return
        editor.view.dispatch(
          editor.view.state.tr.setNodeMarkup(pos, undefined, { model: { ...model, rows } }),
        )
      }

      const onFocusOut = (e: Event) => {
        const next = (e as FocusEvent).relatedTarget as HTMLElement | null
        if (next && dom.contains(next)) return
        commit()
      }
      dom.addEventListener('focusout', onFocusOut)
      window.addEventListener('ai-docs-commit-tables', commit)
      editor.on('update', applyEditable)

      return {
        dom,
        update: (n: PmNode) => {
          if (n.type.name !== 'docNestedTable') return false
          if (!n.eq(currentNode)) {
            currentNode = n
            render()
          } else {
            currentNode = n
          }
          return true
        },
        // edits stay in the DOM until the focusout commit; don't let ProseMirror re-parse them
        ignoreMutation: () => true,
        stopEvent: (event: Event) => {
          const target = event.target as HTMLElement | null
          return !!target?.closest?.('td[contenteditable="true"]')
        },
        destroy: () => {
          window.removeEventListener('ai-docs-commit-tables', commit)
          editor.off('update', applyEditable)
        },
      }
    }
  },
})

/** Delete only an explicitly selected whole table; leave cursors and partial cell selections alone. */
export function deleteSelectedWholeTable(
  state: EditorState,
  dispatch?: (transaction: Transaction) => void,
): boolean {
  const { selection } = state
  if (selection instanceof NodeSelection) {
    if (selection.node.type.spec.tableRole !== 'table') return false
    dispatch?.(state.tr.delete(selection.from, selection.to).scrollIntoView())
    return true
  }
  if (
    selection instanceof CellSelection &&
    selection.isRowSelection() &&
    selection.isColSelection()
  ) {
    return deleteTable(state, dispatch)
  }
  return false
}

export const NativeTableSupport = Extension.create({
  name: 'nativeTableSupport',
  extendNodeSchema(extension) {
    if (extension.name === 'docTable') return { tableRole: 'table' }
    if (extension.name === 'docTableRow') return { tableRole: 'row' }
    if (extension.name === 'docTableCell') return { tableRole: 'cell' }
    if (extension.name === 'docTableHeader') return { tableRole: 'header_cell' }
    return {}
  },
  addKeyboardShortcuts() {
    const deleteWholeTable = () =>
      deleteSelectedWholeTable(this.editor.state, this.editor.view.dispatch)
    return {
      Tab: () => {
        const { view } = this.editor
        if (goToNextCell(1)(this.editor.state, view.dispatch)) return true
        // last cell: Word appends a row and moves into its first cell
        if (!isInTable(this.editor.state)) return false
        if (!addRowAfter(this.editor.state, view.dispatch)) return false
        return goToNextCell(1)(this.editor.state, view.dispatch)
      },
      'Shift-Tab': () => goToNextCell(-1)(this.editor.state, this.editor.view.dispatch),
      Backspace: deleteWholeTable,
      Delete: deleteWholeTable,
    }
  },
  addProseMirrorPlugins() {
    return [
      new Plugin({
        props: {
          handleDOMEvents: {
            mouseup: (view) => {
              const resizeState = columnResizingPluginKey.getState(view.state) as
                { dragging?: unknown; activeHandle?: number } | undefined
              if (resizeState?.dragging == null) return false
              const handle = resizeState.activeHandle ?? -1
              if (handle < 0) return false
              // The resize plugin commits its final width in its window-level mouseup,
              // which runs after this handler and its microtasks; wait a macrotask so
              // the committed grid, not the previous drag frame, is constrained.
              window.setTimeout(() => {
                const raw = getComputedStyle(view.dom).getPropertyValue('--section-content-w')
                const maxWidth = Number.parseFloat(raw)
                if (Number.isFinite(maxWidth) && maxWidth > 0) {
                  constrainTableWidthAtCell(handle, maxWidth)(view.state, view.dispatch)
                }
              }, 0)
              return false
            },
          },
        },
      }),
      columnResizing({ View: null, cellMinWidth: 40, lastColumnResizable: true }),
      tableEditing({ allowTableNodeSelection: true }),
    ]
  },
})

/** Protected whole-unit blocks: images, passthrough (charts, math, ...). */
export const DocProtected = Node.create({
  name: 'docProtected',
  group: 'block',
  atom: true,
  // Editable descendants switch this off on pointer-down; the explicit handle
  // switches it back on for whole-object movement.
  draggable: true,
  selectable: true,
  addAttributes() {
    return {
      docxIndex: { default: null as number | null },
      blockRevision: { default: null as Record<string, string> | null },
      blockType: { default: 'passthrough' },
      label: { default: '' },
      previewText: { default: '' },
      imageDataUrl: { default: null as string | null },
      oleProgId: { default: null as string | null },
      /** display size in CSS px (blockType === 'image'), editable via drag handles */
      imageWidthPx: { default: null as number | null },
      imageHeightPx: { default: null as number | null },
      /** paragraph alignment of the image (w:jc) */
      imageAlign: { default: null as string | null },
      imageWrap: { default: null as string | null },
      /**
       * Free-position offset (EMU) of a floating image with wp:posOffset.
       * Used for drag-to-reposition. Null when the image uses named alignment
       * or is inline.
       */
      imageOffsetXEmu: { default: null as number | null },
      /** margin-relative wp:align preset (Word position gallery) */
      imagePosH: { default: null as string | null },
      imagePosV: { default: null as string | null },
      imageOffsetYEmu: { default: null as number | null },
      /** display-only table structure (blockType === 'table') */
      table: { default: null as TableModel | null },
      /** display-only rendering for field passthrough paragraphs */
      fieldDisplay: { default: null as FieldDisplay | null },
      /** decorative rule drawing: render as a horizontal line, not a chip */
      decorative: { default: false },
      /** display-only anchored textboxes (code boxes, callout cards) */
      textboxes: { default: null as TextboxDisplay[] | null },
      /** editable OMML leaf tokens; formula structure remains protected */
      formulaDisplay: { default: null as FormulaDisplay | null },
      /** embedded chart data model; cached texts/numbers editable, structure protected */
      chartDisplay: { default: null as ChartDisplay | null },
      /** self-contained OOXML fragment for editor-created content (new tables) */
      genXml: { default: null as string | null },
      /** new image awaiting embedding at save time */
      genImage: {
        default: null as { base64: string; mime: string; widthPx: number; heightPx: number } | null,
      },
      /** new chart awaiting embedding at save time (data snapshot; edits live in chartDisplay) */
      genChart: { default: null as NewChart | null },
    }
  },
  parseHTML() {
    return [{ tag: 'div[data-doc-protected]' }]
  },
  renderHTML({ node }) {
    return protectedDomSpec(node) as never
  },
  addNodeView() {
    return ({ node, editor, getPos }) => {
      let currentNode = node
      const dom = buildProtectedDom(currentNode)
      const getNode = () => currentNode
      const pos = getPos as () => number | undefined
      const textboxes = mountTextboxEditors(dom, getNode, pos, editor.view)
      const table = wireTableEditing(dom, getNode, pos, editor.view)
      const field = wireFieldEditing(dom, getNode, pos, editor.view)
      const formula = wireFormulaEditing(dom, getNode, pos, editor.view)
      const chart = wireChartEditing(dom, getNode, pos, editor.view)
      drawChartSvg(dom, currentNode.attrs.chartDisplay as ChartDisplay | null)
      const cleanups = [
        wireProtectedInteractionMode(
          dom,
          pos,
          editor.view,
          textboxes ?? table ?? field ?? formula ?? chart,
        ),
        wireFormulaLatexEdit(dom, getNode, pos),
        table?.cleanup,
        field?.cleanup,
        formula?.cleanup,
        chart?.cleanup,
        textboxes?.cleanup,
      ]
      return {
        dom,
        update: (n: PmNode) => {
          if (n.type.name !== 'docProtected') return false
          if (n.eq(currentNode)) {
            currentNode = n
            return true
          }
          // textbox commits only swap the textboxes attr; the sub-editors are
          // the source of truth, so keep the DOM (and editing session) alive
          if (textboxes && attrsEqualExcept(n.attrs, currentNode.attrs, 'textboxes')) {
            currentNode = n
            textboxes.sync(n.attrs.textboxes as TextboxDisplay[] | null)
            return true
          }
          // other attribute change (resize, table commit, ...): recreate DOM
          return false
        },
        // cell edits live in the DOM until committed on focusout; never re-parse
        ignoreMutation: () => true,
        stopEvent: (event: Event) => {
          const target = event.target as HTMLElement | null
          if (target?.closest?.('.doc-formula-edit')) return true
          // Let ProseMirror plugins receive handle presses; floating-object
          // dragging is implemented at the editor-view level.
          if (target?.closest?.('.doc-move-handle')) return false
          // Shape bodies (prst textboxes) drag-to-move on a plain press (Word
          // parity); a double-click still reaches the inner editor for the caret
          if (
            event.type === 'mousedown' &&
            (event as MouseEvent).detail < 2 &&
            target?.closest?.('.doc-textbox') &&
            !dom.classList.contains('doc-content-editing') &&
            (currentNode.attrs.textboxes as TextboxDisplay[] | null)?.[0]?.prst
          ) {
            return false
          }
          const contentTarget = target?.closest?.(EDITABLE_PROTECTED_SELECTOR)
          return (
            !!contentTarget &&
            (event.type === 'mousedown' || dom.classList.contains('doc-content-editing'))
          )
        },
        destroy: () => cleanups.forEach((c) => c?.()),
      }
    }
  },
  addProseMirrorPlugins() {
    return [imageResizePlugin(), floatingObjectDragPlugin()]
  },
})

/** shared DOM spec for protected blocks (renderHTML + node view) */
function protectedDomSpec(node: PmNode): DomSpec {
  const {
    blockType,
    label,
    previewText,
    imageDataUrl,
    docxIndex,
    table,
    fieldDisplay,
    decorative,
    textboxes,
    formulaDisplay,
    chartDisplay,
  } = node.attrs
  const attrs: Record<string, string> = {
    'data-doc-protected': String(blockType),
    'data-idx': docxIndex === null ? '' : String(docxIndex),
    class: `doc-protected doc-protected-${blockType}`,
  }
  if (decorative) {
    attrs.class += ' doc-protected-rule'
    return ['div', attrs, ['span', { class: 'doc-rule-line' }]]
  }
  if (Array.isArray(textboxes) && textboxes.length > 0) {
    attrs.class += ' doc-protected-textboxes'
    const offsetX = node.attrs.imageOffsetXEmu
    const offsetY = node.attrs.imageOffsetYEmu
    if (offsetX != null || offsetY != null) {
      attrs.style =
        `transform:translate(${Number(offsetX ?? 0) / EMU_PER_PX}px,` +
        `${Number(offsetY ?? 0) / EMU_PER_PX}px)`
    }
    const children: DomSpec[] = (textboxes as TextboxDisplay[]).map(renderTextboxSpec)
    // corner resize handle; multi-box nodes keep per-box autogrow semantics only
    if ((textboxes as TextboxDisplay[]).length === 1) {
      children.push(['span', { class: 'box-resize-handle', contenteditable: 'false' }])
    }
    return ['div', attrs, moveHandleSpec(t('editorMoveTextbox')), ...children]
  }
  // empty section-break paragraphs (page-per-section converter output): Word
  // shows nothing here, so render a near-invisible strip (hover reveals it)
  if (label === 'Section break paragraph' && !previewText) {
    attrs.class += ' doc-protected-sectbreak'
    return ['div', attrs, ['span', { class: 'doc-sectbreak-label' }, t('editorSectionBreak')]]
  }
  // TOC field boundary paragraphs (fldChar begin + instruction / lone fldChar
  // end) have no visible result; Word shows nothing there either, so they get
  // the same near-invisible strip (hover/selection reveals the label)
  if (
    !previewText &&
    !fieldDisplay &&
    (label === 'Auto TOC (updates when opened in Word)' || label === 'Field end marker')
  ) {
    attrs.class += ' doc-protected-sectbreak'
    return ['div', attrs, ['span', { class: 'doc-sectbreak-label' }, String(label)]]
  }
  if (blockType === 'image' && imageDataUrl) {
    const { imageWidthPx, imageHeightPx, imageAlign, imageWrap } = node.attrs
    if (imageAlign === 'center' || imageAlign === 'right') {
      attrs['style'] = `text-align:${imageAlign}`
    }
    if (imageWrap) attrs.class += ` img-wrap-${String(imageWrap)}`
    const imgAttrs: Record<string, string> = {
      src: String(imageDataUrl),
      class: 'doc-protected-img',
    }
    if (imageWidthPx) {
      imgAttrs['style'] =
        `width:${Number(imageWidthPx)}px;` +
        (imageHeightPx ? `height:${Number(imageHeightPx)}px` : 'height:auto')
    }
    return [
      'div',
      attrs,
      moveHandleSpec(t('editorMoveImage')),
      [
        'span',
        { class: 'doc-img-wrap' },
        ['img', imgAttrs],
        ['span', { class: 'img-resize-handle' }],
      ],
    ]
  }
  if (blockType === 'table' && table && (table as TableModel).rows?.length) {
    return [
      'div',
      attrs,
      moveHandleSpec(t('editorMoveTable')),
      renderTableSpec(table as TableModel),
    ]
  }
  if (fieldDisplay) {
    const spec = renderFieldSpec(fieldDisplay as FieldDisplay)
    if (spec) {
      attrs.class += ' doc-protected-field'
      return ['div', attrs, spec]
    }
  }
  if ((formulaDisplay as FormulaDisplay | null)?.tokens?.length) {
    attrs.class += ' doc-protected-formula'
    if ((formulaDisplay as FormulaDisplay).mathml) attrs.class += ' doc-protected-formula-display'
    return [
      'div',
      attrs,
      moveHandleSpec(t('editorMoveEquation')),
      renderFormulaSpec(formulaDisplay as FormulaDisplay),
    ]
  }
  if ((chartDisplay as ChartDisplay | null)?.series?.length) {
    attrs.class += ' doc-protected-chart'
    return [
      'div',
      attrs,
      moveHandleSpec(t('editorMoveChart')),
      renderChartSpec(chartDisplay as ChartDisplay),
      ['span', { class: 'box-resize-handle', contenteditable: 'false' }],
    ]
  }
  // OLE embed with a packaged preview picture: show the picture with a
  // friendly type caption instead of a bare "Embedded object" label
  const oleCaption = label === 'Embedded object' ? oleTypeLabel(node.attrs.oleProgId) : null
  if (imageDataUrl && blockType === 'passthrough') {
    attrs.class += ' doc-protected-ole'
    return [
      'div',
      attrs,
      [
        'span',
        { class: 'doc-ole-wrap' },
        ['img', { src: String(imageDataUrl), class: 'doc-ole-img' }],
      ],
      ['span', { class: 'doc-protected-label' }, oleCaption ?? String(label)],
    ]
  }
  const children: unknown[] = [
    [
      'span',
      { class: 'doc-protected-label' },
      oleCaption ?? String(label || t('editorProtectedContent')),
    ],
  ]
  if (previewText) children.push(['span', { class: 'doc-protected-preview' }, String(previewText)])
  return ['div', attrs, ...children]
}

/** o:OLEObject ProgID → localized friendly kind */
function oleTypeLabel(progId: unknown): string {
  const id = typeof progId === 'string' ? progId : ''
  if (id.startsWith('Excel.')) return t('editorOleExcel')
  if (id.startsWith('Word.')) return t('editorOleWord')
  if (id.startsWith('PowerPoint.')) return t('editorOlePpt')
  if (id.startsWith('AcroExch')) return t('editorOlePdf')
  return id ? `${t('editorOleGeneric')} (${id.split('.')[0]})` : t('editorOleGeneric')
}

function moveHandleSpec(label: string): DomSpec {
  return [
    'span',
    {
      class: 'doc-move-handle',
      title: label,
      'aria-label': label,
      contenteditable: 'false',
    },
    '↕',
  ]
}

export function buildProtectedDom(node: PmNode): HTMLElement {
  const { dom } = DOMSerializer.renderSpec(document, protectedDomSpec(node) as never)
  const el = dom as HTMLElement
  const mathml = (node.attrs.formulaDisplay as FormulaDisplay | null)?.mathml
  const mathHost = el.querySelector?.('.doc-formula-math')
  if (mathml && mathHost) mathHost.innerHTML = mathml
  return el
}

export interface ProtectedContentEditor {
  setEditable(editable: boolean): void
  commit(): void
}

const EDITABLE_PROTECTED_SELECTOR =
  'td, .doc-textbox, .doc-toc-title, .doc-toc-page, .doc-field-text, .doc-formula-token, ' +
  '.doc-formula-math, .doc-chart-title, .doc-chart-cell'

/** Object mode (single click/drag) and text mode (double click). */
function wireProtectedInteractionMode(
  dom: HTMLElement,
  getPos: () => number | undefined,
  view: EditorView,
  contentEditor: ProtectedContentEditor | null,
): (() => void) | null {
  const handle = dom.querySelector('.doc-move-handle') as HTMLElement | null
  if (!handle && !contentEditor) return null

  const selectObject = () => {
    const pos = getPos()
    if (typeof pos !== 'number') return
    view.dispatch(view.state.tr.setSelection(NodeSelection.create(view.state.doc, pos)))
  }
  const setEditing = (editing: boolean) => {
    if (!contentEditor) editing = false
    if (editing === dom.classList.contains('doc-content-editing')) {
      contentEditor?.setEditable(editing)
      dom.draggable = !!handle && !editing
      return
    }
    if (!editing) contentEditor?.commit()
    contentEditor?.setEditable(editing)
    dom.classList.toggle('doc-content-editing', editing)
    dom.draggable = !!handle && !editing
  }
  const onMouseDown = (event: MouseEvent) => {
    if (event.button !== 0) return
    const target = event.target as HTMLElement | null
    if (!target) return
    if (target.closest('.doc-move-handle')) {
      setEditing(false)
      selectObject()
      return
    }
    const content = target.closest(EDITABLE_PROTECTED_SELECTOR)
    if (content && dom.contains(content)) {
      if (!dom.classList.contains('doc-content-editing')) selectObject()
      dom.draggable = !!handle && !dom.classList.contains('doc-content-editing')
      return
    }
    if (dom.classList.contains('doc-content-editing')) setEditing(false)
    selectObject()
  }
  const onDoubleClick = (event: MouseEvent) => {
    if (!contentEditor) return
    const target = event.target as HTMLElement | null
    const content = target?.closest(EDITABLE_PROTECTED_SELECTOR)
    if (!content || !dom.contains(content)) return
    setEditing(true)
  }
  const onDocumentMouseDown = (event: MouseEvent) => {
    const target = event.target
    if (
      target instanceof HTMLElement &&
      !dom.contains(target) &&
      dom.classList.contains('doc-content-editing')
    ) {
      setEditing(false)
    }
  }
  const onKeyDown = (event: KeyboardEvent) => {
    if (event.key !== 'Escape' || !dom.classList.contains('doc-content-editing')) return
    setEditing(false)
    selectObject()
    dom.focus()
  }

  setEditing(false)
  dom.addEventListener('mousedown', onMouseDown, true)
  dom.addEventListener('dblclick', onDoubleClick, true)
  document.addEventListener('mousedown', onDocumentMouseDown, true)
  dom.addEventListener('keydown', onKeyDown, true)
  return () => {
    dom.removeEventListener('mousedown', onMouseDown, true)
    dom.removeEventListener('dblclick', onDoubleClick, true)
    document.removeEventListener('mousedown', onDocumentMouseDown, true)
    dom.removeEventListener('keydown', onKeyDown, true)
  }
}

/** split a contenteditable cell back into paragraph strings */
function tdParas(td: HTMLElement): string[] {
  const text = (td.innerText ?? td.textContent ?? '').replace(/\u00a0/g, ' ').replace(/\n+$/, '')
  const paras = text.split('\n')
  if (paras.length === 1 && paras[0].trim() === '') return ['']
  return paras
}

/**
 * In-place cell editing: cells become contenteditable islands and
 * their text is committed back into the node's TableModel when focus leaves
 * the table (or when App broadcasts 'ai-docs-commit-tables' before saving).
 */
function wireTableEditing(
  dom: HTMLElement,
  getNode: () => PmNode,
  getPos: () => number | undefined,
  view: EditorView,
): (ProtectedContentEditor & { cleanup(): void }) | null {
  const node = getNode()
  if (node.attrs.blockType !== 'table' || !node.attrs.table) return null
  const setEditable = (editable: boolean) => {
    for (const td of Array.from(dom.querySelectorAll('td'))) {
      td.setAttribute('contenteditable', editable ? 'true' : 'false')
    }
  }

  const commit = () => {
    const current = getNode()
    const model = current.attrs.table as TableModel
    const domTds = Array.from(dom.querySelectorAll('td'))
    let k = 0
    let changed = false
    const rows = model.rows.map((row) =>
      row.map((cell) => {
        if (cell.vMerge === 'continue') return cell
        const td = domTds[k++]
        if (!td) return cell
        const paras = tdParas(td as HTMLElement)
        if (paras.join('\n') === cell.paras.join('\n')) return cell
        changed = true
        return { ...cell, paras }
      }),
    )
    if (!changed) return
    const pos = getPos()
    if (typeof pos !== 'number') return
    view.dispatch(
      view.state.tr.setNodeMarkup(pos, undefined, { ...current.attrs, table: { ...model, rows } }),
    )
  }

  dom.addEventListener('focusout', (e) => {
    const next = (e as FocusEvent).relatedTarget as HTMLElement | null
    if (next && dom.contains(next)) return // moving between cells: not yet
    commit()
  })
  window.addEventListener('ai-docs-commit-tables', commit)
  return {
    setEditable,
    commit,
    cleanup: () => window.removeEventListener('ai-docs-commit-tables', commit),
  }
}

export function protectedText(element: HTMLElement): string {
  return (element.innerText ?? element.textContent ?? '').replace(/\u00a0/g, '')
}

export function preventProtectedLineBreak(event: KeyboardEvent) {
  if (event.key === 'Enter') event.preventDefault()
}

/** Edit cached visible field results without exposing field instructions. */
function wireFieldEditing(
  dom: HTMLElement,
  getNode: () => PmNode,
  getPos: () => number | undefined,
  view: EditorView,
): (ProtectedContentEditor & { cleanup(): void }) | null {
  const field = getNode().attrs.fieldDisplay as FieldDisplay | null
  if (!field || field.kind === 'pageBreak') return null
  const targets = Array.from(
    dom.querySelectorAll<HTMLElement>('.doc-toc-title, .doc-toc-page, .doc-field-text'),
  )
  if (targets.length === 0) return null

  const setEditable = (editable: boolean) => {
    for (const target of targets)
      target.setAttribute('contenteditable', editable ? 'true' : 'false')
  }
  const commit = () => {
    const current = getNode()
    const currentField = current.attrs.fieldDisplay as FieldDisplay | null
    if (!currentField) return
    const next: FieldDisplay = { ...currentField }
    if (currentField.kind === 'tocLine') {
      const left = dom.querySelector<HTMLElement>('.doc-toc-title')
      const right = dom.querySelector<HTMLElement>('.doc-toc-page')
      if (left) next.left = protectedText(left)
      if (right) next.right = protectedText(right)
    } else {
      const text = dom.querySelector<HTMLElement>('.doc-field-text')
      if (text) next.left = protectedText(text)
    }
    if (JSON.stringify(next) === JSON.stringify(currentField)) return
    const pos = getPos()
    if (typeof pos !== 'number') return
    view.dispatch(
      view.state.tr.setNodeMarkup(pos, undefined, { ...current.attrs, fieldDisplay: next }),
    )
  }
  for (const target of targets) target.addEventListener('keydown', preventProtectedLineBreak)
  window.addEventListener('ai-docs-commit-tables', commit)
  return {
    setEditable,
    commit,
    cleanup: () => {
      for (const target of targets) target.removeEventListener('keydown', preventProtectedLineBreak)
      window.removeEventListener('ai-docs-commit-tables', commit)
    },
  }
}

/** hover button on display formulas whose LaTeX was recovered: full re-edit */
function wireFormulaLatexEdit(
  dom: HTMLElement,
  getNode: () => PmNode,
  getPos: () => number | undefined,
): (() => void) | null {
  const formula = getNode().attrs.formulaDisplay as FormulaDisplay | null
  if (!formula?.latex) return null
  const host = dom.querySelector('.doc-formula-wrap') ?? dom
  const button = document.createElement('button')
  button.className = 'doc-formula-edit'
  button.title = t('editorEditFormulaLatex')
  button.textContent = t('editorEdit')
  button.setAttribute('contenteditable', 'false')
  button.addEventListener('mousedown', (e) => {
    e.preventDefault()
    e.stopPropagation()
  })
  button.addEventListener('click', (e) => {
    e.preventDefault()
    e.stopPropagation()
    const pos = getPos()
    const latex = (getNode().attrs.formulaDisplay as FormulaDisplay | null)?.latex
    if (typeof pos === 'number' && latex) {
      window.dispatchEvent(
        new CustomEvent('ai-docs-edit-inline-math', {
          detail: { pos, latex, kind: 'block' },
        }),
      )
    }
  })
  host.appendChild(button)
  return () => button.remove()
}

/** Edit OMML leaf tokens while keeping formula structure outside the editor. */
function wireFormulaEditing(
  dom: HTMLElement,
  getNode: () => PmNode,
  getPos: () => number | undefined,
  view: EditorView,
): (ProtectedContentEditor & { cleanup(): void }) | null {
  const formula = getNode().attrs.formulaDisplay as FormulaDisplay | null
  if (!formula?.tokens.length) return null
  const targets = Array.from(dom.querySelectorAll<HTMLElement>('.doc-formula-token'))
  if (targets.length !== formula.tokens.length) return null

  const setEditable = (editable: boolean) => {
    for (const target of targets)
      target.setAttribute('contenteditable', editable ? 'true' : 'false')
  }
  const commit = () => {
    const current = getNode()
    const currentFormula = current.attrs.formulaDisplay as FormulaDisplay | null
    if (!currentFormula || currentFormula.tokens.length !== targets.length) return
    const tokens = targets.map(protectedText)
    if (tokens.every((token, i) => token === currentFormula.tokens[i])) return
    const pos = getPos()
    if (typeof pos !== 'number') return
    // re-derive the 2D preview (and, for editor-created formulas, the OOXML
    // to be saved) from the patched OMML source
    const omml = currentFormula.omml ? patchMathTokens(currentFormula.omml, tokens) : undefined
    const nextFormula: FormulaDisplay = omml
      ? { tokens, omml, mathml: ommlToMathML(omml) }
      : { tokens }
    const attrs: Record<string, unknown> = { ...current.attrs, formulaDisplay: nextFormula }
    if (current.attrs.genXml) {
      attrs.genXml = patchMathTokens(String(current.attrs.genXml), tokens)
    }
    view.dispatch(view.state.tr.setNodeMarkup(pos, undefined, attrs))
  }
  for (const target of targets) target.addEventListener('keydown', preventProtectedLineBreak)
  window.addEventListener('ai-docs-commit-tables', commit)
  return {
    setEditable,
    commit,
    cleanup: () => {
      for (const target of targets) target.removeEventListener('keydown', preventProtectedLineBreak)
      window.removeEventListener('ai-docs-commit-tables', commit)
    },
  }
}

/** node attrs equality ignoring one key (identity per key is enough here) */
function attrsEqualExcept(
  a: Record<string, unknown>,
  b: Record<string, unknown>,
  skip: string,
): boolean {
  const keys = new Set([...Object.keys(a), ...Object.keys(b)])
  for (const key of keys) {
    if (key !== skip && !Object.is(a[key], b[key])) return false
  }
  return true
}

type TextboxPara = TextboxDisplay['paras'][number]

/** ProseMirror doc JSON for one textbox's rich content */
function textboxDocJson(box: TextboxDisplay): Record<string, unknown> {
  const paras = box.paras.map((para) => ({
    type: 'docParagraph',
    attrs: {
      align: para.align ?? null,
      lineSpacing: para.lineSpacing ?? null,
      indentLeft: para.indentLeft ?? null,
      indentRight: para.indentRight ?? null,
      indentFirstLine: para.indentFirstLine ?? null,
      spaceBefore: para.spaceBefore ?? null,
      spaceAfter: para.spaceAfter ?? null,
      shadingFill: para.shadingFill ?? null,
      borders: para.borders ?? null,
    },
    content: runsToInline(para.runs),
  }))
  return { type: 'doc', content: paras.length > 0 ? paras : [{ type: 'docParagraph' }] }
}

/** TextboxDisplay paragraphs from a sub-editor's current doc */
function subEditorParas(sub: Editor): TextboxPara[] {
  return ((sub.getJSON().content ?? []) as PmJson[]).map((p) => {
    const para: TextboxPara = { runs: inlineToRuns(p.content ?? []) }
    const attrs = p.attrs ?? {}
    const keys = [
      'align',
      'lineSpacing',
      'indentLeft',
      'indentRight',
      'indentFirstLine',
      'spaceBefore',
      'spaceAfter',
      'shadingFill',
      'borders',
    ] as const
    for (const key of keys) {
      const value = attrs[key]
      if (value !== null && value !== undefined) Object.assign(para, { [key]: value })
    }
    return para
  })
}

/**
 * Rich textbox editing: each rendered box hosts a full nested
 * Tiptap editor sharing the main schema's marks, so ribbon formatting (color,
 * bold, size, ...) works inside the box. Content commits back into the node's
 * TextboxDisplay model when focus leaves the block (or before saving); the
 * save path regenerates only the changed paragraphs inside w:txbxContent.
 */
function mountTextboxEditors(
  dom: HTMLElement,
  getNode: () => PmNode,
  getPos: () => number | undefined,
  view: EditorView,
):
  | (ProtectedContentEditor & {
      cleanup: () => void
      sync: (boxes: TextboxDisplay[] | null) => void
    })
  | null {
  let knownBoxes = getNode().attrs.textboxes as TextboxDisplay[] | null
  if (!knownBoxes || knownBoxes.length === 0) return null

  const editors: Editor[] = []
  const boxEls = Array.from(dom.querySelectorAll('.doc-textbox')) as HTMLElement[]
  const minHeights = knownBoxes.map((box) => box.minHeightPx ?? box.heightPx)
  const measuredHeights = knownBoxes.map((box) => box.heightPx)
  const resizeFrames: Array<number | undefined> = []
  const measureBox = (index: number) => {
    const el = boxEls[index]
    const minHeight = minHeights[index]
    if (!el || !minHeight) return
    const borderHeight = el.offsetHeight - el.clientHeight
    el.style.height = 'auto'
    const naturalHeight = Math.ceil(el.scrollHeight + borderHeight)
    const nextHeight = Math.max(minHeight, naturalHeight)
    el.style.height = `${nextHeight}px`
    measuredHeights[index] = nextHeight
  }
  const scheduleMeasure = (index: number) => {
    if (resizeFrames[index] !== undefined) cancelAnimationFrame(resizeFrames[index]!)
    resizeFrames[index] = requestAnimationFrame(() => {
      resizeFrames[index] = undefined
      measureBox(index)
    })
  }
  boxEls.forEach((el, i) => {
    const box = knownBoxes![i]
    if (!box) return
    // boxes whose content flattens tables / content controls into display lines
    // keep the static spec: a sub-editor commit would corrupt that structure
    if (box.readOnly) return
    el.replaceChildren() // the static spec children are replaced by the live editor
    const sub: Editor = new Editor({
      element: el,
      extensions: textboxSubExtensions,
      content: textboxDocJson(box),
      editorProps: {
        attributes: { class: 'doc-textbox-editor', spellcheck: 'false' },
      },
      onFocus: () => setActiveSubEditor(sub),
      onTransaction: ({ transaction }) => {
        notifySubEditorState(transaction.docChanged)
        if (transaction.docChanged) scheduleMeasure(i)
      },
    })
    // TipTap initializes the host element and may clear attributes emitted by
    // the static DOM spec, so reapply the shape geometry after mounting.
    el.setAttribute('style', textboxBoxStyle(box))
    editors[i] = sub
  })
  if (editors.length === 0) return null

  const setEditable = (editable: boolean) => {
    for (const sub of editors) {
      if (sub && !sub.isDestroyed) sub.setEditable(editable)
    }
  }

  const commit = () => {
    const current = getNode()
    const model = current.attrs.textboxes as TextboxDisplay[] | null
    if (!model) return
    let changed = false
    const next = model.map((box, i) => {
      const sub = editors[i]
      if (!sub || sub.isDestroyed) return box
      const paras = subEditorParas(sub)
      const same =
        paras.length === box.paras.length &&
        paras.every((p, j) => textboxParaSignature(p) === textboxParaSignature(box.paras[j]))
      const measuredHeight = minHeights[i] ? measuredHeights[i] : box.heightPx
      const sameHeight = measuredHeight === box.heightPx
      if (same && sameHeight) return box
      changed = true
      return { ...box, paras, heightPx: measuredHeight }
    })
    if (!changed) return
    const pos = getPos()
    if (typeof pos !== 'number') return
    knownBoxes = next
    view.dispatch(
      view.state.tr.setNodeMarkup(pos, undefined, { ...current.attrs, textboxes: next }),
    )
  }

  /** external model change (undo of a commit, AI edit): re-feed the sub-editors */
  const sync = (boxes: TextboxDisplay[] | null) => {
    if (!boxes || boxes === knownBoxes) return
    knownBoxes = boxes
    boxes.forEach((box, i) => {
      const sub = editors[i]
      if (sub && !sub.isDestroyed) sub.commands.setContent(textboxDocJson(box))
      const el = boxEls[i]
      if (el) el.setAttribute('style', textboxBoxStyle(box))
      // corner resize only rewrites the attrs; without refreshing the minimum
      // the next autofit would shrink the box back below the resized height
      minHeights[i] = box.minHeightPx ?? box.heightPx
      measuredHeights[i] = box.heightPx
    })
  }

  dom.addEventListener('focusout', (e) => {
    const next = (e as FocusEvent).relatedTarget as HTMLElement | null
    if (next && dom.contains(next)) return // moving between boxes: not yet
    commit()
  })
  window.addEventListener('ai-docs-commit-tables', commit)
  const cleanup = () => {
    window.removeEventListener('ai-docs-commit-tables', commit)
    for (const frame of resizeFrames) {
      if (frame !== undefined) cancelAnimationFrame(frame)
    }
    for (const sub of editors) {
      if (!sub) continue
      dropActiveSubEditor(sub)
      sub.destroy()
    }
  }
  return { cleanup, sync, setEditable, commit }
}

/** drag the corner handle of a selected image to resize it */
function imageResizePlugin(): Plugin {
  return new Plugin({
    props: {
      handleDOMEvents: {
        mousedown: (view, event) => {
          const target = event.target as HTMLElement
          if (target.classList?.contains('box-resize-handle')) {
            const boxWrapper = target.closest('.doc-protected') as HTMLElement | null
            const isChart = !!boxWrapper?.classList.contains('doc-protected-chart')
            const boxEl = boxWrapper?.querySelector(
              isChart ? '.doc-chart-canvas svg' : '.doc-textbox',
            ) as HTMLElement | null
            if (!boxWrapper || !boxEl) return false
            let boxPos = -1
            view.state.doc.descendants((node, p) => {
              if (boxPos !== -1) return false
              if (node.type.name === 'docProtected' && view.nodeDOM(p) === boxWrapper) boxPos = p
              return boxPos === -1
            })
            if (boxPos === -1) return false
            const attrs = view.state.doc.nodeAt(boxPos)?.attrs
            const boxes = attrs?.textboxes as TextboxDisplay[] | null
            const chart = attrs?.chartDisplay as ChartDisplay | null
            if (!isChart && (!Array.isArray(boxes) || boxes.length !== 1)) return false
            if (isChart && !chart) return false
            event.preventDefault()
            view.dispatch(view.state.tr.setSelection(NodeSelection.create(view.state.doc, boxPos)))

            const zoomEl = document.querySelector('.doc-zoom') as HTMLElement | null
            const zoom = zoomEl ? parseFloat(getComputedStyle(zoomEl).zoom || '1') || 1 : 1
            const startRect = boxEl.getBoundingClientRect()
            const startW = isChart
              ? startRect.width / zoom
              : parseFloat(getComputedStyle(boxEl).width) || boxes![0].widthPx || 189
            const startH = isChart
              ? startRect.height / zoom
              : parseFloat(getComputedStyle(boxEl).height) || boxes![0].heightPx || 113
            const startX = event.clientX
            const startY = event.clientY

            // horizontal lines resize in length only (their saved extent is zero-height)
            const lockH = !isChart && isStraightLineKind(boxes![0].prst)
            const minW = isChart ? 120 : 24
            const minH = isChart ? 80 : 8
            // charts never draw wider than the render cap, so don't let the model exceed it
            const maxW = isChart ? CHART_MAX_WIDTH_PX : Infinity
            const sizeAt = (e: MouseEvent) => ({
              w: Math.min(maxW, Math.max(minW, startW + (e.clientX - startX) / zoom)),
              h: lockH ? startH : Math.max(minH, startH + (e.clientY - startY) / zoom),
            })
            const onMove = (e: MouseEvent) => {
              const { w, h } = sizeAt(e)
              boxEl.style.width = `${w}px`
              boxEl.style.height = `${h}px`
            }
            const onUp = (e: MouseEvent) => {
              window.removeEventListener('mousemove', onMove)
              window.removeEventListener('mouseup', onUp)
              const { w, h } = sizeAt(e)
              const node = view.state.doc.nodeAt(boxPos)
              if (!node) return
              if (isChart) {
                const display = node.attrs.chartDisplay as ChartDisplay | null
                if (!display) return
                view.dispatch(
                  view.state.tr.setNodeMarkup(boxPos, undefined, {
                    ...node.attrs,
                    chartDisplay: {
                      ...display,
                      widthPx: Math.round(w),
                      heightPx: Math.round(h),
                    },
                  }),
                )
                return
              }
              const box = (node.attrs.textboxes as TextboxDisplay[] | null)?.[0]
              if (!box) return
              // straight lines keep their zero-height extent: never give them a heightPx
              const next = lockH
                ? { ...box, widthPx: Math.round(w) }
                : {
                    ...box,
                    widthPx: Math.round(w),
                    heightPx: Math.round(h),
                    minHeightPx: Math.round(h),
                  }
              view.dispatch(
                view.state.tr.setNodeMarkup(boxPos, undefined, {
                  ...node.attrs,
                  textboxes: [next],
                }),
              )
            }
            window.addEventListener('mousemove', onMove)
            window.addEventListener('mouseup', onUp)
            return true
          }
          if (!target.classList?.contains('img-resize-handle')) return false
          const wrapper = target.closest('.doc-protected') as HTMLElement | null
          const img = wrapper?.querySelector('img.doc-protected-img') as HTMLImageElement | null
          if (!wrapper || !img) return false
          event.preventDefault()

          let pos = -1
          view.state.doc.descendants((node, p) => {
            if (pos !== -1) return false
            if (node.type.name === 'docProtected' && view.nodeDOM(p) === wrapper) pos = p
            return pos === -1
          })
          if (pos === -1) return false

          view.dispatch(view.state.tr.setSelection(NodeSelection.create(view.state.doc, pos)))

          // CSS `zoom` scales client coordinates; divide it back out
          const zoomEl = document.querySelector('.doc-zoom') as HTMLElement | null
          const zoom = zoomEl ? parseFloat(getComputedStyle(zoomEl).zoom || '1') || 1 : 1
          const startRect = img.getBoundingClientRect()
          const startW = startRect.width / zoom
          const ratio = startRect.height / startRect.width
          const startX = event.clientX

          const widthAt = (e: MouseEvent) => Math.max(24, startW + (e.clientX - startX) / zoom)
          const onMove = (e: MouseEvent) => {
            const w = widthAt(e)
            img.style.width = `${w}px`
            img.style.height = `${w * ratio}px`
          }
          const onUp = (e: MouseEvent) => {
            window.removeEventListener('mousemove', onMove)
            window.removeEventListener('mouseup', onUp)
            const w = Math.round(widthAt(e))
            const node = view.state.doc.nodeAt(pos)
            if (!node) return
            view.dispatch(
              view.state.tr.setNodeMarkup(pos, undefined, {
                ...node.attrs,
                imageWidthPx: w,
                imageHeightPx: Math.round(w * ratio),
              }),
            )
          }
          window.addEventListener('mousemove', onMove)
          window.addEventListener('mouseup', onUp)
          return true
        },
      },
    },
  })
}

const EMU_PER_PX = 9525

/**
 * Drag floating images and textbox shapes (wp:anchor) to update posOffset.
 * Only handles images with numeric posOffset (imageOffsetXEmu/YEmu set).
 * Inline images (no imageWrap) are auto-converted to anchor on drag start
 * with square wrap and the initial offset derived from the drag delta.
 */
function floatingObjectDragPlugin(): Plugin {
  return new Plugin({
    props: {
      handleDOMEvents: {
        mousedown: (view, event) => {
          if (event.button !== 0) return false
          const target = event.target as HTMLElement | null
          if (!target) return false
          // Activate on the move handle of an image block, or (Word parity)
          // anywhere on a textbox/shape body — its text isn't edited in place,
          // so the body gesture is unambiguous
          const handle = target.closest('.doc-move-handle') as HTMLElement | null
          const body = handle ? null : (target.closest('.doc-textbox') as HTMLElement | null)
          if (!handle && !body) return false
          const wrapper = (handle ?? body)!.closest('.doc-protected') as HTMLElement | null
          if (!wrapper) return false

          // Find the ProseMirror node position
          let pos = -1
          view.state.doc.descendants((node, p) => {
            if (pos !== -1) return false
            if (node.type.name === 'docProtected' && view.nodeDOM(p) === wrapper) pos = p
            return pos === -1
          })
          if (pos === -1) return false
          const node = view.state.doc.nodeAt(pos)
          if (!node) return false
          const isImage = node.attrs.blockType === 'image'
          const isTextbox = Array.isArray(node.attrs.textboxes) && node.attrs.textboxes.length > 0
          if (!isImage && !isTextbox) return false
          // Body activation is for prst shapes only: plain text boxes keep
          // click-to-type, images keep their native behavior
          const isShapeBody = isTextbox && !!(node.attrs.textboxes as TextboxDisplay[])[0]?.prst
          if (!handle && !isShapeBody) return false

          const isFloating = !!node.attrs.imageWrap
          const hasNumericOffset =
            node.attrs.imageOffsetXEmu != null && node.attrs.imageOffsetYEmu != null

          // Only handle floating images (anchor) that have numeric posOffset,
          // or inline images (will be converted on first drag)
          if (isImage && isFloating && !hasNumericOffset) return false

          event.preventDefault()
          event.stopPropagation()
          view.dispatch(view.state.tr.setSelection(NodeSelection.create(view.state.doc, pos)))

          const zoomEl = document.querySelector('.doc-zoom') as HTMLElement | null
          const zoom = zoomEl ? parseFloat(getComputedStyle(zoomEl).zoom || '1') || 1 : 1
          const startX = event.clientX
          const startY = event.clientY
          const startOffsetX = Number(node.attrs.imageOffsetXEmu ?? 0)
          const startOffsetY = Number(node.attrs.imageOffsetYEmu ?? 0)

          // Visual feedback: apply CSS translate during drag
          const visual = wrapper.querySelector(
            isTextbox ? '.doc-textbox' : '.doc-protected-img',
          ) as HTMLElement | null

          // 3px threshold keeps plain clicks (select, first click of a
          // double-click-to-edit) from nudging the object
          let dragging = false
          const onMove = (e: MouseEvent) => {
            const dx = (e.clientX - startX) / zoom
            const dy = (e.clientY - startY) / zoom
            if (!dragging && Math.abs(dx) < 3 && Math.abs(dy) < 3) return
            dragging = true
            if (visual) visual.style.transform = `translate(${dx}px, ${dy}px)`
          }

          const onUp = (e: MouseEvent) => {
            window.removeEventListener('mousemove', onMove)
            window.removeEventListener('mouseup', onUp)
            if (visual) visual.style.transform = ''

            const dx = (e.clientX - startX) / zoom
            const dy = (e.clientY - startY) / zoom
            // Only commit a real drag (past the threshold)
            if (!dragging) return

            const newX = Math.round(startOffsetX + dx * EMU_PER_PX)
            const newY = Math.round(startOffsetY + dy * EMU_PER_PX)

            const currentNode = view.state.doc.nodeAt(pos)
            if (!currentNode) return
            if (isTextbox) {
              view.dispatch(
                view.state.tr.setNodeMarkup(pos, undefined, {
                  ...currentNode.attrs,
                  imageWrap: currentNode.attrs.imageWrap ?? 'square-left',
                  imageOffsetXEmu: newX,
                  imageOffsetYEmu: newY,
                  imagePosH: null,
                  imagePosV: null,
                }),
              )
            } else if (!isFloating) {
              // Inline → auto-convert to anchor with square-left wrap
              view.dispatch(
                view.state.tr.setNodeMarkup(pos, undefined, {
                  ...currentNode.attrs,
                  imageWrap: 'square-left',
                  imageOffsetXEmu: Math.max(0, newX),
                  imageOffsetYEmu: Math.max(0, newY),
                  imagePosH: null,
                  imagePosV: null,
                }),
              )
            } else {
              view.dispatch(
                view.state.tr.setNodeMarkup(pos, undefined, {
                  ...currentNode.attrs,
                  imageOffsetXEmu: Math.max(0, newX),
                  imageOffsetYEmu: Math.max(0, newY),
                  imagePosH: null,
                  imagePosV: null,
                }),
              )
            }
          }

          window.addEventListener('mousemove', onMove)
          window.addEventListener('mouseup', onUp)
          return true
        },
      },
    },
  })
}

export type DomSpec = [string, Record<string, string>, ...unknown[]]

/**
 * Paragraph node for textbox sub-editors. Named docParagraph on purpose so
 * ribbon paragraph commands (updateAttributes('docParagraph', ...)) work
 * unchanged whether they target the main editor or a textbox.
 */
const TextboxParagraph = Node.create({
  name: 'docParagraph',
  group: 'block',
  content: 'inline*',
  addAttributes() {
    return {
      align: { default: null as string | null },
      lineSpacing: { default: null as number | null },
      indentLeft: { default: null as number | null },
      indentRight: { default: null as number | null },
      indentFirstLine: { default: null as number | null },
      spaceBefore: { default: null as number | null },
      spaceAfter: { default: null as number | null },
      shadingFill: { default: null as string | null },
      borders: { default: null as string | null },
    }
  },
  parseHTML() {
    return [{ tag: 'div.doc-textbox-para' }, { tag: 'p' }]
  },
  renderHTML({ node }) {
    const attrs: Record<string, string> = {
      class: `doc-textbox-para${node.content.size === 0 ? ' doc-textbox-para-empty' : ''}`,
    }
    const styles = [
      node.attrs.align
        ? `text-align:${node.attrs.align === 'distribute' ? 'justify' : node.attrs.align}`
        : '',
      node.attrs.lineSpacing ? `line-height:${Number(node.attrs.lineSpacing) * 1.2}` : '',
      node.attrs.indentLeft ? `margin-left:${Number(node.attrs.indentLeft) / 20}pt` : '',
      node.attrs.indentRight ? `margin-right:${Number(node.attrs.indentRight) / 20}pt` : '',
      node.attrs.indentFirstLine ? `text-indent:${Number(node.attrs.indentFirstLine) / 20}pt` : '',
      node.attrs.spaceBefore != null ? `margin-top:${Number(node.attrs.spaceBefore) / 20}pt` : '',
      node.attrs.spaceAfter != null ? `margin-bottom:${Number(node.attrs.spaceAfter) / 20}pt` : '',
      node.attrs.shadingFill ? `background-color:#${node.attrs.shadingFill}` : '',
    ]
      .filter(Boolean)
      .join(';')
    if (styles) attrs.style = styles
    return ['div', attrs, 0]
  },
})

/** shared with the main editor: same mark names, so ribbon commands route 1:1 */
const textboxSubExtensions = [
  Node.create({ name: 'doc', topNode: true, content: 'block+' }),
  DocText,
  DocHardBreak,
  TextboxParagraph,
  BoldMark,
  ItalicMark,
  UnderlineMark,
  StrikeMark,
  LinkMark,
  // research-report sidebars keep PAGE/date/REF fields inside textbox tables;
  // without these marks the whole box fails to load into the sub-editor
  RefFieldMark,
  InstrFieldMark,
  TextStyleMark,
  CommentMark,
  UndoRedo,
]

// ---- find & replace highlighting ----

export interface SearchHighlight {
  ranges: Array<{ from: number; to: number }>
  activeIndex: number
}

export const editorExtensions = [
  DocDocument,
  DocText,
  DocHardBreak,
  DocNoteRef,
  DocXeMark,
  DocRuby,
  DocInlineMath,
  DocParagraph,
  DocHeading,
  DocListItem,
  DocTable,
  DocTableRow,
  DocTableCell,
  DocTableHeader,
  DocNestedTable,
  DocProtected,
  BoldMark,
  ItalicMark,
  UnderlineMark,
  StrikeMark,
  LinkMark,
  RefFieldMark,
  InstrFieldMark,
  RprChangeMark,
  TextStyleMark,
  CommentMark,
  InsMark,
  DelMark,
  UndoRedo,
  SearchHighlightExtension,
  PendingCommentHighlightExtension,
  ResolvedCommentsExtension,
  NativeTableSupport,
  TrackChangesExtension,
  LineFactorExtension,
  ListNumberingExtension,
  PaginationGapsExtension,
  TabStopExtension,
  DropCapExtension,
  SdtExtension,
  MoveRevisionExtension,
  PPrChangeExtension,
  RevisionOriginalExtension,
  AutoDirectionExtension,
]
