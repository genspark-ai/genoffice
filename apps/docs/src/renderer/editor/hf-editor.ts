/**
 * Rich header/footer editing: a nested TipTap editor (textbox sub-editor
 * schema + tab stops) hosts the part's text paragraphs, so ribbon formatting
 * routes into it like a shape's text. PAGE / NUMPAGES field sentinels become
 * atom nodes showing the live number; layout-table rows and floating boxes
 * stay display-only and are spliced back around the edited paragraphs.
 */
import { Editor, Extension, Node } from '@tiptap/core'
import { Plugin, PluginKey, TextSelection, type EditorState } from '@tiptap/pm/state'
import type { ResolvedPos } from '@tiptap/pm/model'
import type { EditorView } from '@tiptap/pm/view'
import { Decoration, DecorationSet } from '@tiptap/pm/view'
import {
  NEW_PICTURE_TOKEN_PREFIX,
  PAGE_MARK,
  TOTAL_PAGES_MARK,
  type HeaderFooter,
  type HfCellParaProps,
  type HfParagraph,
  type HfTableCell,
  type HfTableRow,
  type Run,
} from '@genoffice/docx-engine'
import { dropActiveSubEditor, notifySubEditorState, setActiveSubEditor } from './active-editor'
import { inlineToRuns, runsToInline, type PmNode } from './convert'
import { setDkBackground, setDkBackgroundImage, setDkBorder } from './dark-page'
import { TextboxParagraph, textboxSubExtensions } from './extensions'
import { hfCellGeometry, hfRowStyle } from './hf-dom'
import { hfParasOf } from './hf-text'

export type HfFieldKind = 'PAGE' | 'NUMPAGES'

interface HfFieldNumbers {
  page: string
  total: string
}

/** PAGE / NUMPAGES field: one atom per sentinel; the number is drawn from a decoration attribute */
const HfFieldNode = Node.create({
  name: 'hfField',
  group: 'inline',
  inline: true,
  atom: true,
  selectable: true,
  marks: '_',
  addAttributes() {
    return { kind: { default: 'PAGE' as HfFieldKind } }
  },
  parseHTML() {
    return [
      {
        tag: 'span[data-hf-field]',
        getAttrs: (el) => ({ kind: (el as HTMLElement).getAttribute('data-hf-field') }),
      },
    ]
  },
  renderHTML({ node }) {
    return [
      'span',
      { class: 'hf-field', 'data-hf-field': node.attrs.kind as string, contenteditable: 'false' },
    ]
  },
})

const hfFieldNumbersKey = new PluginKey<HfFieldNumbers>('hfFieldNumbers')

/** live page / total numbers on the field atoms (set per mount through a transaction meta) */
const HfFieldNumbersExtension = Extension.create({
  name: 'hfFieldNumbers',
  addProseMirrorPlugins() {
    return [
      new Plugin<HfFieldNumbers>({
        key: hfFieldNumbersKey,
        state: {
          init: () => ({ page: '1', total: '1' }),
          apply: (tr, prev) =>
            (tr.getMeta(hfFieldNumbersKey) as HfFieldNumbers | undefined) ?? prev,
        },
        props: {
          decorations(state) {
            const nums = hfFieldNumbersKey.getState(state)!
            const decos: Decoration[] = []
            state.doc.descendants((node, pos) => {
              if (node.type.name !== 'hfField') return
              const num = node.attrs.kind === 'NUMPAGES' ? nums.total : nums.page
              decos.push(Decoration.node(pos, pos + node.nodeSize, { 'data-hf-num': num }))
            })
            return DecorationSet.create(state.doc, decos)
          },
        },
      }),
    ]
  },
})

/** paragraph attrs carried both ways between HfParagraph and the editor node */
const HF_PARA_KEYS = [
  'styleId',
  'align',
  'lineSpacing',
  'lineRule',
  'lineRawTwips',
  'snapToGrid',
  'indentLeft',
  'indentRight',
  'indentFirstLine',
  'indentFirstLineChars',
  'spaceBefore',
  'spaceAfter',
  'spaceBeforeAuto',
  'spaceAfterAuto',
  'shadingFill',
  'shadingDisplay',
  'shadingClear',
  'borders',
  'bidi',
] as const

const HfParagraphNode = TextboxParagraph.extend({
  addAttributes() {
    return {
      ...this.parent?.(),
      bidi: { default: null as boolean | null },
      tabStops: { default: null as string | null },
      /** index of the source HfParagraph (its unmodeled fields survive the round trip) */
      hfIndex: { default: null as number | null, rendered: false },
    }
  },
  renderHTML(props) {
    const spec = this.parent?.(props) as [string, Record<string, string>, ...unknown[]]
    const attrs = spec[1]
    if (props.node.attrs.tabStops) attrs['data-tab-stops'] = String(props.node.attrs.tabStops)
    if (props.node.attrs.bidi) attrs.dir = 'rtl'
    return spec
  },
})

/** layout-table row: flex columns of cells, each a stack of paragraphs. The
 *  geometry stays on the node (display only); isolating cells keep Backspace /
 *  Enter from merging text across cells, like ProseMirror tables. */
const HfRowNode = Node.create({
  name: 'hfRow',
  group: 'block',
  content: 'hfCell+',
  isolating: true,
  defining: true,
  addAttributes() {
    return {
      hfIndex: { default: null as number | null, rendered: false },
      row: { default: null as HfTableRow | null, rendered: false },
    }
  },
  parseHTML() {
    return [{ tag: 'div.page-hf-row' }]
  },
  renderHTML() {
    return ['div', { class: 'page-hf-para page-hf-row' }, 0]
  },
  addNodeView() {
    return ({ node }) => {
      const dom = document.createElement('div')
      dom.className = 'page-hf-para page-hf-row'
      Object.assign(dom.style, hfRowStyle((node.attrs.row as HfTableRow | null) ?? undefined))
      return { dom, contentDOM: dom }
    }
  },
})

/** caret at the first (side -1) or last (side 1) edge of a layout-table cell's text:
 *  no editor command joins across cells (isolating), so the keystroke must be
 *  swallowed before the browser's native deletion merges the cell elements */
export function atHfCellEdge(state: EditorState, side: -1 | 1, at?: ResolvedPos): boolean {
  if (!state.selection.empty && !at) return false
  const $from = at ?? state.selection.$from
  if (!$from.parent.isTextblock) return false
  const atEdge =
    side < 0 ? $from.parentOffset === 0 : $from.parentOffset === $from.parent.content.size
  if (!atEdge) return false
  for (let d = $from.depth - 1; d > 0; d--) {
    if ($from.node(d).type.name !== 'hfCell') continue
    const index = $from.index(d)
    return side < 0 ? index === 0 : index === $from.node(d).childCount - 1
  }
  return false
}

/** Home / End inside a layout-table cell, moved by the editor: Blink's own
 *  line-edge move after a DOM update in these cells reveals the caret by
 *  scrolling the document to its end */
function moveToCellLineEdge(view: EditorView, side: -1 | 1): boolean {
  const $from = domCaret(view) ?? view.state.selection.$from
  if (!$from.parent.isTextblock) return false
  let inCell = false
  for (let d = $from.depth - 1; d > 0; d--) inCell ||= $from.node(d).type.name === 'hfCell'
  if (!inCell) return false
  const start = $from.start()
  const end = $from.end()
  let pos = side < 0 ? start : end
  try {
    const caret = view.coordsAtPos($from.pos)
    const first = view.coordsAtPos(start)
    const last = view.coordsAtPos(end)
    const sameLine = (a: { top: number }, b: { top: number }) => Math.abs(a.top - b.top) < 2
    if (!sameLine(first, last)) {
      const para = view.domAtPos(start).node as HTMLElement
      const box = (para.nodeType === 1 ? para : para.parentElement!).getBoundingClientRect()
      const hit = view.posAtCoords({
        left: side < 0 ? box.left + 1 : box.right - 1,
        top: (caret.top + caret.bottom) / 2,
      })
      if (hit && hit.pos >= start && hit.pos <= end) pos = hit.pos
    }
  } catch {
    // no layout (tests): paragraph edge
  }
  view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, pos)))
  return true
}

/** the caret as the browser has it: right after a caret-moving key the DOM
 *  selection can be ahead of the editor state's (synced on the next flush) */
function domCaret(view: EditorView): ResolvedPos | undefined {
  const sel = view.dom.ownerDocument.getSelection()
  if (!sel?.anchorNode || !sel.isCollapsed || !view.dom.contains(sel.anchorNode)) return undefined
  try {
    return view.state.doc.resolve(view.posAtDOM(sel.anchorNode, sel.anchorOffset))
  } catch {
    return undefined
  }
}

const HF_CELL_GEOM_KEYS = [
  'widthPct',
  'widthTwips',
  'fill',
  'borders',
  'vAlign',
  'marTwips',
  'align',
] as const

type HfCellGeom = Pick<HfTableCell, (typeof HF_CELL_GEOM_KEYS)[number]>

const HfCellNode = Node.create({
  name: 'hfCell',
  content: 'docParagraph+',
  isolating: true,
  defining: true,
  addAttributes() {
    return { cell: { default: null as HfCellGeom | null, rendered: false } }
  },
  parseHTML() {
    return [{ tag: 'div.page-hf-cell' }]
  },
  renderHTML() {
    return ['div', { class: 'page-hf-cell' }, 0]
  },
  addNodeView() {
    return ({ node }) => {
      const dom = document.createElement('div')
      dom.className = 'page-hf-cell'
      const cell = (node.attrs.cell as HfCellGeom | null) ?? {}
      const geom = hfCellGeometry({ paras: [], ...cell })
      Object.assign(dom.style, geom.style)
      if (cell.fill) setDkBackground(dom, `#${cell.fill}`)
      for (const [side, css] of Object.entries(geom.borders)) {
        setDkBorder(dom, side as 't' | 'b' | 'l' | 'r', css)
      }
      if (geom.diagonals) setDkBackgroundImage(dom, geom.diagonals)
      return { dom, contentDOM: dom }
    }
  },
})

export const hfSubExtensions = [
  ...textboxSubExtensions.filter((ext) => ext.name !== 'docParagraph'),
  HfParagraphNode,
  HfRowNode,
  HfCellNode,
  HfFieldNode,
  HfFieldNumbersExtension,
]

/** paragraphs and table rows the editor takes over; floating-box content and
 *  rows flattened out of nested tables stay display-only */
export function hfEditablePara(para: HfParagraph): boolean {
  if (para.box || para.boxAnchored) return false
  return !para.cells?.some((cell) => cell.locked)
}

/** cell paragraph props carried both ways between HfCellParaProps and the editor node */
const HF_CELL_PARA_KEYS = [
  'styleId',
  'align',
  'indentLeft',
  'indentFirstLine',
  'spaceBefore',
  'spaceAfter',
  'lineRule',
  'lineRawTwips',
  'lineSpacing',
] as const

function fieldSplit(nodes: PmNode[]): PmNode[] {
  const out: PmNode[] = []
  for (const node of nodes) {
    if (node.type !== 'text' || !node.text) {
      out.push(node)
      continue
    }
    const marks = node.marks ? { marks: node.marks } : {}
    for (const seg of node.text.split(new RegExp(`([${PAGE_MARK}${TOTAL_PAGES_MARK}])`))) {
      if (seg === PAGE_MARK) out.push({ type: 'hfField', attrs: { kind: 'PAGE' }, ...marks })
      else if (seg === TOTAL_PAGES_MARK)
        out.push({ type: 'hfField', attrs: { kind: 'NUMPAGES' }, ...marks })
      else if (seg) out.push({ type: 'text', text: seg, ...marks })
    }
  }
  return out
}

function fieldJoin(nodes: PmNode[]): PmNode[] {
  return nodes.map((node) =>
    node.type === 'hfField'
      ? {
          type: 'text',
          text: node.attrs?.kind === 'NUMPAGES' ? TOTAL_PAGES_MARK : PAGE_MARK,
          ...(node.marks ? { marks: node.marks } : {}),
        }
      : node,
  )
}

function cellParaNode(runs: Run[], props: HfCellParaProps | undefined): PmNode {
  const attrs: Record<string, unknown> = {}
  for (const key of HF_CELL_PARA_KEYS) {
    const v = props?.[key]
    if (v !== undefined && v !== null) attrs[key] = v
  }
  if (props?.tabStops?.length) attrs.tabStops = JSON.stringify(props.tabStops)
  return { type: 'docParagraph', attrs, content: fieldSplit(runsToInline(runs)) }
}

function cellNode(cell: HfTableCell): PmNode {
  const geom: Record<string, unknown> = {}
  for (const key of HF_CELL_GEOM_KEYS) {
    if (cell[key] !== undefined) geom[key] = cell[key]
  }
  const paras = cell.paras.length > 0 ? cell.paras : [[]]
  return {
    type: 'hfCell',
    attrs: { cell: geom },
    content: paras.map((runs, k) => cellParaNode(runs, cell.paraProps?.[k])),
  }
}

export function hfDocJson(value: HeaderFooter): PmNode {
  const paras = hfParasOf(value)
  const content: PmNode[] = []
  paras.forEach((para, i) => {
    if (!hfEditablePara(para)) return
    if (para.cells) {
      content.push({
        type: 'hfRow',
        attrs: { hfIndex: i, row: para.row ?? null },
        content: para.cells.map(cellNode),
      })
      return
    }
    const attrs: Record<string, unknown> = { hfIndex: i }
    for (const key of HF_PARA_KEYS) {
      const v = (para as unknown as Record<string, unknown>)[key]
      if (v !== undefined && v !== null) attrs[key] = v
    }
    if (para.tabStops?.length) attrs.tabStops = JSON.stringify(para.tabStops)
    content.push({
      type: 'docParagraph',
      attrs,
      content: fieldSplit(runsToInline(para.runs)),
    })
  })
  if (content.length === 0) content.push({ type: 'docParagraph', attrs: { hfIndex: null } })
  return { type: 'doc', content }
}

/** node JSON with the schema's null defaults dropped, so an untouched cell compares equal */
function canonical(node: PmNode): unknown {
  const attrs = Object.fromEntries(
    Object.entries(node.attrs ?? {}).filter(([, v]) => v !== null && v !== undefined),
  )
  return {
    type: node.type,
    ...(Object.keys(attrs).length > 0 ? { attrs } : {}),
    ...(node.text !== undefined ? { text: node.text } : {}),
    ...(node.marks?.length ? { marks: node.marks } : {}),
    ...(node.content?.length ? { content: node.content.map(canonical) } : {}),
  }
}

/** `initial`: the same cell as the editor loaded it (schema-normalized), so
 *  an untouched cell compares equal despite attribute defaults */
function cellFromNode(node: PmNode, template: HfTableCell, initial?: PmNode): HfTableCell {
  const paraNodes = node.content ?? []
  const paras = paraNodes.map((p) => inlineToRuns(fieldJoin(p.content ?? [])))
  const paraProps = paraNodes.map((p, k): HfCellParaProps | undefined => {
    const props: Record<string, unknown> = { ...template.paraProps?.[k] }
    for (const key of HF_CELL_PARA_KEYS) {
      const v = p.attrs?.[key]
      if (v !== null && v !== undefined) props[key] = v
      else delete props[key]
    }
    if (typeof p.attrs?.tabStops === 'string') {
      try {
        const parsed: unknown = JSON.parse(p.attrs.tabStops)
        if (Array.isArray(parsed) && parsed.length > 0) props.tabStops = parsed
        else delete props.tabStops
      } catch {
        delete props.tabStops
      }
    } else delete props.tabStops
    return Object.keys(props).length > 0 ? (props as HfCellParaProps) : undefined
  })
  const changed =
    JSON.stringify(canonical(node)) !== JSON.stringify(canonical(initial ?? cellNode(template)))
  const propsOf = (props: HfCellParaProps | undefined) => JSON.stringify(props ?? {})
  const rewriteFormat = paraProps
    .map((props, k) => (propsOf(props) !== propsOf(template.paraProps?.[k]) ? k : -1))
    .filter((k) => k >= 0)
  return {
    ...template,
    paras,
    ...(paraProps.some(Boolean) ? { paraProps } : {}),
    ...(changed || template.rewrite ? { rewrite: true as const } : {}),
    ...(rewriteFormat.length > 0 || template.rewriteFormat
      ? {
          rewriteFormat: [...new Set([...(template.rewriteFormat ?? []), ...rewriteFormat])],
          priorParaProps: template.priorParaProps ?? template.paraProps ?? [],
        }
      : {}),
  }
}

function rowFromNode(node: PmNode, template: HfParagraph, initial?: PmNode): HfParagraph {
  const cells = (node.content ?? []).map((cellNodeJson, j) =>
    cellFromNode(cellNodeJson, template.cells?.[j] ?? { paras: [] }, initial?.content?.[j]),
  )
  return { ...template, runs: [], cells }
}

/** bytes of pictures inserted this session, keyed by the token their drawing carries */
function newPicturesOf(paras: HfParagraph[], prev?: HeaderFooter['newPictures']) {
  const out: NonNullable<HeaderFooter['newPictures']> = {}
  const runsOf = (p: HfParagraph) => [p.runs, ...(p.cells?.flatMap((c) => c.paras) ?? [])]
  for (const run of paras.flatMap(runsOf).flat()) {
    const xml = run.image?.xml ?? ''
    if (!xml.includes(NEW_PICTURE_TOKEN_PREFIX)) continue
    const token = new RegExp(`r:embed="(${NEW_PICTURE_TOKEN_PREFIX}[0-9]+)"`).exec(xml)?.[1]
    if (!token) continue
    const bytes = /^data:(image\/(?:png|jpeg|gif));base64,(.+)$/s.exec(run.image?.dataUrl ?? '')
    if (bytes) out[token] = { base64: bytes[2], mime: bytes[1] as 'image/png' }
    else if (prev?.[token]) out[token] = prev[token]
  }
  return Object.keys(out).length > 0 ? out : undefined
}

function paraFromNode(node: PmNode, template: HfParagraph): HfParagraph {
  const para: HfParagraph = {
    ...template,
    runs: inlineToRuns(fieldJoin(node.content ?? [])),
  }
  delete para.cells
  delete para.row
  delete para.box
  delete para.boxAnchored
  const attrs = node.attrs ?? {}
  for (const key of HF_PARA_KEYS) {
    const v = attrs[key]
    if (v !== null && v !== undefined) Object.assign(para, { [key]: v })
    else delete (para as unknown as Record<string, unknown>)[key]
  }
  if (typeof attrs.tabStops === 'string') {
    try {
      const parsed: unknown = JSON.parse(attrs.tabStops)
      if (Array.isArray(parsed) && parsed.length > 0) para.tabStops = parsed
      else delete para.tabStops
    } catch {
      delete para.tabStops
    }
  } else delete para.tabStops
  // a blank line takes its own mark size; the old value would size a line that no longer exists
  if (para.runs.length > 0) delete para.emptyRunSizeHalfPoints
  return para
}

/** editor doc → part paragraphs: rows/boxes go back to their original slots, like applyHfText */
export function hfValueFromDoc(doc: PmNode, base: HeaderFooter, initial?: PmNode): HeaderFooter {
  const paras = hfParasOf(base)
  const initialRow = (hfIndex: unknown) =>
    initial?.content?.find((n) => n.type === 'hfRow' && n.attrs?.hfIndex === hfIndex)
  const editable = paras.filter(hfEditablePara)
  const fallback: HfParagraph = editable[editable.length - 1] ?? { align: 'center', runs: [] }
  const edited: HfParagraph[] = (doc.content ?? []).map((node) => {
    const idx = node.attrs?.hfIndex
    const source =
      typeof idx === 'number' && paras[idx] && hfEditablePara(paras[idx]) ? paras[idx] : null
    if (node.type === 'hfRow')
      return rowFromNode(node, source?.cells ? source : { runs: [] }, initialRow(idx))
    return paraFromNode(node, source && !source.cells ? source : fallback)
  })
  const next: HfParagraph[] = []
  let ei = 0
  for (const p of paras) {
    if (!hfEditablePara(p)) next.push(p)
    else if (ei < edited.length) next.push(edited[ei++])
  }
  next.push(...edited.slice(ei))
  const text = edited
    .filter((p) => !p.cells)
    .map((p) => p.runs.map((r: Run) => r.text).join(''))
    .join('')
  // fields are atoms here, so the flag follows the marks: a removed field must not
  // come back as the save path's appended page-number paragraph
  const allRuns = (p: HfParagraph) => [
    ...p.runs,
    ...(p.cells?.flatMap((c) => c.paras.flat()) ?? []),
  ]
  const pageNumber = next.some((p) => allRuns(p).some((r) => r.text.includes(PAGE_MARK)))
  const newPictures = newPicturesOf(next, base.newPictures)
  const { newPictures: _prev, ...rest } = base
  return { ...rest, text, paras: next, pageNumber, ...(newPictures ? { newPictures } : {}) }
}

/** Picture from File: an inline picture at the caret, fitted to the cell or strip width */
export function insertHfPicture(
  editor: Editor,
  picture: { dataUrl: string; widthPx: number; heightPx: number },
): void {
  const { $from } = editor.state.selection
  let hostWidth = editor.view.dom.clientWidth
  for (let d = $from.depth; d > 0; d--) {
    if ($from.node(d).type.name !== 'hfCell') continue
    const dom = editor.view.nodeDOM($from.before(d)) as HTMLElement | null
    if (dom) {
      const cs = getComputedStyle(dom)
      hostWidth = dom.clientWidth - parseFloat(cs.paddingLeft) - parseFloat(cs.paddingRight)
    }
    break
  }
  const scale = hostWidth > 0 ? Math.min(1, hostWidth / picture.widthPx) : 1
  // the strip is already on screen; the default focus scroll jumps the page
  editor
    .chain()
    .focus(undefined, { scrollIntoView: false })
    .insertContent({
      type: 'docInlineImage',
      attrs: {
        dataUrl: picture.dataUrl,
        widthPx: Math.max(1, Math.round(picture.widthPx * scale)),
        heightPx: Math.max(1, Math.round(picture.heightPx * scale)),
      },
    })
    .run()
}

export function insertHfField(editor: Editor, kind: HfFieldKind): void {
  editor.chain().focus().insertContent({ type: 'hfField', attrs: { kind } }).run()
}

/** DATE / TIME / FILENAME: cached result text under the instrField mark, like the body */
export function insertHfInstrField(editor: Editor, instr: string, value: string): void {
  editor
    .chain()
    .focus()
    .insertContent({
      type: 'text',
      text: value || ' ',
      marks: [{ type: 'instrField', attrs: { instr } }],
    })
    .unsetMark('instrField')
    .run()
}

export const HF_COMMIT_EVENT = 'ai-docs-commit-tables'

/** focus moving here keeps the header editor open (ribbon formatting acts on it, Word-style) */
const KEEP_OPEN_SELECTOR = '.ribbon, [data-rb-panel], .modal, .modal-backdrop, .ctx-menu'

export interface HfEditorHandle {
  editor: Editor
  /** current content when it differs from the mounted value */
  changed(): HeaderFooter | null
  /** commit (when changed), destroy the editor and report the exit */
  exit(): void
  /** destroy the editor without committing (the strip's ownership changed underneath it) */
  discard(): void
}

export function mountHfEditor(
  host: HTMLElement,
  opts: {
    value: HeaderFooter
    pageNo: string
    pageTotal: string
    spellcheck: boolean
    onCommit: (next: HeaderFooter) => void
    onExit: () => void
  },
): HfEditorHandle {
  const editor: Editor = new Editor({
    element: host,
    extensions: hfSubExtensions,
    content: hfDocJson(opts.value),
    editorProps: {
      attributes: { class: 'page-hf-editor', spellcheck: opts.spellcheck ? 'true' : 'false' },
      handleKeyDown: (view, event) => {
        if (
          (event.key === 'End' || event.key === 'Home') &&
          !event.shiftKey &&
          !event.altKey &&
          !event.ctrlKey &&
          !event.metaKey
        ) {
          if (!moveToCellLineEdge(view, event.key === 'End' ? 1 : -1)) return false
          event.preventDefault()
          return true
        }
        if (event.key === 'Backspace' || event.key === 'Delete') {
          if (!atHfCellEdge(view.state, event.key === 'Backspace' ? -1 : 1, domCaret(view)))
            return false
          event.preventDefault()
          return true
        }
        if (event.key !== 'Escape') return false
        event.preventDefault()
        queueMicrotask(exit)
        return true
      },
    },
    onFocus: () => setActiveSubEditor(editor),
    onTransaction: ({ transaction }) => notifySubEditorState(transaction.docChanged),
  })
  editor.view.dispatch(
    editor.state.tr.setMeta(hfFieldNumbersKey, { page: opts.pageNo, total: opts.pageTotal }),
  )
  // compared after schema normalization, so untouched null attrs never read as an edit
  const initialDoc = editor.getJSON() as PmNode
  const initialSig = JSON.stringify(initialDoc)

  const changed = (): HeaderFooter | null => {
    const doc = editor.getJSON() as PmNode
    return JSON.stringify(doc) === initialSig ? null : hfValueFromDoc(doc, opts.value, initialDoc)
  }
  let exited = false
  const onCommitEvent = () => {
    const next = changed()
    if (next) opts.onCommit(next)
    exit()
  }
  const onFocusOut = (e: FocusEvent) => {
    const next = e.relatedTarget as HTMLElement | null
    if (next && (host.contains(next) || next.closest(KEEP_OPEN_SELECTOR))) return
    if (next) return exit()
    // no relatedTarget: either a real blur or Blink blurring the strip a
    // re-pagination is about to detach (dispatched synchronously, host still
    // connected). Decide once the removal has run: a detached host is re-hosted
    // or exited by its owner
    queueMicrotask(() => {
      if (exited || !host.isConnected) return
      if (document.activeElement && host.contains(document.activeElement)) return
      exit()
    })
  }
  const close = (commit: boolean) => {
    if (exited) return
    exited = true
    host.removeEventListener('focusout', onFocusOut)
    window.removeEventListener(HF_COMMIT_EVENT, onCommitEvent)
    const next = commit ? changed() : null
    dropActiveSubEditor(editor)
    editor.destroy()
    if (next) opts.onCommit(next)
    opts.onExit()
  }
  const exit = () => close(true)
  host.addEventListener('focusout', onFocusOut)
  window.addEventListener(HF_COMMIT_EVENT, onCommitEvent)
  editor.commands.focus('end')
  return { editor, changed, exit, discard: () => close(false) }
}
