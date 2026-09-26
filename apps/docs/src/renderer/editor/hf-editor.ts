/**
 * Rich header/footer editing: a nested TipTap editor (textbox sub-editor
 * schema + tab stops) hosts the part's text paragraphs, so ribbon formatting
 * routes into it like a shape's text. PAGE / NUMPAGES field sentinels become
 * atom nodes showing the live number; layout-table rows and floating boxes
 * stay display-only and are spliced back around the edited paragraphs.
 */
import { Editor, Extension, Node } from '@tiptap/core'
import { Plugin, PluginKey } from '@tiptap/pm/state'
import { Decoration, DecorationSet } from '@tiptap/pm/view'
import {
  PAGE_MARK,
  TOTAL_PAGES_MARK,
  type HeaderFooter,
  type HfParagraph,
  type Run,
} from '@genoffice/docx-engine'
import { dropActiveSubEditor, notifySubEditorState, setActiveSubEditor } from './active-editor'
import { inlineToRuns, runsToInline, type PmNode } from './convert'
import { TextboxParagraph, textboxSubExtensions } from './extensions'
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

export const hfSubExtensions = [
  ...textboxSubExtensions.filter((ext) => ext.name !== 'docParagraph'),
  HfParagraphNode,
  HfFieldNode,
  HfFieldNumbersExtension,
]

/** paragraphs the editor takes over; rows and floating-box content stay display-only */
export function hfEditablePara(para: HfParagraph): boolean {
  return !para.cells && !para.box && !para.boxAnchored
}

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

export function hfDocJson(value: HeaderFooter): PmNode {
  const paras = hfParasOf(value)
  const content: PmNode[] = []
  paras.forEach((para, i) => {
    if (!hfEditablePara(para)) return
    const attrs: Record<string, unknown> = { hfIndex: i }
    for (const key of HF_PARA_KEYS) {
      const v = (para as unknown as Record<string, unknown>)[key]
      if (v !== undefined && v !== null) attrs[key] = v
    }
    if (para.tabStops?.length) attrs.tabStops = JSON.stringify(para.tabStops)
    content.push({ type: 'docParagraph', attrs, content: fieldSplit(runsToInline(para.runs)) })
  })
  if (content.length === 0) content.push({ type: 'docParagraph', attrs: { hfIndex: null } })
  return { type: 'doc', content }
}

function paraFromNode(node: PmNode, template: HfParagraph): HfParagraph {
  const para: HfParagraph = { ...template, runs: inlineToRuns(fieldJoin(node.content ?? [])) }
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
export function hfValueFromDoc(doc: PmNode, base: HeaderFooter): HeaderFooter {
  const paras = hfParasOf(base)
  const editable = paras.filter(hfEditablePara)
  const fallback: HfParagraph = editable[editable.length - 1] ?? { align: 'center', runs: [] }
  const edited: HfParagraph[] = (doc.content ?? []).map((node) => {
    const idx = node.attrs?.hfIndex
    const template =
      typeof idx === 'number' && paras[idx] && hfEditablePara(paras[idx]) ? paras[idx] : fallback
    return paraFromNode(node, template)
  })
  const next: HfParagraph[] = []
  let ei = 0
  for (const p of paras) {
    if (!hfEditablePara(p)) next.push(p)
    else if (ei < edited.length) next.push(edited[ei++])
  }
  next.push(...edited.slice(ei))
  const text = edited.map((p) => p.runs.map((r: Run) => r.text).join('')).join('')
  // fields are atoms here, so the flag follows the marks: a removed field must not
  // come back as the save path's appended page-number paragraph
  const pageNumber = next.some((p) => p.runs.some((r) => r.text.includes(PAGE_MARK)))
  return { ...base, text, paras: next, pageNumber }
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
      handleKeyDown: (_view, event) => {
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
  const initialSig = JSON.stringify(editor.getJSON())

  const changed = (): HeaderFooter | null => {
    const doc = editor.getJSON() as PmNode
    return JSON.stringify(doc) === initialSig ? null : hfValueFromDoc(doc, opts.value)
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
