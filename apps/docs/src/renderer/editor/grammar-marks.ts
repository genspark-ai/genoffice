/**
 * Grammar findings shown on the page as Word shows them: a blue double
 * underline. They are drawn with the CSS Custom Highlight API, so the page's
 * DOM is not touched: nothing reaches a saved file, the layout does not move
 * and Chromium keeps its spelling markers.
 */
import { Extension } from '@tiptap/core'
import type { Editor } from '@tiptap/core'
import type { Node as PmNode } from '@tiptap/pm/model'
import { Plugin, PluginKey, type EditorState, type Transaction } from '@tiptap/pm/state'
import type { EditorView } from '@tiptap/pm/view'
import { liveTextPieces, type GrammarFinding } from '../ai/grammar'
import { grammarHidden } from '../grammar-pref'
import { pageHighlight } from './page-highlight'

export interface GrammarMarksState {
  findings: readonly GrammarFinding[]
  hidden: boolean
  ignored: ReadonlySet<string>
}

type GrammarMeta =
  | { set: GrammarFinding[] }
  | { remove: string; ignore?: boolean }
  | { hidden: boolean }
  | { resetIgnored: true }
  | { accept: string }

export const grammarKey = new PluginKey<GrammarMarksState>('grammarMarks')
export const GRAMMAR_HIGHLIGHT = 'docs-grammar'

const quoteAt = (doc: PmNode, from: number, to: number) =>
  liveTextPieces(doc, from, to)
    .map((p) => p.text)
    .join('')

const blockText = (block: PmNode) => block.textBetween(0, block.content.size, undefined, '￼')

/** an edited paragraph loses its marks; accepting one suggestion keeps the paragraph's other marks that still read the same */
function carryFindings(
  findings: readonly GrammarFinding[],
  tr: Transaction,
  oldDoc: PmNode,
  accepted: string | null,
): GrammarFinding[] {
  const out: GrammarFinding[] = []
  for (const finding of findings) {
    if (finding.id === accepted) continue
    const from = tr.mapping.map(finding.from, 1)
    const to = tr.mapping.map(finding.to, -1)
    if (from >= to) continue
    const before = oldDoc.resolve(finding.from).parent
    const after = tr.doc.resolve(from).parent
    const same =
      before === after ||
      (accepted !== null
        ? quoteAt(tr.doc, from, to) === finding.quote
        : blockText(before) === blockText(after))
    if (same)
      out.push(from === finding.from && to === finding.to ? finding : { ...finding, from, to })
  }
  return out
}

function metaOf(tr: Transaction): GrammarMeta | undefined {
  const own = tr.getMeta(grammarKey) as GrammarMeta | undefined
  if (own) return own
  // Track Changes rewrites an accepted suggestion in a follow-up transaction
  const source = tr.getMeta('appendedTransaction') as Transaction | undefined
  const inherited = source?.getMeta(grammarKey) as GrammarMeta | undefined
  return inherited && 'accept' in inherited ? inherited : undefined
}

function apply(tr: Transaction, prev: GrammarMarksState, oldState: EditorState): GrammarMarksState {
  const meta = metaOf(tr)
  let next = prev
  if (tr.docChanged && prev.findings.length > 0) {
    const accepted = meta && 'accept' in meta ? meta.accept : null
    next = { ...next, findings: carryFindings(prev.findings, tr, oldState.doc, accepted) }
  }
  if (!meta) return next
  if ('set' in meta)
    return {
      ...next,
      // the document may have moved on while the model was answering
      findings: meta.set.filter(
        (f) => f.to <= tr.doc.content.size && quoteAt(tr.doc, f.from, f.to) === f.quote,
      ),
    }
  if ('remove' in meta) {
    const gone = next.findings.find((f) => f.id === meta.remove)
    if (!gone) return next
    return {
      ...next,
      findings: next.findings.filter((f) => f !== gone),
      ignored: meta.ignore ? new Set([...next.ignored, gone.key]) : next.ignored,
    }
  }
  if ('hidden' in meta) return { ...next, hidden: meta.hidden }
  if ('resetIgnored' in meta) return { ...next, ignored: new Set() }
  return next
}

function drawMarks(view: EditorView, drawn: AbstractRange[]): AbstractRange[] {
  const highlight = pageHighlight(GRAMMAR_HIGHLIGHT)
  if (!highlight) return []
  for (const range of drawn) highlight.delete(range)
  const state = grammarKey.getState(view.state)
  if (!state || state.hidden) return []
  const next: AbstractRange[] = []
  for (const finding of state.findings) {
    try {
      const start = view.domAtPos(finding.from, 1)
      const end = view.domAtPos(finding.to, -1)
      const range = document.createRange()
      range.setStart(start.node, start.offset)
      range.setEnd(end.node, end.offset)
      highlight.add(range)
      next.push(range)
    } catch {
      // a position the view has no DOM for yet is drawn on the next update
    }
  }
  return next
}

export const GrammarMarksExtension = Extension.create({
  name: 'grammarMarks',

  addProseMirrorPlugins() {
    return [
      new Plugin<GrammarMarksState>({
        key: grammarKey,
        state: {
          init: () => ({ findings: [], hidden: grammarHidden(), ignored: new Set<string>() }),
          apply,
        },
        view(view) {
          let drawn = drawMarks(view, [])
          return {
            update(updated) {
              // the view re-renders text nodes without a state change of ours, so the ranges are rebuilt each time
              if (drawn.length > 0 || grammarKey.getState(updated.state)?.findings.length)
                drawn = drawMarks(updated, drawn)
            },
            destroy() {
              const highlight = pageHighlight(GRAMMAR_HIGHLIGHT)
              for (const range of drawn) highlight?.delete(range)
            },
          }
        },
      }),
    ]
  },
})

const send = (editor: Editor, meta: GrammarMeta) =>
  editor.view.dispatch(editor.state.tr.setMeta(grammarKey, meta).setMeta('addToHistory', false))

export const grammarState = (state: EditorState): GrammarMarksState | undefined =>
  grammarKey.getState(state)

export function setGrammarFindings(editor: Editor, findings: GrammarFinding[]): void {
  send(editor, { set: findings })
}

export function setGrammarMarksHidden(editor: Editor, hidden: boolean): void {
  if (grammarState(editor.state)?.hidden !== hidden) send(editor, { hidden })
}

/** Ignore Once: the finding goes, and the next check does not bring it back */
export function ignoreGrammarFinding(editor: Editor, id: string): void {
  send(editor, { remove: id, ignore: true })
}

export function resetIgnoredGrammar(editor: Editor): void {
  if (grammarState(editor.state)?.ignored.size) send(editor, { resetIgnored: true })
}

/** the finding whose mark is under a position; none while the marks are hidden */
export function grammarFindingAt(state: EditorState, pos: number): GrammarFinding | null {
  const marks = grammarState(state)
  if (!marks || marks.hidden) return null
  return marks.findings.find((f) => f.from <= pos && pos <= f.to) ?? null
}

/** Replaces the finding's words with a suggestion as one undo step; inline objects and deleted text inside it stay. */
export function applyGrammarSuggestion(editor: Editor, id: string, replacement: string): boolean {
  const finding = grammarState(editor.state)?.findings.find((f) => f.id === id)
  if (!finding || !editor.isEditable) return false
  const pieces = liveTextPieces(editor.state.doc, finding.from, finding.to)
  if (pieces.map((p) => p.text).join('') !== finding.quote) {
    send(editor, { remove: id })
    return false
  }
  return editor
    .chain()
    .focus()
    .command(({ tr }) => {
      for (const piece of pieces.slice(1).reverse()) tr.delete(piece.from, piece.to)
      tr.insertText(replacement, pieces[0].from, pieces[0].to)
      tr.setMeta(grammarKey, { accept: id })
      return true
    })
    .run()
}
