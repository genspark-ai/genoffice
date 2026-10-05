import { forwardRef, useEffect, useImperativeHandle, useRef } from 'react'
import { Annotation, Compartment, EditorState } from '@codemirror/state'
import {
  EditorView,
  keymap,
  lineNumbers,
  highlightActiveLine,
  drawSelection,
  dropCursor,
  rectangularSelection,
  crosshairCursor,
  highlightSpecialChars,
  highlightActiveLineGutter,
} from '@codemirror/view'
import {
  defaultKeymap,
  history,
  historyKeymap,
  indentWithTab,
  redo,
  redoDepth,
  undo,
  undoDepth,
} from '@codemirror/commands'
import {
  HighlightStyle,
  bracketMatching,
  indentOnInput,
  syntaxHighlighting,
} from '@codemirror/language'
import { closeBrackets, closeBracketsKeymap } from '@codemirror/autocomplete'
import { json } from '@codemirror/lang-json'
import { tags } from '@lezer/highlight'
import type { FindTarget } from '@genoffice/ui'
import type { TextMode } from '../../shared/text-mode'
import { buildFindTarget, findHighlight } from './plain-find'

/** Marks the load dispatch so replacing the buffer is not read back as a user edit */
const Loaded = Annotation.define<boolean>()

export interface PlainTextEditorHandle {
  /** replace the buffer without marking the document dirty (the initial file load) */
  setDoc(text: string): void
  focus(): void
  undo(): boolean
  redo(): boolean
  canUndo(): boolean
  canRedo(): boolean
  findTarget(): FindTarget | null
}

interface Props {
  initialText: string
  mode: Extract<TextMode, 'plain' | 'json'>
  onChange: (text: string) => void
  className?: string
  spellcheck: boolean
  /**
   * A right-click over a non-empty selection. The block editor's own menu
   * cannot serve here: it hangs off `editor.view.dom`, and this surface is
   * mounted *instead* of that editor, so the listener would never fire.
   */
  onSelectionRequest?: (text: string) => void
}

/** Colors come from the app's own tokens so the source view follows the light/dark theme */
const highlight = HighlightStyle.define([
  { tag: tags.propertyName, color: 'var(--accent)' },
  { tag: tags.string, color: 'var(--success)' },
  { tag: tags.number, color: 'var(--danger)' },
  { tag: tags.bool, color: 'var(--danger)' },
  { tag: tags.null, color: 'var(--text-muted)' },
  { tag: tags.bracket, color: 'var(--text-muted)' },
  { tag: tags.punctuation, color: 'var(--text-muted)' },
  { tag: tags.invalid, color: 'var(--danger)' },
])

const theme = EditorView.theme({
  '&': { height: '100%', backgroundColor: 'var(--surface)', color: 'var(--text)' },
  '.cm-scroller': {
    fontFamily: "'SF Mono', Menlo, Consolas, monospace",
    fontSize: '13px',
    lineHeight: '1.6',
  },
  '.cm-content': { caretColor: 'var(--text)' },
  '.cm-cursor, .cm-dropCursor': { borderLeftColor: 'var(--text)' },
  '&.cm-focused > .cm-scroller > .cm-selectionLayer .cm-selectionBackground, .cm-selectionBackground':
    { backgroundColor: 'var(--accent-soft)' },
  '.cm-activeLine': { backgroundColor: 'var(--hover)' },
  '.cm-gutters': {
    backgroundColor: 'var(--surface)',
    color: 'var(--text-muted)',
    borderRight: '1px solid var(--border)',
  },
  '.cm-activeLineGutter': { backgroundColor: 'var(--hover)' },
})

function extensionsFor(mode: Props['mode']) {
  return [
    lineNumbers(),
    highlightActiveLineGutter(),
    highlightSpecialChars(),
    history(),
    drawSelection(),
    dropCursor(),
    indentOnInput(),
    syntaxHighlighting(highlight),
    bracketMatching(),
    closeBrackets(),
    rectangularSelection(),
    crosshairCursor(),
    highlightActiveLine(),
    keymap.of([...closeBracketsKeymap, ...defaultKeymap, ...historyKeymap, indentWithTab]),
    // JSON gets its own grammar; plain text gets none, so there is nothing to
    // mistake ordinary prose for code
    ...(mode === 'json' ? [json()] : []),
    findHighlight,
    theme,
    EditorView.lineWrapping,
  ]
}

export const PlainTextEditor = forwardRef<PlainTextEditorHandle, Props>(function PlainTextEditor(
  { initialText, mode, onChange, className, spellcheck, onSelectionRequest },
  ref,
) {
  const hostRef = useRef<HTMLDivElement>(null)
  const viewRef = useRef<EditorView | null>(null)
  const spellRef = useRef<Compartment | null>(null)
  const findTargetRef = useRef<FindTarget | null>(null)
  const onChangeRef = useRef(onChange)
  onChangeRef.current = onChange
  const selectionRequestRef = useRef(onSelectionRequest)
  selectionRequestRef.current = onSelectionRequest

  useEffect(() => {
    const host = hostRef.current
    if (!host) return
    const listeners = new Set<() => void>()
    // spellcheck is a live preference, so it lives in a compartment rather
    // than in the extension set that identifies this editor
    const spellCompartment = new Compartment()
    const view = new EditorView({
      parent: host,
      state: EditorState.create({
        doc: initialText,
        extensions: [
          ...extensionsFor(mode),
          spellCompartment.of(EditorView.contentAttributes.of({ spellcheck: String(spellcheck) })),
          // Right-click over a selection is the same gesture the block
          // editor's menu offers, so it goes to the same dialog. Suppressing
          // the browser menu is the editor's call, not this handler's: a
          // right-click with nothing selected must leave the browser's menu
          // alone, and only the editor knows the selection.
          EditorView.domEventHandlers({
            contextmenu: (event, view) => {
              const onRequest = selectionRequestRef.current
              if (!onRequest) return false
              const { from, to } = view.state.selection.main
              if (from === to) return false
              const selected = view.state.sliceDoc(from, to)
              if (!selected.trim()) return false
              onRequest(selected)
              return true
            },
          }),
          EditorView.updateListener.of((update) => {
            // the load dispatch must not read back as a user edit
            if (update.transactions.some((tr) => tr.annotation(Loaded))) return
            if (!update.docChanged) return
            onChangeRef.current(view.state.doc.toString())
            for (const listener of listeners) listener()
          }),
        ],
      }),
    })
    viewRef.current = view
    spellRef.current = spellCompartment
    findTargetRef.current = buildFindTarget(view, (listener) => {
      listeners.add(listener)
      return () => listeners.delete(listener)
    })
    return () => {
      view.destroy()
      viewRef.current = null
      spellRef.current = null
      findTargetRef.current = null
    }
    // the buffer is seeded once; later external replacements go through setDoc
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [mode])

  useEffect(() => {
    spellRef.current?.reconfigure(
      EditorView.contentAttributes.of({ spellcheck: String(spellcheck) }),
    )
  }, [spellcheck])

  useImperativeHandle(ref, () => ({
    setDoc(text) {
      const view = viewRef.current
      if (!view || view.state.doc.toString() === text) return
      view.dispatch({
        changes: { from: 0, to: view.state.doc.length, insert: text },
        annotations: Loaded.of(true),
      })
    },
    focus: () => viewRef.current?.focus(),
    undo: () => (viewRef.current ? undo(viewRef.current) : false),
    redo: () => (viewRef.current ? redo(viewRef.current) : false),
    canUndo: () => (viewRef.current ? undoDepth(viewRef.current.state) > 0 : false),
    canRedo: () => (viewRef.current ? redoDepth(viewRef.current.state) > 0 : false),
    findTarget: () => findTargetRef.current,
  }))

  return <div ref={hostRef} className={className} />
})
