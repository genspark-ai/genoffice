import {
  forwardRef,
  useEffect,
  useImperativeHandle,
  useLayoutEffect,
  useRef,
  useState,
} from 'react'
import { EditorState } from '@codemirror/state'
import { EditorView } from '@codemirror/view'
import { redo, redoDepth, undo, undoDepth } from '@codemirror/commands'
import { External, buildExtensions } from './cm-setup'
import { addAiRanges, clearAiRanges } from './cm-highlight'
import { cmFindTarget } from './find-target'
import { useI18n } from '../i18n/locale'
import type { FindTarget } from '@genoffice/ui'
import type { Patch } from '../document/patch'

/** a source range and the text it covers, as the reader sees it */
export interface SourceRange {
  from: number
  to: number
  text: string
}

export interface SourceEditorHandle {
  /** replace the whole document without touching the undo history (file load) */
  setDoc(text: string): void
  /** replace the whole document as one undoable step (AI rewrite, rollback) */
  replaceDoc(text: string, highlight: boolean): void
  /** apply validated patches as one undoable step; returns the post-edit ranges */
  applyPatches(patches: readonly Patch[], highlight: boolean): Array<[number, number]>
  clearHighlights(): void
  /** select and scroll a source range into view; focus only when the user asked for the editor */
  revealRange(from: number, to: number, focus?: boolean): void
  /** the main selection in source offsets; null until the editor is mounted */
  getSelection(): (SourceRange & { empty: boolean }) | null
  undo(): boolean
  redo(): boolean
  canUndo(): boolean
  canRedo(): boolean
  focus(): void
  /** adapter for the find/replace panel; null until the editor is mounted */
  findTarget(): FindTarget | null
}

export interface CursorInfo {
  line: number
  col: number
  /** document offset of the cursor head */
  pos: number
}

interface Props {
  initialText: string
  onChange: (text: string) => void
  onCursor: (cursor: CursorInfo) => void
  onBeforeReplace?: () => void
  /** only the editor's own transactions report changes; External-annotated ones are already known to the caller */
  className?: string
  /**
   * Right-click a selection to withhold it from the model.
   *
   * The menu is here, in the source pane, and not in the preview, on purpose: a
   * page's secrets are usually not text the reader can select. They sit in
   * `<script>` strings, in `content=` / `src=` / `onclick=` values and in
   * comments, and a rendered page puts no caret in any of them — the source can
   * select all of them, so this is the surface that reaches the real cases.
   */
  onRedact?: (range: SourceRange) => void
}

interface MenuState extends SourceRange {
  x: number
  y: number
}

export const SourceEditor = forwardRef<SourceEditorHandle, Props>(function SourceEditor(
  { initialText, onChange, onCursor, onBeforeReplace, className, onRedact },
  ref,
) {
  const { t } = useI18n()
  const hostRef = useRef<HTMLDivElement>(null)
  const viewRef = useRef<EditorView | null>(null)
  const findTargetRef = useRef<FindTarget | null>(null)
  const docListeners = useRef(new Set<() => void>())
  const onChangeRef = useRef(onChange)
  const onCursorRef = useRef(onCursor)
  onChangeRef.current = onChange
  onCursorRef.current = onCursor
  const [menu, setMenu] = useState<MenuState | null>(null)
  const [menuPos, setMenuPos] = useState({ left: 0, top: 0 })
  const menuRef = useRef<HTMLDivElement>(null)
  const onRedactRef = useRef(onRedact)
  onRedactRef.current = onRedact

  useEffect(() => {
    const host = hostRef.current
    if (!host) return
    const view = new EditorView({
      parent: host,
      state: EditorState.create({
        doc: initialText,
        extensions: buildExtensions((v) => {
          const head = v.state.selection.main.head
          const line = v.state.doc.lineAt(head)
          onCursorRef.current({ line: line.number, col: head - line.from + 1, pos: head })
        }),
      }),
      dispatchTransactions: (trs, v) => {
        v.update(trs)
        if (trs.some((tr) => tr.docChanged && !tr.annotation(External))) {
          onChangeRef.current(v.state.doc.toString())
        }
        if (trs.some((tr) => tr.docChanged)) for (const l of docListeners.current) l()
      },
    })
    viewRef.current = view
    findTargetRef.current = cmFindTarget(
      view,
      (listener) => {
        docListeners.current.add(listener)
        return () => docListeners.current.delete(listener)
      },
      onBeforeReplace,
    )
    return () => {
      view.destroy()
      viewRef.current = null
      findTargetRef.current = null
    }
    // the document is seeded once; later external replacements go through setDoc
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [])

  /**
   * A bare caret has nothing to withhold, and the source pane keeps Chromium's own
   * menu in that case — copy, paste and the spellchecker are still worth having on
   * a selection we are not claiming.
   */
  const onContextMenu = (event: React.MouseEvent<HTMLDivElement>) => {
    const view = viewRef.current
    if (!view || !onRedactRef.current) return
    const sel = view.state.selection.main
    if (sel.empty) return
    event.preventDefault()
    setMenu({
      x: event.clientX,
      y: event.clientY,
      from: sel.from,
      to: sel.to,
      text: view.state.sliceDoc(sel.from, sel.to),
    })
  }

  // keep the menu on screen: flip it up rather than open it off the bottom edge
  useLayoutEffect(() => {
    if (!menu) return
    const rect = menuRef.current?.getBoundingClientRect()
    if (!rect) return
    setMenuPos({
      left: Math.max(8, Math.min(menu.x, window.innerWidth - rect.width - 8)),
      top:
        menu.y + rect.height > window.innerHeight - 8 ? Math.max(8, menu.y - rect.height) : menu.y,
    })
  }, [menu])

  useEffect(() => {
    if (!menu) return
    const onMouseDown = (e: MouseEvent) => {
      if (!menuRef.current?.contains(e.target as Node)) setMenu(null)
    }
    const onKeyDown = (e: KeyboardEvent) => {
      if (e.key === 'Escape') setMenu(null)
    }
    window.addEventListener('mousedown', onMouseDown)
    window.addEventListener('keydown', onKeyDown, true)
    return () => {
      window.removeEventListener('mousedown', onMouseDown)
      window.removeEventListener('keydown', onKeyDown, true)
    }
  }, [menu])

  useImperativeHandle(ref, () => ({
    setDoc(text) {
      const view = viewRef.current
      if (!view || view.state.doc.toString() === text) return
      view.dispatch({
        changes: { from: 0, to: view.state.doc.length, insert: text },
        annotations: [External.of(true)],
      })
    },
    replaceDoc(text, highlight) {
      const view = viewRef.current
      if (!view) return
      const len = view.state.doc.length
      view.dispatch({
        changes: { from: 0, to: len, insert: text },
        annotations: [External.of(true)],
        effects: highlight
          ? [clearAiRanges.of(null), addAiRanges.of([[0, text.length]])]
          : [clearAiRanges.of(null)],
      })
    },
    applyPatches(patches, highlight) {
      const view = viewRef.current
      if (!view) return []
      const sorted = [...patches].sort((a, b) => a.from - b.from || a.to - b.to)
      const ranges: Array<[number, number]> = []
      let delta = 0
      for (const p of sorted) {
        ranges.push([p.from + delta, p.from + delta + p.text.length])
        delta += p.text.length - (p.to - p.from)
      }
      view.dispatch({
        changes: sorted.map((p) => ({ from: p.from, to: p.to, insert: p.text })),
        annotations: [External.of(true)],
        effects: highlight ? [addAiRanges.of(ranges)] : [],
      })
      return ranges
    },
    clearHighlights() {
      viewRef.current?.dispatch({ effects: [clearAiRanges.of(null)] })
    },
    revealRange(from, to, focus = true) {
      const view = viewRef.current
      if (!view) return
      const len = view.state.doc.length
      const a = Math.max(0, Math.min(from, len))
      const b = Math.max(a, Math.min(to, len))
      view.dispatch({
        selection: { anchor: a, head: b },
        effects: EditorView.scrollIntoView(a, { y: 'center' }),
        annotations: [External.of(true)],
      })
      if (focus) view.focus()
    },
    getSelection() {
      const view = viewRef.current
      if (!view) return null
      const sel = view.state.selection.main
      return {
        from: sel.from,
        to: sel.to,
        empty: sel.empty,
        text: view.state.sliceDoc(sel.from, sel.to),
      }
    },
    undo: () => (viewRef.current ? undo(viewRef.current) : false),
    redo: () => (viewRef.current ? redo(viewRef.current) : false),
    canUndo: () => (viewRef.current ? undoDepth(viewRef.current.state) > 0 : false),
    canRedo: () => (viewRef.current ? redoDepth(viewRef.current.state) > 0 : false),
    focus: () => viewRef.current?.focus(),
    findTarget: () => findTargetRef.current,
  }))

  return (
    <>
      <div ref={hostRef} className={className} onContextMenu={onContextMenu} />
      {menu && (
        <div
          ref={menuRef}
          className="source-ctx"
          role="menu"
          style={{ left: menuPos.left, top: menuPos.top }}
          onContextMenu={(e) => e.preventDefault()}
        >
          <button
            className="source-ctx-item"
            role="menuitem"
            onClick={() => {
              const range: SourceRange = { from: menu.from, to: menu.to, text: menu.text }
              setMenu(null)
              onRedactRef.current?.(range)
            }}
          >
            {t('redactMenuLabel')}
          </button>
        </div>
      )}
    </>
  )
})
