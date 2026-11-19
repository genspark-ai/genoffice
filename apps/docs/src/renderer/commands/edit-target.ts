import { useEffect, useState } from 'react'
import type { Editor } from '@tiptap/core'
import { getActiveSubEditor } from '../editor/active-editor'
import { onEditableChange } from '../editor/editable'

/** what Cut, Copy and Paste would act on right now */
export interface EditTarget {
  /** something is selected to cut or copy */
  hasSelection: boolean
  /** the target takes edits */
  editable: boolean
  /** the focus is in a text field rather than the document */
  field: boolean
}

export function readEditTarget(editor: Editor | null): EditTarget {
  // another view (the AI panel, a dialog window) owns the focus: its state is not ours to judge
  if (!document.hasFocus()) return { hasSelection: true, editable: true, field: true }
  const active = document.activeElement
  if (active instanceof HTMLInputElement || active instanceof HTMLTextAreaElement) {
    return {
      hasSelection: active.selectionStart !== active.selectionEnd,
      editable: !active.readOnly && !active.disabled,
      field: true,
    }
  }
  const dom = window.getSelection()
  const domSelection = Boolean(dom && !dom.isCollapsed)
  if (
    active instanceof HTMLElement &&
    active.isContentEditable &&
    !active.closest('.ProseMirror')
  ) {
    return { hasSelection: domSelection, editable: true, field: true }
  }
  // a picture or a block of cells is a selection the DOM does not always show as a range
  const target = getActiveSubEditor() ?? editor
  return {
    hasSelection: target ? !target.state.selection.empty || domSelection : domSelection,
    editable: Boolean(target?.isEditable),
    field: false,
  }
}

const same = (a: EditTarget, b: EditTarget) =>
  a.hasSelection === b.hasSelection && a.editable === b.editable && a.field === b.field

/** follows the selection and the focus, wherever they are in this window */
export function useEditTarget(editor: Editor | null): EditTarget {
  const [target, setTarget] = useState(() => readEditTarget(editor))
  useEffect(() => {
    let frame = 0
    const refresh = () => {
      cancelAnimationFrame(frame)
      frame = requestAnimationFrame(() => {
        const next = readEditTarget(editor)
        setTarget((prev) => (same(prev, next) ? prev : next))
      })
    }
    refresh()
    const docEvents = ['selectionchange', 'focusin', 'focusout', 'input'] as const
    for (const name of docEvents) document.addEventListener(name, refresh)
    window.addEventListener('focus', refresh)
    window.addEventListener('blur', refresh)
    editor?.on('transaction', refresh)
    const offEditable = editor ? onEditableChange(editor, refresh) : undefined
    return () => {
      cancelAnimationFrame(frame)
      for (const name of docEvents) document.removeEventListener(name, refresh)
      window.removeEventListener('focus', refresh)
      window.removeEventListener('blur', refresh)
      editor?.off('transaction', refresh)
      offEditable?.()
    }
  }, [editor])
  return target
}
