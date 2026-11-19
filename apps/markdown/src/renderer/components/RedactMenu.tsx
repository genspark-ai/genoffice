import { useEffect, useState } from 'react'
import type { Editor } from '@tiptap/core'
import { useI18n } from '../i18n/locale'
import { hasRedactionIn } from '../editor/Redaction'
import { RedactDialog } from '../editor/RedactDialog'

interface Props {
  editor: Editor | null
  enabled: boolean
}

/**
 * The editor's own right-click menu, carrying the one command that needs a
 * decision rather than a keystroke: withholding a selection from the model.
 *
 * The browser's menu is suppressed rather than merged, because this edit is
 * not one to make by accident. The dialog repeats the decision and shows the
 * resulting marker before anything is committed.
 */
export function RedactMenu({ editor, enabled }: Props) {
  const { t } = useI18n()
  const [point, setPoint] = useState<{ x: number; y: number } | null>(null)
  const [seed, setSeed] = useState('')
  const [dialogOpen, setDialogOpen] = useState(false)
  /**
   * Whether the selection the menu was opened over already carries the mark.
   *
   * Read once when the menu opens, alongside the seed, because the selection
   * can move while the menu is up and the item has to describe the click the
   * reader made. A bare caret never gets here — the handler returns before this
   * point — so the answer is always about a real selection.
   */
  const [isRedacted, setIsRedacted] = useState(false)

  useEffect(() => {
    if (!editor) return
    const dom = editor.view.dom
    const onContextMenu = (event: MouseEvent) => {
      if (!enabled) return
      // nothing to withhold without a selection, so leave the browser's menu
      if (editor.state.selection.empty) return
      event.preventDefault()
      const { from, to } = editor.state.selection
      setSeed(editor.state.doc.textBetween(from, to, ' ').trim().slice(0, 24))
      setIsRedacted(hasRedactionIn(editor, from, to))
      setPoint({ x: event.clientX, y: event.clientY })
    }
    dom.addEventListener('contextmenu', onContextMenu)
    return () => dom.removeEventListener('contextmenu', onContextMenu)
  }, [editor, enabled])

  // The menu is dismissed the way a context menu anywhere else is: Escape, or a
  // click anywhere outside it. The listeners land on the next tick so the very
  // click that opened the menu does not immediately close it.
  useEffect(() => {
    if (!point) return
    const close = () => setPoint(null)
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') close()
    }
    const timer = window.setTimeout(() => {
      window.addEventListener('mousedown', close)
      window.addEventListener('keydown', onKey)
    }, 0)
    return () => {
      window.clearTimeout(timer)
      window.removeEventListener('mousedown', close)
      window.removeEventListener('keydown', onKey)
    }
  }, [point])

  // The dialog's lifetime must not depend on the menu still being open. The
  // window-level `mousedown` listener above dismisses the menu, and mousedown
  // fires before click — so gating the dialog on `point` unmounted it in the
  // same tick that opened it, and the item did nothing at all.
  return (
    <>
      {point && (
        <div
          className="redact-menu"
          style={{ left: point.x, top: point.y }}
          // stopPropagation as well as preventDefault: this listener is on
          // window in the bubble phase, and preventDefault does not stop the
          // event reaching it
          onMouseDown={(e) => {
            e.preventDefault()
            e.stopPropagation()
          }}
        >
          <button
            type="button"
            className="redact-menu-item"
            onClick={() => {
              setPoint(null)
              setDialogOpen(true)
            }}
          >
            {t('redactMenuLabel')}
          </button>
          {/* The mirror of the item above, and the other half of the pair:
              offered only when the selection really carries the mark. Without
              it the sole way to stop withholding a span was to delete the
              words, which loses the reader's own text to undo a decision they
              made in the editor. No dialog — the words were never replaced, so
              there is no label to ask about and nothing to confirm. */}
          {isRedacted && (
            <button
              type="button"
              className="redact-menu-item"
              onClick={() => {
                setPoint(null)
                editor!.chain().focus().unsetRedaction().run()
              }}
            >
              {t('redactShowLabel')}
            </button>
          )}
        </div>
      )}
      {dialogOpen && editor && (
        <RedactDialog
          seed={seed}
          onCancel={() => setDialogOpen(false)}
          onSubmit={(label) => {
            setDialogOpen(false)
            editor.chain().focus().setRedaction(label).run()
          }}
        />
      )}
    </>
  )
}
