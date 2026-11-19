import { useEffect, useState } from 'react'
import type { ReactNode } from 'react'
import type { Editor } from '@tiptap/core'
import { useEditorState } from '@tiptap/react'
import { useI18n } from '../i18n/locale'
import {
  IconColDelete,
  IconColInsertLeft,
  IconColInsertRight,
  IconHeaderRow,
  IconRowDelete,
  IconRowInsertAbove,
  IconRowInsertBelow,
  IconTableDelete,
} from './icons'
import { uiOp, type TableAction } from '../editor/ops'

interface Props {
  editor: Editor | null
  /** scroll container of the editor canvas, for repositioning on scroll */
  scrollRef: React.RefObject<HTMLElement | null>
  /** Reposition the viewport-anchored menu after document zoom changes. */
  zoom: number
}

function Btn({
  title,
  danger,
  onClick,
  children,
}: {
  title: string
  danger?: boolean
  onClick: () => void
  children: ReactNode
}) {
  // data-tip drives the fast custom tooltip (the native title is too slow to
  // explain icon-only buttons)
  return (
    <button
      type="button"
      className={`tm-btn${danger ? ' danger' : ''}`}
      aria-label={title}
      data-tip={title}
      onMouseDown={(e) => e.preventDefault()}
      onClick={onClick}
    >
      {children}
    </button>
  )
}

/**
 * Floating table toolbar shown while the caret is inside a table, docked to
 * the table's top-right corner. Only markdown-expressible operations: row/
 * column insert & delete, header-row toggle, delete table (no merge — GFM
 * tables cannot serialize spans).
 */
export function TableMenu({ editor, scrollRef, zoom }: Props) {
  const { t } = useI18n()
  const [rect, setRect] = useState<{ top: number; left: number } | null>(null)

  const inTable = useEditorState({
    editor,
    selector: ({ editor: e }) => Boolean(e?.isActive('table')),
  })

  useEffect(() => {
    if (!editor || !inTable) {
      setRect(null)
      return
    }
    const reposition = () => {
      const { from } = editor.state.selection
      const dom = editor.view.domAtPos(from).node
      const el = dom instanceof HTMLElement ? dom : dom.parentElement
      const table = el?.closest('table')
      if (!table) {
        setRect(null)
        return
      }
      const r = table.getBoundingClientRect()
      setRect({ top: r.top - 40, left: Math.max(8, r.right - 8) })
    }
    reposition()
    editor.on('selectionUpdate', reposition)
    editor.on('update', reposition)
    const scroller = scrollRef.current
    scroller?.addEventListener('scroll', reposition, { passive: true })
    window.addEventListener('resize', reposition)
    return () => {
      editor.off('selectionUpdate', reposition)
      editor.off('update', reposition)
      scroller?.removeEventListener('scroll', reposition)
      window.removeEventListener('resize', reposition)
    }
  }, [editor, inTable, scrollRef, zoom])

  if (!editor || !inTable || !rect) return null
  const run = (action: TableAction) =>
    uiOp(editor, { op: 'editTable', target: 'selection', action })
  const ICON = 15

  return (
    <div
      className="table-menu"
      style={{ position: 'fixed', top: rect.top, left: rect.left, transform: 'translateX(-100%)' }}
      onMouseDown={(e) => e.preventDefault()}
    >
      <Btn title={t('tableRowAbove')} onClick={() => run('addRowBefore')}>
        <IconRowInsertAbove size={ICON} />
      </Btn>
      <Btn title={t('tableRowBelow')} onClick={() => run('addRowAfter')}>
        <IconRowInsertBelow size={ICON} />
      </Btn>
      <Btn title={t('tableDeleteRow')} danger onClick={() => run('deleteRow')}>
        <IconRowDelete size={ICON} />
      </Btn>
      <span className="tm-sep" />
      <Btn title={t('tableColLeft')} onClick={() => run('addColumnBefore')}>
        <IconColInsertLeft size={ICON} />
      </Btn>
      <Btn title={t('tableColRight')} onClick={() => run('addColumnAfter')}>
        <IconColInsertRight size={ICON} />
      </Btn>
      <Btn title={t('tableDeleteCol')} danger onClick={() => run('deleteColumn')}>
        <IconColDelete size={ICON} />
      </Btn>
      <span className="tm-sep" />
      <Btn title={t('tableToggleHeaderRow')} onClick={() => run('toggleHeaderRow')}>
        <IconHeaderRow size={ICON} />
      </Btn>
      <Btn title={t('tableDeleteTable')} danger onClick={() => run('deleteTable')}>
        <IconTableDelete size={ICON} />
      </Btn>
    </div>
  )
}
