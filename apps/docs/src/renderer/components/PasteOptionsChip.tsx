/**
 * Post-paste floating "Paste options" chip (Word parity): after pasting
 * foreign HTML (web pages, Word, mail), a small button appears at the end of
 * the pasted content offering the three paste modes — keep source formatting,
 * merge formatting, keep text only. Choosing one undoes the paste, restores
 * the insertion point and re-applies the SAME clipboard payload through the
 * editor's full paste pipeline; "always use this choice" persists the mode
 * as the default for pasting from other programs (editor/paste-options.ts).
 *
 * Mounted at the app root with fixed positioning (coordsAtPos speaks viewport
 * coordinates, the AiAskPopover pattern); repositions while the document
 * scrolls and hides on any other document change.
 */
import React, { useEffect, useRef, useState } from 'react'
import type { Editor } from '@tiptap/core'
import type { EditorState, Transaction } from '@tiptap/pm/state'
import { useI18n, type StringKey } from '../i18n/locale'
import {
  beginForeignPaste,
  capturePasteRestore,
  defaultPasteMode,
  forceNextPasteMode,
  revertPaste,
  setDefaultPasteMode,
  takePastePayload,
  type PasteMode,
  type PastePayload,
  type PasteRestore,
} from '../editor/paste-options'
import { IconPaste } from './icons'

type ChipState = {
  payload: PastePayload
  from: number
  to: number
  mode: PasteMode
  restore: PasteRestore
}

type PastedRange = Pick<ChipState, 'from' | 'to' | 'restore'>

const MODE_KEYS: Array<[PasteMode, StringKey]> = [
  ['source', 'appPasteKeepSource'],
  ['merge', 'appPasteMergeFormat'],
  ['text', 'appPasteTextOnly'],
]

const GAP = 6
const EDGE = 8
const REVERT_META = 'pasteOptionsRevert'

/** The range a transaction's steps touched, in final-doc coordinates. */
function changedRangeOf(tr: Transaction): { from: number; to: number } | null {
  let from = Infinity
  let to = -Infinity
  const maps = tr.mapping.maps
  maps.forEach((stepMap, index) => {
    stepMap.forEach((_oldFrom, _oldTo, newFrom, newTo) => {
      let start = newFrom
      let end = newTo
      for (let later = index + 1; later < maps.length; later++) {
        start = maps[later]!.map(start, 1)
        end = maps[later]!.map(end, -1)
      }
      from = Math.min(from, start)
      to = Math.max(to, end)
    })
  })
  return from <= to ? { from, to } : null
}

export function PasteOptionsChip({ editor }: { editor: Editor }) {
  const { t } = useI18n()
  const [chip, setChip] = useState<ChipState | null>(null)
  const [open, setOpen] = useState(false)
  const [pos, setPos] = useState<{ left: number; top: number } | null>(null)
  // setDefaultPasteMode writes localStorage — bump to re-read it in render
  const [, setPrefTick] = useState(0)
  const reapplying = useRef(false)
  // the state a document-changing transaction started from (editor.state is
  // already the new one when 'transaction' fires) and the paste a re-apply
  // produced (null when it produced none), read back once pasteHTML returns
  const beforeState = useRef<EditorState | null>(null)
  const reapplied = useRef<PastedRange | null>(null)
  const rootRef = useRef<HTMLDivElement | null>(null)

  // show on a paste transaction (payload stashed by the paste entry points),
  // track re-applies, hide on any other document change
  useEffect(() => {
    const onBeforeTransaction = ({ transaction }: { transaction: Transaction }) => {
      if (transaction.docChanged) beforeState.current = editor.state
    }
    const onTransaction = ({
      transaction: tr,
      appendedTransactions,
    }: {
      transaction: Transaction
      appendedTransactions: Transaction[]
    }) => {
      if (!tr.docChanged) return
      if (tr.getMeta(REVERT_META)) {
        reapplied.current = null
        return
      }
      const payload = reapplying.current ? null : takePastePayload()
      if (!payload && !reapplying.current) {
        setChip(null)
        setOpen(false)
        return
      }
      const range = changedRangeOf(tr)
      const restore =
        range && beforeState.current
          ? capturePasteRestore(beforeState.current, [tr, ...appendedTransactions])
          : null
      const pasted = range && restore ? { from: range.from, to: range.to, restore } : null
      if (reapplying.current) {
        reapplied.current = pasted
        return
      }
      setChip(payload && pasted ? { payload, mode: payload.mode, ...pasted } : null)
      setOpen(false)
    }
    editor.on('beforeTransaction', onBeforeTransaction)
    editor.on('transaction', onTransaction)
    return () => {
      editor.off('beforeTransaction', onBeforeTransaction)
      editor.off('transaction', onTransaction)
    }
  }, [editor])

  // anchor at the end of the pasted content; follow scrolling
  useEffect(() => {
    if (!chip) {
      setPos(null)
      return
    }
    const place = () => {
      try {
        const max = editor.state.doc.content.size
        const coords = editor.view.coordsAtPos(Math.min(chip.to, max))
        const left = Math.min(coords.right + GAP, window.innerWidth - EDGE - 32)
        const top = Math.min(coords.bottom + GAP, window.innerHeight - EDGE - 32)
        setPos({ left: Math.max(EDGE, left), top: Math.max(EDGE, top) })
      } catch {
        setChip(null)
      }
    }
    place()
    window.addEventListener('scroll', place, true)
    window.addEventListener('resize', place)
    return () => {
      window.removeEventListener('scroll', place, true)
      window.removeEventListener('resize', place)
    }
  }, [editor, chip])

  // Esc closes the menu, then the chip; outside clicks close the menu
  useEffect(() => {
    if (!chip) return
    const onKey = (event: KeyboardEvent) => {
      if (event.key !== 'Escape') return
      if (open) setOpen(false)
      else {
        setChip(null)
        // focus may sit inside the chip (menu buttons, the checkbox) —
        // dismissing returns it to the document like Word does
        editor.commands.focus()
      }
    }
    const onPointerDown = (event: MouseEvent) => {
      if (rootRef.current && !rootRef.current.contains(event.target as Node)) setOpen(false)
    }
    window.addEventListener('keydown', onKey)
    window.addEventListener('mousedown', onPointerDown)
    return () => {
      window.removeEventListener('keydown', onKey)
      window.removeEventListener('mousedown', onPointerDown)
    }
  }, [editor, chip, open])

  if (!chip || !pos) return null

  const applyMode = (mode: PasteMode) => {
    if (mode === chip.mode) {
      setOpen(false)
      return
    }
    const view = editor.view
    const revert = revertPaste(view.state, chip.restore)
    if (!revert) {
      setChip(null)
      setOpen(false)
      return
    }
    reapplying.current = true
    try {
      view.dispatch(revert.setMeta(REVERT_META, true))
      forceNextPasteMode(mode)
      beginForeignPaste(chip.payload.html)
      const data = new DataTransfer()
      data.setData('text/html', chip.payload.html)
      if (chip.payload.text) data.setData('text/plain', chip.payload.text)
      view.pasteHTML(
        chip.payload.html,
        new ClipboardEvent('paste', { clipboardData: data }) as ClipboardEvent,
      )
    } finally {
      reapplying.current = false
    }
    const pasted = reapplied.current
    setChip((prev) => (prev && pasted ? { ...prev, ...pasted, mode } : null))
    setOpen(false)
    editor.commands.focus()
  }

  const remembered = defaultPasteMode() === chip.mode

  return (
    <div
      ref={rootRef}
      className="paste-chip"
      style={{ left: pos.left, top: pos.top }}
      data-testid="paste-options-chip"
    >
      <button
        className="paste-chip-btn"
        data-tip={t('appPasteOptions')}
        aria-label={t('appPasteOptions')}
        aria-expanded={open}
        onClick={() => setOpen((v) => !v)}
      >
        <IconPaste size={16} />
        <span className="paste-chip-caret">▾</span>
      </button>
      {open && (
        <div className="paste-chip-menu" role="menu">
          {MODE_KEYS.map(([mode, key]) => (
            <button
              key={mode}
              role="menuitemradio"
              aria-checked={chip.mode === mode}
              className={`paste-chip-item ${chip.mode === mode ? 'active' : ''}`}
              onClick={() => applyMode(mode)}
            >
              <span className="paste-chip-check">{chip.mode === mode ? '✓' : ''}</span>
              {t(key)}
            </button>
          ))}
          <div className="paste-chip-sep" />
          <label className="paste-chip-remember">
            <input
              type="checkbox"
              checked={remembered}
              onChange={(event) => {
                setDefaultPasteMode(event.target.checked ? chip.mode : 'source')
                setPrefTick((tick) => tick + 1)
              }}
            />
            {t('appPasteRememberDefault')}
          </label>
        </div>
      )}
    </div>
  )
}
