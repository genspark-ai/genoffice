/** Paste ▸ Paste Special…: the clipboard flavors actually present, pasted as the chosen one. */
import { useEffect, useState } from 'react'
import type { Editor } from '@tiptap/core'
import {
  pastePayload,
  readClipboardPayload,
  type ClipboardPayload,
  type PasteChoice,
} from '../editor/paste-actions'
import { useI18n, type StringKey } from '../i18n/locale'
import { useModalKeys } from './modal-keys'

type Flavor = 'html' | 'text' | 'picture'

const LABELS: Record<Flavor, StringKey> = {
  html: 'ribbonPasteFormatHtml',
  text: 'ribbonPasteFormatText',
  picture: 'ribbonPasteFormatPicture',
}

const CHOICES: Record<Flavor, PasteChoice> = { html: 'source', text: 'text', picture: 'picture' }

export function availableFlavors(payload: ClipboardPayload): Flavor[] {
  const flavors: Flavor[] = []
  if (payload.html) flavors.push('html')
  if (payload.text.trim() || payload.html) flavors.push('text')
  if (payload.image) flavors.push('picture')
  return flavors
}

export function PasteSpecialDialog({ editor, onClose }: { editor: Editor; onClose: () => void }) {
  const { t } = useI18n()
  const keys = useModalKeys(onClose)
  const [payload, setPayload] = useState<ClipboardPayload | null>(null)
  const [flavor, setFlavor] = useState<Flavor | null>(null)

  useEffect(() => {
    let alive = true
    void readClipboardPayload().then((read) => {
      if (!alive) return
      setPayload(read)
      setFlavor(availableFlavors(read)[0] ?? null)
    })
    return () => {
      alive = false
    }
  }, [])

  const flavors = payload ? availableFlavors(payload) : []
  const ok = () => {
    if (payload && flavor) pastePayload(editor, payload, CHOICES[flavor])
    onClose()
  }

  return (
    <div
      className="modal-backdrop"
      ref={keys.ref}
      onKeyDown={keys.onKeyDown}
      onMouseDown={(e) => e.target === e.currentTarget && onClose()}
    >
      <div
        className="modal protect-dialog paste-special-dialog"
        role="dialog"
        aria-label={t('ribbonPasteSpecialTitle')}
      >
        <h2>{t('ribbonPasteSpecialTitle')}</h2>
        <p className="modal-desc">{t('ribbonPasteSpecialAs')}</p>
        {payload && flavors.length === 0 && (
          <p className="fld-hint">{t('ribbonPasteSpecialEmpty')}</p>
        )}
        {flavors.map((f) => (
          <label key={f} className="protect-check">
            <input
              type="radio"
              name="paste-special-flavor"
              checked={flavor === f}
              onChange={() => setFlavor(f)}
              onDoubleClick={ok}
            />
            <span className="ctl" aria-hidden="true" />
            {t(LABELS[f])}
          </label>
        ))}
        <div className="modal-actions">
          <button type="button" onClick={onClose}>
            {t('ribbonCancel')}
          </button>
          <button type="button" className="primary" onClick={ok} disabled={!flavor}>
            {t('ribbonOk')}
          </button>
        </div>
      </div>
    </div>
  )
}
