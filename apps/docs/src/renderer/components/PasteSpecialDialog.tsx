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
import { Unavailable } from './Unavailable'

type Flavor = 'html' | 'text' | 'picture'

const LABELS: Record<Flavor, StringKey> = {
  html: 'ribbonPasteFormatHtml',
  text: 'ribbonPasteFormatText',
  picture: 'ribbonPasteFormatPicture',
}

const CHOICES: Record<Flavor, PasteChoice> = { html: 'source', text: 'text', picture: 'picture' }

const DESCRIPTIONS: Record<Flavor, StringKey> = {
  html: 'ribbonPasteDescHtml',
  text: 'ribbonPasteDescText',
  picture: 'ribbonPasteDescPicture',
}

/** the application that put the HTML on the clipboard, when it names itself */
export function clipboardSource(html: string): string {
  const meta = /<meta\b[^>]*\bname\s*=\s*["']?generator["']?[^>]*>/i.exec(html)?.[0] ?? ''
  return (
    /\bcontent\s*=\s*(?:"([^"]*)"|'([^']*)'|([^\s>]+))/i
      .exec(meta)
      ?.slice(1)
      .find(Boolean)
      ?.trim() ?? ''
  )
}

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
        <p className="paste-special-source">
          {t('ribbonPasteSpecialSource')}{' '}
          <span>
            {(payload && clipboardSource(payload.html)) || t('ribbonPasteSpecialUnknown')}
          </span>
        </p>
        {/* Paste alone is not a choice, so it comes back together with Paste link */}
        <Unavailable id="paste-special.paste-link">
          <div className="paste-special-modes">
            <label className="protect-check">
              <input type="radio" name="paste-special-mode" checked readOnly />
              <span className="ctl" aria-hidden="true" />
              {t('ribbonPaste')}
            </label>
            <label className="protect-check">
              <input type="radio" name="paste-special-mode" disabled />
              <span className="ctl" aria-hidden="true" />
              {t('ribbonPasteLink')}
            </label>
          </div>
        </Unavailable>
        <label className="paste-special-as">
          {t('ribbonPasteSpecialAs')}
          <select
            size={4}
            aria-label={t('ribbonPasteSpecialAs')}
            value={flavor ?? ''}
            onChange={(e) => setFlavor(e.target.value as Flavor)}
            onDoubleClick={ok}
          >
            {flavors.map((f) => (
              <option key={f} value={f}>
                {t(LABELS[f])}
              </option>
            ))}
          </select>
        </label>
        <Unavailable id="paste-special.display-as-icon">
          <label className="protect-check">
            <input type="checkbox" disabled />
            <span className="ctl" aria-hidden="true" />
            {t('ribbonPasteDisplayAsIcon')}
          </label>
        </Unavailable>
        <p className="fld-hint paste-special-result">
          {payload && flavors.length === 0
            ? t('ribbonPasteSpecialEmpty')
            : flavor
              ? t(DESCRIPTIONS[flavor])
              : ''}
        </p>
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
