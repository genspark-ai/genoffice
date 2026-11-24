/** Paste ▸ Paste Special…: the clipboard flavors actually present, pasted as the chosen one. */
import { useEffect, useState } from 'react'
import type { Editor } from '@tiptap/core'
import {
  pastePayload,
  readClipboardPayload,
  type ClipboardPayload,
  type PasteChoice,
} from '../editor/paste-actions'
import { canPasteLink, pasteLink } from '../editor/paste-link'
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
  // a bitmap pastes as itself; copied text or a table is drawn as one (Word's Paste as Picture)
  if (payload.image || payload.html || payload.text.trim()) flavors.push('picture')
  return flavors
}

export function PasteSpecialDialog({ editor, onClose }: { editor: Editor; onClose: () => void }) {
  const { t } = useI18n()
  const keys = useModalKeys(onClose)
  const [payload, setPayload] = useState<ClipboardPayload | null>(null)
  const [flavor, setFlavor] = useState<Flavor | null>(null)
  /** the clipboard was copied from a saved document that is still open: it can be linked to */
  const [linkable, setLinkable] = useState(false)
  const [asLink, setAsLink] = useState(false)
  const [gone, setGone] = useState(false)

  useEffect(() => {
    let alive = true
    void readClipboardPayload().then(async (read) => {
      const link = await canPasteLink(read)
      if (!alive) return
      setPayload(read)
      setFlavor(availableFlavors(read)[0] ?? null)
      setLinkable(link)
    })
    return () => {
      alive = false
    }
  }, [])

  // a link is to text: the picture flavor is not one of its forms
  const flavors = (payload ? availableFlavors(payload) : []).filter(
    (f) => !asLink || f !== 'picture',
  )
  const chosen = flavor && flavors.includes(flavor) ? flavor : (flavors[0] ?? null)
  const ok = async () => {
    if (payload && chosen && asLink) {
      if (!(await pasteLink(editor, payload, chosen === 'text'))) {
        setGone(true)
        return
      }
    } else if (payload && chosen) pastePayload(editor, payload, CHOICES[chosen])
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
        <div className="paste-special-modes">
          <label className="protect-check">
            <input
              type="radio"
              name="paste-special-mode"
              checked={!asLink}
              onChange={() => setAsLink(false)}
            />
            <span className="ctl" aria-hidden="true" />
            {t('ribbonPaste')}
          </label>
          <label className="protect-check">
            <input
              type="radio"
              name="paste-special-mode"
              checked={asLink}
              disabled={!linkable}
              onChange={() => {
                setAsLink(true)
                setGone(false)
              }}
            />
            <span className="ctl" aria-hidden="true" />
            {t('ribbonPasteLink')}
          </label>
        </div>
        <label className="paste-special-as">
          {t('ribbonPasteSpecialAs')}
          <select
            size={4}
            aria-label={t('ribbonPasteSpecialAs')}
            value={chosen ?? ''}
            onChange={(e) => setFlavor(e.target.value as Flavor)}
            onDoubleClick={() => void ok()}
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
            : chosen
              ? `${t(DESCRIPTIONS[chosen])}${asLink ? ` ${t('ribbonPasteLinkResult')}` : ''}`
              : ''}
        </p>
        {gone && (
          <p className="fld-hint paste-special-gone" role="alert">
            {t('ribbonPasteLinkGone')}
          </p>
        )}
        <div className="modal-actions">
          <button type="button" onClick={onClose}>
            {t('ribbonCancel')}
          </button>
          <button type="button" className="primary" onClick={() => void ok()} disabled={!chosen}>
            {t('ribbonOk')}
          </button>
        </div>
      </div>
    </div>
  )
}
