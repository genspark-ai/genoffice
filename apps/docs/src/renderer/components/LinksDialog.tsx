import { useEffect, useMemo, useState } from 'react'
import type { Editor } from '@tiptap/core'
import type { Block, HfPictureLink } from '@genoffice/docx-engine'
import { HF_COMMIT_EVENT } from '../editor/hf-editor'
import type { InsertFileNotes } from '../editor/insert-file'
import {
  breakLinks,
  collectLinks,
  hfPictureLink,
  linkSourceOf,
  linkStatusKey as statusKey,
  linkTypeOf,
  readLinkedPicture,
  readsDocument,
  linkUpdatable,
  setLinksAutomatic,
  setLinksLocked,
  setLinkSource,
  updateLink,
  type DocLink,
  type LinkUpdateResult,
} from '../editor/link-fields'
import type { HfPictureLinkRow } from '../hf-sections'
import { useI18n, type StringKey } from '../i18n/locale'
import { useModalKeys } from './modal-keys'

const TYPE_KEYS = {
  document: 'ribbonDocTitle',
  picture: 'ribbonPicture',
  object: 'ribbonLinksTypeObject',
} as const satisfies Record<string, StringKey>

const fileName = (source: string): string => source.split(/[\\/]/).pop() || source

/** picks the file a link is to name from now on */
async function pickSource(link: DocLink): Promise<{ path: string; url: string } | null> {
  if (link.type === 'picture') {
    const picked = await window.desktop.pickImage()
    return picked ? { path: picked.path, url: picked.url } : null
  }
  if (readsDocument(link.info)) {
    const picked = await window.desktop.pickDocx()
    return picked && 'path' in picked ? { path: picked.path, url: '' } : null
  }
  const picked = await window.desktop.pickLinkFile()
  return picked ? { path: picked.path, url: picked.url } : null
}

/** the linked pictures of the headers and footers, and the way a change to one reaches its strip */
export interface HfLinkHost {
  rows: HfPictureLinkRow[]
  set: (row: HfPictureLinkRow, change: Pick<HfPictureLink, 'target' | 'bytes'>) => boolean
}

/**
 * Word's Edit > Links: the content of the document that comes from other
 * files, each with its source, and what can be done about it. Every action
 * is an edit of the document, undone like any other; one on a picture of a
 * header or footer is kept with its strip, as a move of that picture is.
 */
export function LinksDialog({
  editor,
  blocks,
  notes,
  hf,
  onClose,
}: {
  editor: Editor
  blocks: Block[]
  notes: InsertFileNotes
  hf?: HfLinkHost
  onClose: () => void
}) {
  const { t } = useI18n()
  const keys = useModalKeys(onClose)
  const [version, setVersion] = useState(0)
  const [picked, setPicked] = useState<number[]>([0])
  const [anchor, setAnchor] = useState(0)
  const [busy, setBusy] = useState(false)
  // by row; cleared when the rows themselves change
  const [status, setStatus] = useState<Record<number, StringKey>>({})
  // a strip being typed in holds its pictures as they were when it was opened: it is committed first
  useEffect(() => {
    window.dispatchEvent(new Event(HF_COMMIT_EVENT))
  }, [])
  const hfRows = hf?.rows
  // the document's links, then those of its headers and footers
  const links = useMemo(
    () => [...collectLinks(editor.state.doc, blocks), ...(hfRows ?? []).map(hfPictureLink)],
    // eslint-disable-next-line react-hooks/exhaustive-deps -- version stands for the document's state
    [editor, blocks, version, hfRows],
  )
  /** the link of a row as it stands now; a header's keeps its row while the document's move */
  const linkAt = (at: number): DocLink | undefined => {
    const inDoc = collectLinks(editor.state.doc, blocks)
    return at < inDoc.length ? inDoc[at] : links[at]?.type === 'picture' ? links[at] : undefined
  }
  const inList = picked.filter((i) => i < links.length)
  const rows = inList.length > 0 || links.length === 0 ? inList : [links.length - 1]
  const chosen = rows.map((i) => links[i])
  // Open Source and Change Source take one link
  const link = chosen.length === 1 ? chosen[0] : undefined
  const row = link ? rows[0] : -1
  const refresh = () => setVersion((v) => v + 1)
  const editable = editor.isEditable && !busy
  const fields = chosen.filter((l) => l.type !== 'picture')
  const oles = fields.filter((l) => l.info.kind === 'LINK')
  const automatic = oles.length > 0 && oles.every((l) => l.info.automatic)
  const locked = fields.length > 0 && fields.every((l) => l.locked)

  const pick = (i: number, e: { shiftKey: boolean; metaKey: boolean; ctrlKey: boolean }) => {
    if (e.shiftKey) {
      const [from, to] = anchor < i ? [anchor, i] : [i, anchor]
      setPicked(Array.from({ length: to - from + 1 }, (_, n) => from + n))
      return
    }
    setAnchor(i)
    if (!(e.metaKey || e.ctrlKey)) setPicked([i])
    else if (!rows.includes(i)) setPicked([...rows, i].sort((a, b) => a - b))
    else if (rows.length > 1) setPicked(rows.filter((r) => r !== i))
  }

  /** a header's picture names `target` and takes the bytes read from it; unread, it only names it */
  const updateInHf = async (row: HfPictureLinkRow, target: string): Promise<LinkUpdateResult> => {
    const read = await readLinkedPicture(target)
    const bytes = 'status' in read ? undefined : { base64: read.base64, mime: read.mime }
    const taken = hf?.set(row, { target, ...(bytes ? { bytes } : {}) })
    if ('status' in read) return read
    return taken ? { status: 'updated', leftOut: 0 } : { status: 'failed', reason: 'failed' }
  }
  const update = async (target: DocLink, at: number, source?: string) => {
    setBusy(true)
    try {
      const result =
        target.type === 'picture' && target.hf
          ? await updateInHf(target.hf, source ?? target.target)
          : await updateLink(editor, target, notes)
      setStatus((s) => ({ ...s, [at]: statusKey(result) }))
    } finally {
      setBusy(false)
      refresh()
    }
  }
  const updateChosen = async () => {
    // from the last row: what an update brings in moves only the rows after it
    for (const at of [...rows].reverse()) {
      const target = linkAt(at)
      if (target && linkUpdatable(target)) await update(target, at)
    }
  }
  const changeSource = async () => {
    if (!link) return
    const source = await pickSource(link)
    if (!source) return
    if (link.type === 'picture' && link.hf) return update(link, row, source.url)
    if (!setLinkSource(editor, link, blocks, source)) return
    const next = collectLinks(editor.state.doc, blocks)[row]
    if (next && linkUpdatable(next)) await update(next, row)
    else refresh()
  }
  const openSource = async () => {
    if (!link) return
    const opened = await window.desktop.openLinkSource(linkSourceOf(link))
    if (!opened) setStatus((s) => ({ ...s, [row]: 'ribbonLinksMissing' }))
  }
  const setAutomatic = (on: boolean) => {
    if (setLinksAutomatic(editor, oles, blocks, on) > 0) refresh()
  }
  const breakLink = () => {
    if (chosen.length === 0 || !window.confirm(t('ribbonLinksBreakConfirm'))) return
    breakLinks(editor, chosen, blocks)
    for (const entry of chosen)
      if (entry.type === 'picture' && entry.hf) hf?.set(entry.hf, { target: null })
    setStatus({})
    const next = Math.max(0, rows[0] - 1)
    setPicked([next])
    setAnchor(next)
    refresh()
  }

  return (
    <div
      className="modal-backdrop"
      ref={keys.ref}
      onKeyDown={keys.onKeyDown}
      onMouseDown={(e) => e.target === e.currentTarget && onClose()}
    >
      <div className="modal gs-form tco-dialog links-dialog" role="dialog" aria-modal="true">
        <h2>{t('ribbonLinksTitle')}</h2>
        <div className="links-body">
          <div
            className="links-list"
            role="listbox"
            aria-multiselectable="true"
            aria-label={t('ribbonLinksTitle')}
          >
            <div className="links-row links-head" aria-hidden="true">
              <span>{t('ribbonLinksSource')}</span>
              <span>{t('ribbonLinksItem')}</span>
              <span>{t('ribbonLinksType')}</span>
              <span>{t('ribbonLinksUpdate')}</span>
            </div>
            {links.map((entry, i) => (
              <button
                type="button"
                key={i}
                role="option"
                aria-selected={rows.includes(i)}
                className={`links-row${rows.includes(i) ? ' selected' : ''}`}
                onClick={(e) => pick(i, e)}
              >
                <span title={linkSourceOf(entry)}>{fileName(linkSourceOf(entry))}</span>
                <span>{entry.type === 'picture' ? '' : entry.info.item}</span>
                <span>{t(TYPE_KEYS[linkTypeOf(entry)])}</span>
                <span className="links-update">
                  {status[i]
                    ? t(status[i])
                    : entry.type !== 'picture' && entry.locked
                      ? t('ribbonLinksLocked')
                      : entry.type !== 'picture' && entry.info.automatic
                        ? t('ribbonLinksAutomatic')
                        : t('ribbonLinksManual')}
                </span>
              </button>
            ))}
          </div>
          <div className="links-actions">
            <button
              type="button"
              name="update"
              disabled={!editable || !chosen.some(linkUpdatable)}
              onClick={() => void updateChosen()}
            >
              {t('ribbonLinksUpdateNow')}
            </button>
            <button
              type="button"
              name="open"
              disabled={!link || busy}
              onClick={() => void openSource()}
            >
              {t('ribbonLinksOpenSource')}
            </button>
            <button
              type="button"
              name="change"
              disabled={!editable || !link}
              onClick={() => void changeSource()}
            >
              {t('ribbonLinksChangeSource')}
            </button>
            <button
              type="button"
              name="break"
              disabled={!editable || chosen.length === 0}
              onClick={breakLink}
            >
              {t('ribbonLinksBreak')}
            </button>
          </div>
        </div>
        <p className="links-source">{link ? linkSourceOf(link) : ''}</p>
        <div className="links-options">
          <span>{t('ribbonLinksUpdate')}:</span>
          <label className="tco-check">
            <input
              type="radio"
              name="links-update"
              checked={automatic}
              disabled={!editable || oles.length === 0}
              onChange={() => setAutomatic(true)}
            />
            {t('ribbonLinksAutomatic')}
          </label>
          <label className="tco-check">
            <input
              type="radio"
              name="links-update"
              checked={chosen.length > 0 && !automatic}
              disabled={!editable || oles.length === 0}
              onChange={() => setAutomatic(false)}
            />
            {t('ribbonLinksManual')}
          </label>
          <label className="tco-check">
            <input
              type="checkbox"
              name="links-locked"
              checked={locked}
              disabled={!editable || fields.length === 0}
              onChange={(e) => {
                if (setLinksLocked(editor, fields, blocks, e.target.checked) > 0) refresh()
              }}
            />
            {t('ribbonLinksLocked')}
          </label>
        </div>
        <div className="modal-actions">
          <span />
          <div className="tco-actions-right">
            <button type="button" className="primary" name="close" onClick={onClose}>
              {t('ribbonClose')}
            </button>
          </div>
        </div>
      </div>
    </div>
  )
}
