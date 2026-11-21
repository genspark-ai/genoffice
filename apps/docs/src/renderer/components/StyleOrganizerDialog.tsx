import { useEffect, useState } from 'react'
import {
  listStylesXml,
  readDocxStylesXml,
  styleCopyUpserts,
  upsertStyleXml,
  writeDocxStyles,
  type OrganizerStyle,
  type StyleUpsert,
} from '@genoffice/docx-engine'
import { fetchDocBytes } from '../doc-bytes'
import { useI18n } from '../i18n/locale'
import { useModalKeys } from './modal-keys'

type Side = 'doc' | 'other'

interface OtherFile {
  name: string
  path: string
  bytes: Uint8Array
  /** what Close File writes back */
  edits: StyleUpsert[]
}

const listed = (s: OrganizerStyle): boolean => s.type !== 'numbering'
const sameStyle = (a: OrganizerStyle, b: OrganizerStyle): boolean =>
  a.type === b.type && a.name.toLowerCase() === b.name.toLowerCase()

/**
 * Word's Organizer (Styles tab): the styles of this document beside those of
 * another document or template, copied either way, deleted or renamed. This
 * document's changes join the pending style definitions and are written on
 * save; the other file is written when it is closed here.
 */
export function StyleOrganizerDialog({
  docName,
  docBytes,
  pending,
  onUpsert,
  onDelete,
  onNotice,
  onClose,
}: {
  docName: string
  docBytes: Uint8Array
  pending: Readonly<Record<string, StyleUpsert>>
  onUpsert: (up: StyleUpsert) => string | null
  /** takes the style off the text that uses it and out of the document */
  onDelete: (styleId: string, type: 'paragraph' | 'character') => void
  onNotice: (message: string) => void
  onClose: () => void
}) {
  const { t } = useI18n()
  const [docXml, setDocXml] = useState('')
  const [other, setOther] = useState<OtherFile | null>(null)
  const [otherXml, setOtherXml] = useState('')
  const [side, setSide] = useState<Side>('doc')
  const [picked, setPicked] = useState<string[]>([])
  const [renaming, setRenaming] = useState<string | null>(null)

  useEffect(() => {
    let live = true
    void readDocxStylesXml(docBytes, Object.values(pending)).then((xml) => {
      if (live) setDocXml(xml)
    })
    return () => {
      live = false
    }
  }, [docBytes, pending])
  useEffect(() => {
    let live = true
    if (!other) {
      setOtherXml('')
      return
    }
    void readDocxStylesXml(other.bytes, other.edits).then((xml) => {
      if (live) setOtherXml(xml)
    })
    return () => {
      live = false
    }
  }, [other])

  const docStyles = listStylesXml(docXml).filter(listed)
  const otherStyles = listStylesXml(otherXml).filter(listed)
  const from = side === 'doc' ? docStyles : otherStyles
  const chosen = from.filter((s) => picked.includes(s.styleId))
  const editable = chosen.length > 0 && chosen.every((s) => s.custom && !s.isDefault)
  const deletable = editable && (side === 'other' || chosen.every((s) => s.type !== 'table'))

  const closeOther = async (): Promise<void> => {
    if (!other) return
    if (
      other.edits.length > 0 &&
      window.confirm(t('appOrganizerSaveConfirm', { name: other.name }))
    ) {
      try {
        const bytes = await writeDocxStyles(other.bytes, other.edits)
        const result = await window.desktop.organizerSave(
          other.path,
          bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer,
        )
        onNotice(
          result.ok
            ? t('appOrganizerSaved', { name: other.name })
            : t('appSaveFailed', { error: result.error ?? '' }),
        )
      } catch (err) {
        onNotice(t('appSaveFailed', { error: String(err) }))
      }
    }
    setOther(null)
    if (side === 'other') {
      setSide('doc')
      setPicked([])
    }
  }
  const openOther = async (): Promise<void> => {
    const file = await window.desktop.organizerOpen()
    if (!file) return
    if ('needsPassword' in file) {
      onNotice(t('appDocPwdTitle'))
      return
    }
    setOther({
      name: file.name,
      path: file.path,
      bytes: await fetchDocBytes(file.dataUrl),
      edits: [],
    })
  }
  const close = (): void => {
    void closeOther().then(onClose)
  }
  const keys = useModalKeys(close)

  const editOther = (ups: StyleUpsert[]): void => {
    if (other && ups.length > 0) setOther({ ...other, edits: [...other.edits, ...ups] })
  }
  const copy = (): void => {
    const ups: StyleUpsert[] = []
    let targetXml = side === 'doc' ? otherXml : docXml
    for (const style of chosen) {
      const clash = listStylesXml(targetXml).some((d) => sameStyle(d, style))
      if (clash && !window.confirm(t('appOrganizerOverwrite', { name: style.name }))) continue
      const batch = styleCopyUpserts(side === 'doc' ? docXml : otherXml, style.styleId, targetXml)
      ups.push(...batch)
      // a later pick sees what an earlier one brought along
      for (const up of batch) targetXml = upsertStyleXml(targetXml, up)
    }
    if (side === 'doc') editOther(ups)
    else for (const up of ups) onUpsert(up)
  }
  const remove = (): void => {
    if (!window.confirm(t('appStyleDeleteConfirm', { name: chosen.map((s) => s.name).join(', ') })))
      return
    if (side === 'other') editOther(chosen.map((s) => ({ styleId: s.styleId, remove: true })))
    else
      for (const s of chosen)
        onDelete(s.styleId, s.type === 'character' ? 'character' : 'paragraph')
    setPicked([])
  }
  const rename = (): void => {
    const name = renaming?.trim()
    const style = chosen[0]
    setRenaming(null)
    if (!name || !style || name === style.name) return
    if (
      from.some((s) => s.styleId !== style.styleId && s.name.toLowerCase() === name.toLowerCase())
    ) {
      onNotice(t('appOrganizerNameTaken', { name }))
      return
    }
    const up: StyleUpsert = { styleId: style.styleId, name }
    if (side === 'other') editOther([up])
    else onUpsert(up)
  }

  const list = (which: Side, styles: OrganizerStyle[], label: string) => (
    <label className="organizer-list">
      <span className="organizer-col-title">{label}</span>
      <select
        multiple
        size={12}
        name={which === 'doc' ? 'docStyles' : 'otherStyles'}
        disabled={which === 'other' && !other}
        value={side === which ? picked : []}
        onChange={(e) => {
          setSide(which)
          setRenaming(null)
          setPicked([...e.currentTarget.selectedOptions].map((o) => o.value))
        }}
      >
        {styles.map((s) => (
          <option key={s.styleId} value={s.styleId}>
            {s.name}
          </option>
        ))}
      </select>
    </label>
  )

  return (
    <div
      className="modal-backdrop"
      ref={keys.ref}
      onKeyDown={keys.onKeyDown}
      onMouseDown={(e) => e.target === e.currentTarget && close()}
    >
      <div className="modal tco-dialog organizer-dialog">
        <h2>{t('ribbonStyleOrganizer').replace(/…$/, '')}</h2>
        <div className="organizer-cols">
          <div className="organizer-col">
            {list('doc', docStyles, t('appOrganizerIn', { name: docName }))}
          </div>
          <div className="organizer-buttons">
            <button
              className="btn-ghost"
              name="copy"
              disabled={chosen.length === 0 || !other}
              onClick={copy}
            >
              {side === 'doc' ? `${t('appCopy')} →` : `← ${t('appCopy')}`}
            </button>
            <button className="btn-ghost" name="delete" disabled={!deletable} onClick={remove}>
              {t('ribbonGroupDelete')}
            </button>
            <button
              className="btn-ghost"
              name="rename"
              disabled={!editable || chosen.length !== 1}
              onClick={() => setRenaming(chosen[0]?.name ?? '')}
            >
              {t('ribbonRenameTitle')}…
            </button>
          </div>
          <div className="organizer-col">
            {list(
              'other',
              otherStyles,
              other ? t('appOrganizerIn', { name: other.name }) : t('appOrganizerNoFile'),
            )}
            <button
              className="btn-ghost"
              name="file"
              onClick={() => void (other ? closeOther() : openOther())}
            >
              {other ? t('appOrganizerCloseFile') : t('appOrganizerOpenFile')}
            </button>
          </div>
        </div>
        {renaming !== null && (
          <label className="organizer-rename">
            {t('ribbonRenameTitle')}
            <input
              autoFocus
              name="newName"
              value={renaming}
              onChange={(e) => setRenaming(e.target.value)}
              onKeyDown={(e) => {
                if (e.key === 'Enter') rename()
                else if (e.key === 'Escape') {
                  e.stopPropagation()
                  setRenaming(null)
                }
              }}
            />
            <button className="btn-ghost" name="renameOk" onClick={rename}>
              {t('ribbonOk')}
            </button>
          </label>
        )}
        <div className="modal-actions">
          <span />
          <div className="tco-actions-right">
            <button className="btn-primary" name="close" onClick={close}>
              {t('appClose')}
            </button>
          </div>
        </div>
      </div>
    </div>
  )
}
