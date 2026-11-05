import type { MergeDataSourceFile } from '../../shared/ipc'
import { useI18n } from '../i18n/locale'
import { useModalKeys } from './modal-keys'

/**
 * The data source of a merge: the file the recipient list comes from, and the
 * sheet of it the records are read from.
 *
 * The file is picked first — the main process owns the file system — and the
 * sheet after it: a workbook holds several and the user names the one the
 * merge reads. A csv has no sheets and is used whole.
 */

export interface MergeDataSourceDialogProps {
  /** the file the user picked, null when none yet */
  file: MergeDataSourceFile | null
  /** the sheet the user confirmed */
  sheet: string
  onSheet: (sheet: string) => void
  /** pick another file */
  onPickFile: () => void
  /** confirm the file and the sheet */
  onConfirm: () => void
  /** why the last read failed, empty when there was no failure */
  error?: string
  onClose: () => void
}

export function MergeDataSourceDialog({
  file,
  sheet,
  onSheet,
  onPickFile,
  onConfirm,
  error,
  onClose,
}: MergeDataSourceDialogProps) {
  const { t } = useI18n()
  const modalKeys = useModalKeys(onClose)
  const sheets = file?.sheets ?? []

  return (
    <div
      className="modal-backdrop"
      ref={modalKeys.ref}
      onKeyDown={modalKeys.onKeyDown}
      onMouseDown={(e) => e.target === e.currentTarget && onClose()}
    >
      <div className="modal modal-mailmerge">
        <h2>{t('ribbonMailingsDataSourceTitle')}</h2>

        <div className="mailmerge-section">
          <div className="mailmerge-label">{t('ribbonMailingsFile')}</div>
          <div className="modal-row">
            <span className="mailmerge-source-name" title={file?.path}>
              {file ? file.path : t('ribbonMailingsNoSource')}
            </span>
            <button onClick={onPickFile}>{t('ribbonMailingsChooseFile')}</button>
          </div>
        </div>

        <div className="mailmerge-section">
          <div className="mailmerge-label">{t('ribbonMailingsSheet')}</div>
          {sheets.length > 0 ? (
            <div className="mailmerge-sheets">
              {sheets.map((name) => (
                <label key={name} className="mailmerge-radio">
                  <input
                    type="radio"
                    name="mailmerge-sheet"
                    checked={name === sheet}
                    onChange={() => onSheet(name)}
                  />
                  {name}
                </label>
              ))}
            </div>
          ) : (
            <div className="mailmerge-empty">
              {file ? t('ribbonMailingsNoSheet') : t('ribbonMailingsNoSource')}
            </div>
          )}
        </div>

        {error && <div className="modal-error">{error}</div>}

        <div className="modal-actions">
          <button className="btn-primary" disabled={!file} onClick={onConfirm}>
            {t('ribbonMailingsConfirmSource')}
          </button>
          <button onClick={onClose}>{t('ribbonMailingsCancel')}</button>
        </div>
      </div>
    </div>
  )
}
