/**
 * The pane listing notes that carry a picture, with the picture itself. The
 * image is fetched by id through readWorkbookMedia, the same route an in-cell
 * picture takes, so the bytes never travel with the workbook metadata — unless
 * the picture was put into the note this session and is not written to the
 * file yet, which is the one case where the bytes are already at hand.
 */
import { useSyncExternalStore } from 'react'

import { formatAddress } from '@genoffice/xlsx-gateway/domain/cell-address'

import type { WorkbookFile } from '../shared/desktop-api'
import { useI18n } from './i18n/locale'
import {
  collectNotePictures,
  notePicturesPane,
  type NotePicture,
  type PendingNotePicture,
} from './note-pictures'
import { useWorkbookMediaUrl } from './workbook-media'
import { revealCellBelowFreeze } from './univer-sync'
import type { UniverRuntime } from './univer-state'

/** A picture that is in the file, fetched by id. A failure leaves the text. */
function FilePicture({
  sessionId,
  imageId,
}: {
  readonly sessionId: string
  readonly imageId: string
}) {
  const { t } = useI18n()
  const { url, unavailable } = useWorkbookMediaUrl(sessionId, imageId)
  if (url) return <img className="np-image" src={url} alt="" />
  if (unavailable) return <p className="np-image-missing">{t('appNotePictureUnavailable')}</p>
  return null
}

function Picture({
  sessionId,
  picture,
}: {
  readonly sessionId: string
  readonly picture: NotePicture
}) {
  // a picture put into the note this session is not in the file yet, so its
  // bytes are already at hand
  if (picture.dataUrl) return <img className="np-image" src={picture.dataUrl} alt="" />
  if (!picture.imageId) return null
  return <FilePicture sessionId={sessionId} imageId={picture.imageId} />
}

function Card({
  picture,
  sessionId,
  onJump,
  onRemove,
}: {
  readonly picture: NotePicture
  readonly sessionId: string
  readonly onJump: (sheetId: string, row: number, column: number) => void
  readonly onRemove: (sheetId: string, row: number, column: number) => void
}) {
  const { t } = useI18n()
  return (
    <article className="np-card">
      <header>
        <button
          type="button"
          className="np-cell"
          onClick={() => onJump(picture.sheetId, picture.row, picture.column)}
          title={picture.sheetName}
        >
          {picture.sheetName}!{formatAddress(picture.row, picture.column)}
        </button>
        {picture.author && <span className="np-author">{picture.author}</span>}
        <button
          type="button"
          className="np-remove"
          aria-label={t('appNotePictureRemove')}
          title={t('appNotePictureRemove')}
          onClick={() => onRemove(picture.sheetId, picture.row, picture.column)}
        >
          ×
        </button>
      </header>
      <Picture sessionId={sessionId} picture={picture} />
      {picture.text && <p className="np-text">{picture.text}</p>}
    </article>
  )
}

export function NotePicturesPane({
  file,
  pending,
  onRemove,
  getRuntime,
}: {
  readonly file: WorkbookFile | null
  readonly pending: readonly PendingNotePicture[]
  readonly onRemove: (sheetId: string, row: number, column: number) => void
  readonly getRuntime: () => UniverRuntime | null
}): React.JSX.Element | null {
  const { t } = useI18n()
  const pane = useSyncExternalStore(notePicturesPane.subscribe, notePicturesPane.getState)
  const pictures = collectNotePictures(file, pending)
  if (!pane.open) return null

  const sessionId = file?.sessionId
  const jump = (sheetId: string, row: number, column: number): void => {
    const workbook = getRuntime()?.univerAPI.getActiveWorkbook()
    const sheet = workbook?.getSheetBySheetId(sheetId)
    if (!workbook || !sheet) return
    if (sheet.getSheetId() !== workbook.getActiveSheet()?.getSheetId())
      workbook.setActiveSheet(sheet)
    sheet.getRange(row, column, 1, 1).activate()
    void revealCellBelowFreeze(sheet, row, column)
  }

  return (
    <aside className="note-pictures-pane" aria-label={t('appNotePicturesPaneTitle')}>
      <header>
        <span>{t('appNotePicturesPaneTitle')}</span>
        <button
          type="button"
          className="np-close"
          aria-label={t('appNotePicturesClose')}
          onClick={() => notePicturesPane.close()}
        >
          ×
        </button>
      </header>
      <div className="np-list">
        {sessionId === undefined || pictures.length === 0 ? (
          <p className="np-empty">{t('appNotePicturesNone')}</p>
        ) : (
          pictures.map((picture) => (
            <Card
              key={`${picture.sheetId}:${picture.row}:${picture.column}`}
              picture={picture}
              sessionId={sessionId}
              onJump={jump}
              onRemove={onRemove}
            />
          ))
        )}
      </div>
    </aside>
  )
}
