/**
 * The pane listing notes that carry a picture, with the picture itself. The
 * image is fetched by id through readWorkbookMedia, the same route an in-cell
 * picture takes, so the bytes never travel with the workbook metadata.
 */
import { useSyncExternalStore } from 'react'

import { formatAddress } from '@genoffice/xlsx-gateway/domain/cell-address'

import type { WorkbookFile } from '../shared/desktop-api'
import { useI18n } from './i18n/locale'
import { collectNotePictures, notePicturesPane, type NotePicture } from './note-pictures'
import { useWorkbookMediaUrl } from './workbook-media'
import { revealCellBelowFreeze } from './univer-sync'
import type { UniverRuntime } from './univer-state'

/** One picture, at the width the pane gives it. A failure leaves the text. */
function Picture({ sessionId, imageId }: { readonly sessionId: string; readonly imageId: string }) {
  const { url, unavailable } = useWorkbookMediaUrl(sessionId, imageId)
  if (url) return <img className="np-image" src={url} alt="" />
  if (unavailable) return <p className="np-image-missing">{unavailable}</p>
  return null
}

function Card({
  picture,
  sessionId,
  onJump,
}: {
  readonly picture: NotePicture
  readonly sessionId: string
  readonly onJump: (row: number, column: number) => void
}) {
  return (
    <article className="np-card">
      <header>
        <button
          type="button"
          className="np-cell"
          onClick={() => onJump(picture.row, picture.column)}
          title={picture.sheetName}
        >
          {picture.sheetName}!{formatAddress(picture.row, picture.column)}
        </button>
        {picture.author && <span className="np-author">{picture.author}</span>}
      </header>
      <Picture sessionId={sessionId} imageId={picture.imageId} />
      {picture.text && <p className="np-text">{picture.text}</p>}
    </article>
  )
}

export function NotePicturesPane({
  file,
  getRuntime,
}: {
  readonly file: WorkbookFile | null
  readonly getRuntime: () => UniverRuntime | null
}): React.JSX.Element | null {
  const { t } = useI18n()
  const pane = useSyncExternalStore(notePicturesPane.subscribe, notePicturesPane.getState)
  const pictures = collectNotePictures(file)
  if (!pane.open) return null

  const sessionId = file?.sessionId
  const jump = (row: number, column: number): void => {
    const sheet = getRuntime()?.univerAPI.getActiveWorkbook()?.getActiveSheet()
    if (!sheet) return
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
            />
          ))
        )}
      </div>
    </aside>
  )
}
