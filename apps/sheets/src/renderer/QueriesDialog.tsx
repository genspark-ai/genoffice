import { useEffect, useState } from 'react'

import type { WorkbookQueriesResult } from '../shared/desktop-api'
import { useI18n } from './i18n/locale'
import { useModalDialog } from './modal-dialog'

type Query = WorkbookQueriesResult[number]

const KIND_KEY = {
  powerQuery: 'dlgQueriesKindPowerQuery',
  web: 'dlgQueriesKindWeb',
  text: 'dlgQueriesKindText',
  database: 'dlgQueriesKindDatabase',
  other: 'dlgQueriesKindOther',
} as const

/// `'My Sheet'!A1:C21`, the form Go To resolves
function qualifiedRange(destination: NonNullable<Query['destination']>): string {
  return `'${destination.sheet.replaceAll("'", "''")}'!${destination.range}`
}

/// Data › Queries & Connections, read-only (genoffice#2051): the workbook's
/// Power Query queries and data connections, and where each one's rows are
/// loaded. Clicking a loaded query selects its rows. Refresh is not offered:
/// there is no query engine yet, and the parts stay in the file for Excel.
export function QueriesDialog({
  load,
  onGo,
  onClose,
}: {
  readonly load: () => Promise<WorkbookQueriesResult>
  /// Returns an error message, or null on success (Go To's resolver).
  readonly onGo: (ref: string) => string | null
  readonly onClose: () => void
}): React.JSX.Element {
  const { t } = useI18n()
  const [queries, setQueries] = useState<WorkbookQueriesResult | null>(null)
  const [error, setError] = useState<string | null>(null)

  useEffect(() => {
    let alive = true
    load().then(
      (result) => alive && setQueries(result),
      (err: unknown) => alive && setError(err instanceof Error ? err.message : String(err)),
    )
    return () => {
      alive = false
    }
    // load once per open; the dialog is remounted for the next workbook
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [])

  const go = (destination: NonNullable<Query['destination']>): void => {
    const failure = onGo(qualifiedRange(destination))
    if (failure === null) onClose()
    else setError(failure)
  }

  const modal = useModalDialog(onClose)
  return (
    <div className="dialog-backdrop" onClick={onClose}>
      <div
        className="format-cells-dialog"
        role="dialog"
        {...modal}
        aria-label={t('dlgQueriesTitle')}
        onClick={(event) => event.stopPropagation()}
      >
        <header>{t('dlgQueriesTitle')}</header>
        <section className="dialog-body">
          {queries?.length === 0 && <p className="dialog-note">{t('dlgQueriesEmpty')}</p>}
          {queries && queries.length > 0 && (
            <ul className="goto-name-list">
              {queries.map((query, index) => {
                const destination = query.destination
                const detail = [
                  t(KIND_KEY[query.kind]),
                  destination
                    ? t('dlgQueriesLoaded', {
                        rows: destination.rows,
                        target: `${destination.table ?? destination.sheet} (${qualifiedRange(destination)})`,
                      })
                    : t('dlgQueriesConnectionOnly'),
                  query.source,
                ]
                  .filter(Boolean)
                  .join(' · ')
                return (
                  <li key={`${index}:${query.name}`}>
                    <button
                      type="button"
                      disabled={!destination}
                      data-tip={destination ? t('dlgQueriesGoHint') : query.description}
                      onClick={() => destination && go(destination)}
                    >
                      <strong>{query.name}</strong>
                      <span>{detail}</span>
                    </button>
                  </li>
                )
              })}
            </ul>
          )}
          {queries && queries.length > 0 && <p className="dialog-note">{t('dlgQueriesNote')}</p>}
          {error && (
            <p className="dialog-note dialog-error" role="alert">
              {t('dlgQueriesError', { error })}
            </p>
          )}
        </section>
        <div className="dialog-actions">
          <button className="primary-action" autoFocus onClick={onClose}>
            {t('dlgClose')}
          </button>
        </div>
      </div>
    </div>
  )
}
