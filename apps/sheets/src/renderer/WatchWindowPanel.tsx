import { useEffect, useState } from 'react'

import { formatAddress } from '../domain/cell-address'
import { useI18n } from './i18n/locale'

/// One watched cell, addressed by stable ids so sheet renames survive.
export interface WatchCell {
  readonly sheetId: string
  readonly row: number
  readonly column: number
}

export interface WatchRowValue {
  readonly sheetName: string
  readonly value: string
  readonly formula: string
}

export function watchKey(cell: WatchCell): string {
  return `${cell.sheetId}:${cell.row}:${cell.column}`
}

/// Excel's Watch Window: a floating, non-modal panel pinning cells whose
/// value/formula update live. Values are polled while open — recalc has no
/// single completion event the panel could subscribe to from here, and the
/// handful of watched cells make polling effectively free.
export function WatchWindowPanel({
  watches,
  onResolve,
  onAddSelection,
  onRemove,
  onClose,
}: {
  readonly watches: readonly WatchCell[]
  /// null = the sheet is gone; the row renders struck through.
  readonly onResolve: (cell: WatchCell) => WatchRowValue | null
  readonly onAddSelection: () => void
  readonly onRemove: (key: string) => void
  readonly onClose: () => void
}): React.JSX.Element {
  const { t } = useI18n()
  const [, setTick] = useState(0)
  useEffect(() => {
    const timer = setInterval(() => setTick((value) => value + 1), 800)
    return () => clearInterval(timer)
  }, [])
  return (
    <section className="watch-window" role="dialog" aria-label={t('appWatchWindow')}>
      <header>
        <span className="watch-title">{t('appWatchWindow')}</span>
        <button className="watch-add" onClick={onAddSelection}>
          {t('appWatchAdd')}
        </button>
        <button className="watch-close" data-tip={t('appClose')} onClick={onClose}>
          ✕
        </button>
      </header>
      {watches.length === 0 ? (
        <p className="watch-empty">{t('appWatchEmpty')}</p>
      ) : (
        <table className="watch-table">
          <thead>
            <tr>
              <th>{t('appWatchSheet')}</th>
              <th>{t('appWatchCell')}</th>
              <th>{t('appWatchValue')}</th>
              <th>{t('appWatchFormula')}</th>
              <th />
            </tr>
          </thead>
          <tbody>
            {watches.map((cell) => {
              const resolved = onResolve(cell)
              return (
                <tr key={watchKey(cell)} className={resolved ? '' : 'watch-gone'}>
                  <td>{resolved?.sheetName ?? '—'}</td>
                  <td>{formatAddress(cell.row, cell.column)}</td>
                  <td>{resolved?.value ?? ''}</td>
                  <td className="watch-formula">{resolved?.formula ?? ''}</td>
                  <td>
                    <button
                      className="watch-remove"
                      data-tip={t('appDeleteLabel')}
                      onClick={() => onRemove(watchKey(cell))}
                    >
                      ✕
                    </button>
                  </td>
                </tr>
              )
            })}
          </tbody>
        </table>
      )}
    </section>
  )
}
