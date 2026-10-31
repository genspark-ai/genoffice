import { useState, type ReactElement } from 'react'
import { useI18n } from './i18n/locale'
import type { ExternalLinkRow } from '../shared/desktop-api'

/// Excel's Edit Links as a panel: the workbook's `[n]` links with their
/// stored source path, the resolved local path, and the source's sheet names.
/// Reachability is unknown ("—") until the explicit Update: listing the links
/// must not probe the sources on disk (a dead network share can block the
/// open). Update re-reads reachable sources; broken links stay on their
/// cached values (Excel's own behaviour). Break turns the link's formulas
/// into their cached literals (Excel asks for confirmation because the file
/// cannot undo it); Change Source repoints the link at another file without
/// touching formulas or cached values.
export function ExternalLinksPanel({
  links,
  busy,
  onUpdate,
  onReveal,
  onBreak,
  onChangeSource,
  onClose,
}: {
  readonly links: readonly ExternalLinkRow[]
  readonly busy: boolean
  readonly onUpdate: () => void
  readonly onReveal: (source: string) => void
  readonly onBreak: (index: number) => void
  readonly onChangeSource: (index: number) => void
  readonly onClose: () => void
}): ReactElement {
  const { t } = useI18n()
  const [confirming, setConfirming] = useState<number | null>(null)
  return (
    <section className="watch-window" role="dialog" aria-label={t('appLinksPanel')}>
      <header>
        <span className="watch-title">{t('appLinksPanel')}</span>
        <button className="watch-add" disabled={busy} onClick={onUpdate}>
          {t('appLinksUpdate')}
        </button>
        <button className="watch-close" data-tip={t('appClose')} onClick={onClose}>
          ✕
        </button>
      </header>
      {links.length === 0 ? (
        <p className="watch-empty">{t('appLinksNone')}</p>
      ) : (
        <table className="watch-table">
          <thead>
            <tr>
              <th>{t('appLinksSource')}</th>
              <th>{t('appLinksStatus')}</th>
              <th>{t('appLinksSheets')}</th>
              <th />
            </tr>
          </thead>
          <tbody>
            {links.map((link) => (
              <tr key={link.index}>
                <td className="watch-formula">
                  {link.rawTarget}
                  {link.source !== null && link.source !== link.rawTarget ? (
                    <span className="watch-gone"> → {link.source}</span>
                  ) : null}
                </td>
                <td>
                  {link.reachable === undefined
                    ? '—'
                    : link.reachable
                      ? t('appLinksReachable')
                      : t('appLinksBroken')}
                </td>
                <td>{link.sheetNames.join(', ')}</td>
                <td>
                  {link.reachable === true && link.source !== null ? (
                    <button
                      className="watch-remove"
                      data-tip={t('appLinksLocate')}
                      onClick={() => onReveal(link.source!)}
                    >
                      📂
                    </button>
                  ) : null}
                  <button
                    className="watch-remove"
                    data-tip={t('appLinksChangeSource')}
                    disabled={busy}
                    onClick={() => onChangeSource(link.index)}
                  >
                    🔗
                  </button>
                  {confirming === link.index ? (
                    <button
                      className="watch-remove"
                      disabled={busy}
                      onClick={() => {
                        setConfirming(null)
                        onBreak(link.index)
                      }}
                    >
                      {t('appLinksBreakConfirm')}
                    </button>
                  ) : (
                    <button
                      className="watch-remove"
                      data-tip={t('appLinksBreak')}
                      disabled={busy}
                      onClick={() => setConfirming(link.index)}
                    >
                      ✂
                    </button>
                  )}
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      )}
    </section>
  )
}
