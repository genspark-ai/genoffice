import type { TFunc } from './i18n/locale'
import type { CompareResult } from './compare-text'

/**
 * View → Compare: the open document against a picked one, pages aligned by
 * index. Read-only — rows jump to the page in the open document, nothing here
 * touches the file.
 */
export function ComparePanel({
  t,
  result,
  otherName,
  busy,
  onPick,
  onSwap,
  onGoToPage,
}: {
  t: TFunc
  result: CompareResult | null
  otherName: string | null
  busy: boolean
  onPick: () => void
  /** treat the picked document as the old side instead of the new one */
  onSwap: () => void
  onGoToPage: (pageIndex: number) => void
}): React.JSX.Element {
  const differing = result?.pages.filter((p) => p.addedCount > 0 || p.removedCount > 0) ?? []
  return (
    <div className="pdf-compare">
      <button
        type="button"
        className="pdf-modal-btn primary pdf-compare-pick"
        onClick={onPick}
        disabled={busy}
      >
        {busy ? t('loading') : t('comparePick')}
      </button>
      {otherName && (
        <button type="button" className="pdf-compare-swap" onClick={onSwap}>
          {t('compareSwap')}
        </button>
      )}
      {result && (
        <div className="pdf-modal-hint">
          {result.equal
            ? t('compareIdentical')
            : t('compareDiffPages', { count: differing.length })}
          {otherName && <div className="pdf-compare-other">{otherName}</div>}
        </div>
      )}
      <ul className="pdf-compare-list">
        {differing.map((page) => (
          <li key={page.pageIndex}>
            <button
              type="button"
              className="pdf-compare-row"
              onClick={() => onGoToPage(page.pageIndex + 1)}
            >
              <span className="pdf-compare-page">
                {t('comparePage', { page: page.pageIndex + 1 })}
              </span>
              {page.truncated ? (
                <span className="pdf-compare-truncated">{t('compareTruncated')}</span>
              ) : (
                <>
                  {page.addedCount > 0 && (
                    <span className="pdf-compare-added">
                      {t('compareAdded', { count: page.addedCount })}
                    </span>
                  )}
                  {page.removedCount > 0 && (
                    <span className="pdf-compare-removed">
                      {t('compareRemoved', { count: page.removedCount })}
                    </span>
                  )}
                </>
              )}
            </button>
          </li>
        ))}
      </ul>
    </div>
  )
}
