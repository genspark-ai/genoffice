/**
 * Chart data edit dialog: editable grid of categories × series (modeled on PowerPoint Edit Data).
 * Values are stashed as strings (tolerating in-progress input) and parsed with Number() on OK;
 * invalid/empty values count as 0.
 */
import { useState } from 'react'
import { useI18n } from '../i18n/locale'

export interface ChartDataInit {
  categories: string[]
  series: Array<{ name: string; values: number[] }>
}

interface Props {
  init: ChartDataInit
  onConfirm: (categories: string[], series: Array<{ name: string; values: number[] }>) => void
  onClose: () => void
}

export function ChartDataDialog({ init, onConfirm, onClose }: Props) {
  const { t } = useI18n()
  const [cats, setCats] = useState<string[]>(() =>
    init.categories.length ? [...init.categories] : [t('paneChartCategoryN', { n: 1 })],
  )
  const [names, setNames] = useState<string[]>(() =>
    init.series.length ? init.series.map((s) => s.name) : [t('paneChartSeriesN', { n: 1 })],
  )
  const [vals, setVals] = useState<string[][]>(() => {
    const rows = init.categories.length || 1
    const sers = init.series.length ? init.series : [{ name: '', values: [] as number[] }]
    return sers.map((s) =>
      Array.from({ length: rows }, (_, i) => (s.values[i] != null ? String(s.values[i]) : '')),
    )
  })

  const setVal = (si: number, ci: number, v: string) =>
    setVals((prev) =>
      prev.map((row, i) => (i === si ? row.map((x, j) => (j === ci ? v : x)) : row)),
    )

  const addCategory = () => {
    setCats((c) => [...c, t('paneChartCategoryN', { n: c.length + 1 })])
    setVals((prev) => prev.map((row) => [...row, '']))
  }
  const removeCategory = (ci: number) => {
    setCats((c) => c.filter((_, i) => i !== ci))
    setVals((prev) => prev.map((row) => row.filter((_, j) => j !== ci)))
  }
  const addSeries = () => {
    setNames((n) => [...n, t('paneChartSeriesN', { n: n.length + 1 })])
    setVals((prev) => [...prev, cats.map(() => '')])
  }
  const removeSeries = (si: number) => {
    setNames((n) => n.filter((_, i) => i !== si))
    setVals((prev) => prev.filter((_, i) => i !== si))
  }

  const confirm = () => {
    // Empty names are legal (kept empty end to end); never substitute placeholders
    const categories = cats.map((c) => c.trim())
    const series = names.map((name, si) => ({
      name: name.trim(),
      values: categories.map((_, ci) => {
        const n = Number(vals[si]?.[ci] ?? '')
        return Number.isFinite(n) ? n : 0
      }),
    }))
    onConfirm(categories, series)
  }

  return (
    <div className="modal-backdrop" onClick={onClose}>
      <div className="modal chart-data-modal" onClick={(e) => e.stopPropagation()}>
        <h2>{t('paneChartEditTitle')}</h2>
        <div className="chart-data-scroll">
          <table className="chart-data-grid">
            <thead>
              <tr>
                <th className="chart-data-corner" />
                {names.map((name, si) => (
                  <th key={si}>
                    <div className="chart-data-ser-head">
                      <input
                        value={name}
                        onChange={(e) =>
                          setNames((n) => n.map((x, i) => (i === si ? e.target.value : x)))
                        }
                      />
                      <button
                        className="chart-data-del"
                        data-tip={t('paneChartDelSeries')}
                        aria-label={t('paneChartDelSeries')}
                        disabled={names.length <= 1}
                        onClick={() => removeSeries(si)}
                      >
                        ×
                      </button>
                    </div>
                  </th>
                ))}
                <th className="chart-data-addcol">
                  <button
                    className="chart-data-add"
                    data-tip={t('paneChartAddSeries')}
                    onClick={addSeries}
                  >
                    {t('paneChartAddSeriesBtn')}
                  </button>
                </th>
              </tr>
            </thead>
            <tbody>
              {cats.map((cat, ci) => (
                <tr key={ci}>
                  <th>
                    <div className="chart-data-ser-head">
                      <input
                        value={cat}
                        onChange={(e) =>
                          setCats((c) => c.map((x, i) => (i === ci ? e.target.value : x)))
                        }
                      />
                      <button
                        className="chart-data-del"
                        data-tip={t('paneChartDelCategory')}
                        aria-label={t('paneChartDelCategory')}
                        disabled={cats.length <= 1}
                        onClick={() => removeCategory(ci)}
                      >
                        ×
                      </button>
                    </div>
                  </th>
                  {names.map((_, si) => (
                    <td key={si}>
                      <input
                        inputMode="decimal"
                        value={vals[si]?.[ci] ?? ''}
                        onChange={(e) => setVal(si, ci, e.target.value)}
                      />
                    </td>
                  ))}
                  <td />
                </tr>
              ))}
              <tr>
                <th className="chart-data-addcol">
                  <button
                    className="chart-data-add"
                    data-tip={t('paneChartAddCategory')}
                    onClick={addCategory}
                  >
                    {t('paneChartAddCategoryBtn')}
                  </button>
                </th>
                <td colSpan={names.length + 1} />
              </tr>
            </tbody>
          </table>
        </div>
        <div className="modal-actions">
          <button className="btn-ghost" onClick={onClose}>
            {t('paneCancel')}
          </button>
          <button className="btn-primary" onClick={confirm}>
            {t('paneOk')}
          </button>
        </div>
      </div>
    </div>
  )
}
