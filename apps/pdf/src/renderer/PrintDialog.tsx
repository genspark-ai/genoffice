/**
 * The PDF print dialog (#2461's split of the giant renderer components):
 * printer/copies/duplex + the queue's driver option groups, then the page
 * range and odd/even subset. Every field comes straight off usePrintJob's
 * return value — the dialog owns no state of its own.
 */
import { useI18n } from './i18n/locale'
import { PRINT_DUPLEX_MODES, PRINT_SYSTEM_DIALOG, type PrintJobApi } from './usePrintJob'
import type { PageParity } from './print'
import type { PrintDuplex } from '../shared/ipc'

const IS_MAC = navigator.platform.toLowerCase().includes('mac')

export function PrintDialog({
  print,
  printCurrentPage,
}: {
  print: PrintJobApi
  printCurrentPage: number
}) {
  const { t } = useI18n()
  return (
    <div className="pdf-modal-mask" onClick={() => print.setPrintDlg(false)}>
      <div className="pdf-modal" onClick={(e) => e.stopPropagation()}>
        <div className="pdf-modal-title">{t('print')}</div>
        <label className="pdf-modal-row pdf-print-field">
          <span>{t('printPrinter')}</span>
          <select
            className="pdf-modal-input"
            value={print.printDevice}
            onChange={(e) => {
              const request = ++print.driverOptionsRequest.current
              print.setPrintDevice(e.target.value)
              print.setDriverChoices({})
              print.setDriverOptions([])
              if (e.target.value && IS_MAC)
                void window.pdfApi
                  .printerOptions(e.target.value)
                  .then((options) => {
                    if (request === print.driverOptionsRequest.current)
                      print.setDriverOptions(options)
                  })
                  .catch(() => {
                    if (request === print.driverOptionsRequest.current) print.setDriverOptions([])
                  })
            }}
          >
            <option value="">{t('printDefaultPrinter')}</option>
            {print.printers.map((p) => (
              <option key={p.name} value={p.name}>
                {p.displayName}
              </option>
            ))}
            <option value={PRINT_SYSTEM_DIALOG}>{t('printSystemDialog')}</option>
          </select>
        </label>
        <div className="pdf-modal-row pdf-print-field">
          <label className="pdf-print-copies">
            <span>{t('printCopies')}</span>
            <input
              className="pdf-modal-input"
              type="number"
              min={1}
              max={999}
              value={print.printCopies}
              disabled={print.printDevice === PRINT_SYSTEM_DIALOG}
              onChange={(e) =>
                print.setPrintCopies(
                  Math.min(999, Math.max(1, Math.round(Number(e.target.value)) || 1)),
                )
              }
            />
          </label>
          <label className="pdf-modal-check">
            <input
              type="checkbox"
              checked={print.printCollate}
              disabled={print.printDevice === PRINT_SYSTEM_DIALOG || print.printCopies < 2}
              onChange={(e) => print.setPrintCollate(e.target.checked)}
            />
            {t('printCollate')}
          </label>
        </div>
        <label className="pdf-modal-row pdf-print-field">
          <span>{t('printDuplex')}</span>
          <select
            className="pdf-modal-input"
            value={print.printDuplex}
            disabled={print.printDevice === PRINT_SYSTEM_DIALOG}
            onChange={(e) => print.setPrintDuplex(e.target.value as PrintDuplex)}
          >
            {PRINT_DUPLEX_MODES.map((m) => (
              <option key={m} value={m}>
                {t(
                  m === 'simplex'
                    ? 'printDuplexOff'
                    : m === 'longEdge'
                      ? 'printDuplexLong'
                      : 'printDuplexShort',
                )}
              </option>
            ))}
          </select>
        </label>
        {print.driverOptions.map((group) => (
          <label key={group.name} className="pdf-modal-row pdf-print-field">
            <span>{group.translation}</span>
            <select
              className="pdf-modal-input"
              value={print.driverChoices[group.name] ?? group.defaultValue ?? group.values[0]}
              onChange={(e) =>
                print.setDriverChoices((c) => ({ ...c, [group.name]: e.target.value }))
              }
            >
              {group.values.map((v) => (
                <option key={v} value={v}>
                  {v}
                </option>
              ))}
            </select>
          </label>
        ))}
        <div className="pdf-print-divider" />
        <label className="pdf-modal-row">
          <input
            type="radio"
            name="print-range"
            checked={print.printMode === 'all'}
            onChange={() => {
              print.setPrintMode('all')
              print.setPrintInvalid(false)
            }}
          />
          <span>{t('printRangeAll')}</span>
        </label>
        <label className="pdf-modal-row">
          <input
            type="radio"
            name="print-range"
            checked={print.printMode === 'current'}
            onChange={() => {
              print.setPrintMode('current')
              print.setPrintInvalid(false)
            }}
          />
          <span>
            {t('printRangeCurrent')} ({printCurrentPage})
          </span>
        </label>
        <label className="pdf-modal-row">
          <input
            type="radio"
            name="print-range"
            checked={print.printMode === 'custom'}
            onChange={() => print.setPrintMode('custom')}
          />
          <span>{t('printRangeCustom')}</span>
        </label>
        {print.printMode === 'custom' && (
          <input
            className={`pdf-modal-input${print.printInvalid ? ' invalid' : ''}`}
            value={print.printInput}
            placeholder={t('printRangeHint')}
            autoFocus
            onChange={(e) => {
              print.setPrintInput(e.target.value)
              print.setPrintInvalid(false)
            }}
            onKeyDown={(e) => {
              if (e.key === 'Enter') print.confirmPrint()
              else if (e.key === 'Escape') print.setPrintDlg(false)
            }}
          />
        )}
        <label className="pdf-modal-row pdf-print-field">
          <span>{t('printSubset')}</span>
          <select
            className="pdf-modal-input"
            value={print.printParity}
            onChange={(e) => {
              print.setPrintParity(e.target.value as PageParity)
              print.setPrintInvalid(false)
            }}
          >
            <option value="all">{t('printSubsetAll')}</option>
            <option value="odd">{t('printSubsetOdd')}</option>
            <option value="even">{t('printSubsetEven')}</option>
          </select>
        </label>
        {print.printInvalid && print.printMode !== 'custom' && (
          <div className="pdf-modal-hint pdf-print-empty">{t('printSubsetEmpty')}</div>
        )}
        <div className="pdf-modal-actions">
          <button className="pdf-modal-btn" onClick={() => print.setPrintDlg(false)}>
            {t('cancel')}
          </button>
          <button className="pdf-modal-btn primary" onClick={print.confirmPrint}>
            {t('print')}
          </button>
        </div>
      </div>
    </div>
  )
}
