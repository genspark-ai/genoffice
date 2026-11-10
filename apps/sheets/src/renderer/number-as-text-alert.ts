/**
 * Univer's "number stored as text" cell alerts (text-format cell holding a
 * number, or a force-string cell) are hardcoded to ERROR: red icon, "Error"
 * title. The cell renders correctly, Excel treats it as a hint, so present it
 * as a warning.
 */
import { CellAlertManagerService, CellAlertType, type ICellAlert } from '@univerjs/sheets-ui'

const NUMBER_AS_TEXT_ALERT_KEYS = new Set(['SHEET_NUMFMT_ALERT', 'SHEET_FORCE_STRING_ALERT'])

export function numberAsTextAlertType(alert: Pick<ICellAlert, 'key' | 'type'>): CellAlertType {
  return NUMBER_AS_TEXT_ALERT_KEYS.has(alert.key) && alert.type === CellAlertType.ERROR
    ? CellAlertType.WARNING
    : alert.type
}

let installed = false

export function installNumberAsTextAlertSeverity(): void {
  if (installed) return
  installed = true
  const proto = CellAlertManagerService.prototype as { showAlert(alert: ICellAlert): void }
  const origShowAlert = proto.showAlert
  proto.showAlert = function (alert: ICellAlert) {
    return origShowAlert.call(this, { ...alert, type: numberAsTextAlertType(alert) })
  }
}
