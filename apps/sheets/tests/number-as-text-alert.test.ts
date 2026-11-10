import { CellAlertManagerService, CellAlertType, type ICellAlert } from '@univerjs/sheets-ui'
import { describe, expect, it, vi } from 'vitest'

import {
  installNumberAsTextAlertSeverity,
  numberAsTextAlertType,
} from '../src/renderer/number-as-text-alert'

const location = { unitId: 'u', subUnitId: 's', row: 0, col: 0 }

function alert(key: string, type: CellAlertType): ICellAlert {
  return { key, type, title: 't', message: 'm', location, width: 200, height: 74 }
}

describe('numberAsTextAlertType', () => {
  it('downgrades both number-stored-as-text alerts to a warning', () => {
    expect(numberAsTextAlertType(alert('SHEET_NUMFMT_ALERT', CellAlertType.ERROR))).toBe(
      CellAlertType.WARNING,
    )
    expect(numberAsTextAlertType(alert('SHEET_FORCE_STRING_ALERT', CellAlertType.ERROR))).toBe(
      CellAlertType.WARNING,
    )
  })

  it('leaves other alerts alone', () => {
    expect(numberAsTextAlertType(alert('SOME_OTHER_ALERT', CellAlertType.ERROR))).toBe(
      CellAlertType.ERROR,
    )
    expect(numberAsTextAlertType(alert('SHEET_NUMFMT_ALERT', CellAlertType.INFO))).toBe(
      CellAlertType.INFO,
    )
  })
})

describe('installNumberAsTextAlertSeverity', () => {
  it('rewrites the type before the original showAlert runs', () => {
    const proto = CellAlertManagerService.prototype as unknown as {
      showAlert(alert: ICellAlert): void
    }
    const orig = vi.fn()
    proto.showAlert = orig
    installNumberAsTextAlertSeverity()
    installNumberAsTextAlertSeverity()
    proto.showAlert(alert('SHEET_NUMFMT_ALERT', CellAlertType.ERROR))
    proto.showAlert(alert('OTHER', CellAlertType.ERROR))
    expect(orig.mock.calls.map(([a]) => (a as ICellAlert).type)).toEqual([
      CellAlertType.WARNING,
      CellAlertType.ERROR,
    ])
  })
})
