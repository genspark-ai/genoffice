import { describe, expect, it } from 'vitest'

import { applyPageSetupState } from '../src/gateway/xlsx-page-setup'

const SHEET_WITH_VIEW =
  '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
  '<sheetViews><sheetView tabSelected="1" workbookViewId="0"/></sheetViews>' +
  '<sheetData/></worksheet>'

const SHEET_WITHOUT_VIEW =
  '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
  '<sheetData/></worksheet>'

const SHEET_WITH_PANE =
  '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
  '<sheetViews><sheetView tabSelected="1" workbookViewId="0">' +
  '<pane ySplit="3" topLeftCell="A4" activePane="bottomLeft" state="frozen"/>' +
  '<selection pane="bottomLeft" activeCell="A4" sqref="A4"/>' +
  '</sheetView></sheetViews><sheetData/></worksheet>'

describe('applyPageSetupState frozen panes', () => {
  it('writes a frozen pane for rows only', () => {
    const xml = applyPageSetupState(SHEET_WITH_VIEW, {
      sheetName: 'Sheet1',
      frozenRows: 1,
      frozenColumns: 0,
    })
    expect(xml).toContain(
      '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>',
    )
    expect(xml).toContain('<sheetView tabSelected="1" workbookViewId="0">')
  })

  it('writes both splits with a bottomRight active pane', () => {
    const xml = applyPageSetupState(SHEET_WITH_VIEW, {
      sheetName: 'Sheet1',
      frozenRows: 2,
      frozenColumns: 3,
    })
    expect(xml).toContain(
      '<pane xSplit="3" ySplit="2" topLeftCell="D3" activePane="bottomRight" state="frozen"/>',
    )
  })

  it('creates the sheetViews section when missing', () => {
    const xml = applyPageSetupState(SHEET_WITHOUT_VIEW, {
      sheetName: 'Sheet1',
      frozenRows: 0,
      frozenColumns: 1,
    })
    expect(xml).toContain('<sheetViews><sheetView workbookViewId="0">')
    expect(xml).toContain(
      '<pane xSplit="1" topLeftCell="B1" activePane="topRight" state="frozen"/>',
    )
    expect(xml.indexOf('<sheetViews>')).toBeLessThan(xml.indexOf('<sheetData'))
  })

  it('replaces an existing pane and drops pane-scoped selections', () => {
    const xml = applyPageSetupState(SHEET_WITH_PANE, {
      sheetName: 'Sheet1',
      frozenRows: 1,
      frozenColumns: 0,
    })
    expect(xml).not.toContain('ySplit="3"')
    expect(xml).not.toContain('<selection')
    expect(xml).toContain('ySplit="1"')
  })

  it('removes the pane entirely on 0/0', () => {
    const xml = applyPageSetupState(SHEET_WITH_PANE, {
      sheetName: 'Sheet1',
      frozenRows: 0,
      frozenColumns: 0,
    })
    expect(xml).not.toContain('<pane')
    expect(xml).not.toContain('<selection')
    expect(xml).toContain('<sheetView tabSelected="1" workbookViewId="0">')
  })

  it('leaves the sheet untouched when 0/0 and no view exists', () => {
    const xml = applyPageSetupState(SHEET_WITHOUT_VIEW, {
      sheetName: 'Sheet1',
      frozenRows: 0,
      frozenColumns: 0,
    })
    expect(xml).toBe(SHEET_WITHOUT_VIEW)
  })
})
