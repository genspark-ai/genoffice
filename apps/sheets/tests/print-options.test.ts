import { describe, expect, it } from 'vitest'
import { printOptionsFor } from '../src/main/print-options'
import type { WorkbookExportPdfRequest } from '../src/shared/desktop-api'

const REQUEST: WorkbookExportPdfRequest = {
  fileName: 'Book1.pdf',
  html: '<html></html>',
  landscape: true,
  pageSize: 'A4',
  margins: { top: 0.75, bottom: 0.75, left: 0.7, right: 0.7 },
  scale: 0.8,
}

describe('printOptionsFor', () => {
  it('carries the page setup over in the units webContents.print expects', () => {
    expect(printOptionsFor(REQUEST)).toEqual({
      silent: false,
      printBackground: true,
      landscape: true,
      pageSize: 'A4',
      margins: { marginType: 'custom', top: 72, bottom: 72, left: 67, right: 67 },
      scaleFactor: 80,
    })
  })

  it('converts a custom paper size from inches to microns', () => {
    expect(printOptionsFor({ ...REQUEST, pageSize: { width: 8.5, height: 14 } }).pageSize).toEqual({
      width: 215900,
      height: 355600,
    })
  })
})
