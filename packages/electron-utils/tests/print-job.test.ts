import { describe, expect, it } from 'vitest'
import { printJobOptions } from '../src/print-job'

describe('printJobOptions', () => {
  it('keeps the system dialog for null/omitted deviceName and drops its job fields', () => {
    expect(printJobOptions(undefined)).toEqual({ silent: false })
    expect(printJobOptions({ deviceName: null, copies: 3, duplex: 'longEdge' })).toEqual({
      silent: false,
    })
  })

  it("prints silently to the OS default printer for ''", () => {
    expect(
      printJobOptions({ deviceName: '', copies: 2, collate: false, duplex: 'shortEdge' }),
    ).toEqual({
      silent: true,
      copies: 2,
      collate: false,
      duplexMode: 'shortEdge',
    })
  })

  it('names the chosen printer and rejects bad copies/duplex instead of clamping', () => {
    expect(printJobOptions({ deviceName: 'HP', copies: 0, duplex: 'bogus' })).toEqual({
      silent: true,
      deviceName: 'HP',
    })
    expect(printJobOptions({ deviceName: 'HP', copies: 1000 })).toEqual({
      silent: true,
      deviceName: 'HP',
    })
    expect(printJobOptions({ deviceName: 'HP', copies: 2.5 })).toEqual({
      silent: true,
      deviceName: 'HP',
    })
    expect(printJobOptions({ deviceName: 'HP', copies: 1, duplex: 'simplex' })).toEqual({
      silent: true,
      deviceName: 'HP',
      copies: 1,
      duplexMode: 'simplex',
    })
  })

  // Chromium prints on Letter unless the job states the document's own paper,
  // and a sheet taller than the paper is split in two (genoffice#2563)
  it("states the document's paper in the unit Electron takes", () => {
    const opts = printJobOptions({ deviceName: 'HP', pageSizeIn: { width: 8.27, height: 11.69 } })
    expect(opts.pageSize).toEqual({
      width: Math.round(8.27 * 25_400),
      height: Math.round(11.69 * 25_400),
    })
    // A4, not the Letter Chromium prints on by default (8.5in x 11in)
    expect(opts.pageSize?.width).not.toBe(Math.round(8.5 * 25_400))
    expect(opts.pageSize?.height).not.toBe(Math.round(11 * 25_400))
  })

  // the paper is not the system dialog's business: it opens on Letter too
  it('states the paper on the system-dialog route too', () => {
    const opts = printJobOptions({ deviceName: null, pageSizeIn: { width: 8.27, height: 11.69 } })
    expect(opts.pageSize).toEqual({
      width: Math.round(8.27 * 25_400),
      height: Math.round(11.69 * 25_400),
    })
  })

  it('states no paper when the job carries none', () => {
    expect(printJobOptions({ deviceName: 'HP' }).pageSize).toBeUndefined()
    expect(
      printJobOptions({ deviceName: 'HP', pageSizeIn: { width: 0, height: 11 } }).pageSize,
    ).toBeUndefined()
    expect(
      printJobOptions({ deviceName: 'HP', pageSizeIn: { width: 8, height: 900 } }).pageSize,
    ).toBeUndefined()
  })
})
