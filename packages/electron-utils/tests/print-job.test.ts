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
})
