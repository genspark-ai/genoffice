import { describe, expect, it } from 'vitest'

import { externalFileTarget } from '../src/renderer/external-link-target'

describe('externalFileTarget', () => {
  it('keeps the Windows drive colon (Excel writes file:///C:/Data/src.xlsx)', () => {
    expect(externalFileTarget('C:\\Data\\src.xlsx')).toBe('file:///C:/Data/src.xlsx')
    expect(externalFileTarget('D:/My Data/source.xlsx')).toBe('file:///D:/My%20Data/source.xlsx')
  })

  it('never percent-encodes the drive colon', () => {
    // the sidecar's resolver checks for the drive letter before decoding, so
    // C%3A would be read as a POSIX path and the link would show broken
    expect(externalFileTarget('C:\\Data\\src.xlsx')).not.toContain('%3A')
  })

  it('encodes POSIX path segments and keeps the absolute path', () => {
    expect(externalFileTarget('/Users/me/My Book/source.xlsx')).toBe(
      'file:///Users/me/My%20Book/source.xlsx',
    )
  })

  it('keeps the UNC spelling', () => {
    expect(externalFileTarget('\\\\server\\share\\src.xlsx')).toBe('file:////server/share/src.xlsx')
  })
})
