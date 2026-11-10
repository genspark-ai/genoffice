import { beforeEach, describe, expect, it, vi } from 'vitest'

const electron = vi.hoisted(() => ({
  clipboard: {
    readBuffer: vi.fn(),
    readHTML: vi.fn(),
    write: vi.fn(),
  },
  nativeImage: {
    createFromBuffer: vi.fn(),
  },
}))

vi.mock('electron', () => electron)

import {
  canWriteElementClipboardImage,
  elementClipboardMarkerMatches,
  writeElementClipboardImage,
} from '../src/main/element-clipboard'

const token = 'copy-token-1'

beforeEach(() => {
  electron.clipboard.readBuffer.mockReset().mockReturnValue(Buffer.alloc(0))
  electron.clipboard.readHTML.mockReset().mockReturnValue('')
  electron.clipboard.write.mockReset()
  electron.nativeImage.createFromBuffer.mockReset().mockImplementation((bytes: Buffer) => ({
    isEmpty: () => bytes.length === 0 || bytes[0] === 0,
    toPNG: () => Buffer.from('normalized-png'),
  }))
})

describe('element clipboard OS marker', () => {
  it('only recognizes the exact custom-format token as internal', () => {
    electron.clipboard.readBuffer.mockReturnValue(Buffer.from(token))
    expect(elementClipboardMarkerMatches(token)).toBe(true)

    electron.clipboard.readBuffer.mockReturnValue(Buffer.from(`${token}-other`))
    expect(elementClipboardMarkerMatches(token)).toBe(false)
  })

  it('recognizes the token in the HTML image marker when custom formats are absent', () => {
    electron.clipboard.readHTML.mockReturnValue(
      '<img src="data:image/png;base64,cG5n" data-genoffice-slides-elements="copy-token-1">',
    )

    expect(elementClipboardMarkerMatches(token)).toBe(true)
  })

  it('rejects a stale token and an externally replaced clipboard before image write', () => {
    electron.clipboard.readBuffer.mockReturnValue(Buffer.from('new-token'))
    const current = { token: 'new-token', senderId: 42 }

    expect(canWriteElementClipboardImage(current, 42, token)).toBe(false)
    expect(canWriteElementClipboardImage(current, 7, 'new-token')).toBe(false)
    expect(canWriteElementClipboardImage(current, 42, 'new-token')).toBe(true)

    electron.clipboard.readBuffer.mockReturnValue(Buffer.from('external-copy'))
    expect(canWriteElementClipboardImage(current, 42, 'new-token')).toBe(false)
  })
})

describe('element clipboard PNG writer', () => {
  it('writes normalized PNG and HTML marker in one clipboard operation', () => {
    const decoded = { isEmpty: () => false, toPNG: () => Buffer.from('normalized-png') }
    electron.nativeImage.createFromBuffer.mockReturnValue(decoded)
    electron.clipboard.readBuffer.mockReturnValue(Buffer.from('copy token'))

    expect(writeElementClipboardImage('copy token', 'cG5n')).toBe(true)

    expect(electron.clipboard.write).toHaveBeenCalledOnce()
    expect(electron.clipboard.write).toHaveBeenCalledWith({
      image: decoded,
      html: '<img src="data:image/png;base64,bm9ybWFsaXplZC1wbmc=" data-genoffice-slides-elements="copy%20token">',
    })
    expect(electron.nativeImage.createFromBuffer).toHaveBeenCalledOnce()
  })

  it('does not overwrite an external copy made while PNG normalization runs', () => {
    electron.clipboard.readBuffer.mockReturnValue(Buffer.from(token))
    electron.nativeImage.createFromBuffer.mockReturnValue({
      isEmpty: () => false,
      toPNG: () => {
        electron.clipboard.readBuffer.mockReturnValue(Buffer.from('external-copy'))
        return Buffer.from('normalized-png')
      },
    })

    expect(writeElementClipboardImage(token, 'cG5n')).toBe(false)
    expect(electron.clipboard.write).not.toHaveBeenCalled()
  })

  it('leaves the clipboard untouched for an empty or undecodable PNG', () => {
    expect(writeElementClipboardImage(token, '')).toBe(false)
    expect(writeElementClipboardImage(token, 'AA==')).toBe(false)

    expect(electron.clipboard.write).not.toHaveBeenCalled()
  })

  it('leaves the clipboard untouched when PNG decoding throws', () => {
    electron.nativeImage.createFromBuffer.mockImplementationOnce(() => {
      throw new Error('invalid PNG')
    })

    expect(writeElementClipboardImage(token, 'cG5n')).toBe(false)
    expect(electron.clipboard.write).not.toHaveBeenCalled()
  })
})
