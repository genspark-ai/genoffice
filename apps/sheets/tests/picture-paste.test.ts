import { describe, expect, it } from 'vitest'

import {
  fitPictureFrame,
  pickTransferImage,
  pictureFileProblem,
  pointInRect,
  transferHasImage,
  type TransferItemLike,
  type TransferLike,
} from '../src/renderer/picture-paste'

function imageItem(type = 'image/png', withFile = true): TransferItemLike {
  const file = withFile ? ({ name: 'image.png', type, size: 10 } as unknown as File) : null
  return { kind: 'file', type, getAsFile: () => file }
}

function stringItem(type: string): TransferItemLike {
  return { kind: 'string', type, getAsFile: () => null }
}

function transfer(items: TransferItemLike[], data: Record<string, string> = {}): TransferLike {
  return { items, getData: (format) => data[format] ?? '' }
}

describe('pickTransferImage', () => {
  it('takes a bare screenshot', () => {
    const file = pickTransferImage(transfer([imageItem()]))
    expect(file?.type).toBe('image/png')
  })

  it('takes a browser "Copy image" (bitmap plus an <img> fragment)', () => {
    const data = transfer([stringItem('text/html'), imageItem()], {
      'text/html': '<meta charset="utf-8"><img src="https://example.test/a.png" alt="a">',
    })
    expect(pickTransferImage(data)).not.toBeNull()
  })

  it('leaves spreadsheet cells (table markup plus a rendered bitmap) to the cell paste', () => {
    const data = transfer([stringItem('text/plain'), stringItem('text/html'), imageItem()], {
      'text/plain': '1\t2\n3\t4',
      'text/html': '<table><tr><td>1</td><td>2</td></tr></table>',
    })
    expect(pickTransferImage(data)).toBeNull()
  })

  it('leaves plain text with a stray bitmap to the cell paste', () => {
    const data = transfer([stringItem('text/plain'), imageItem('image/tiff')], {
      'text/plain': 'hello',
    })
    expect(pickTransferImage(data)).toBeNull()
  })

  it('ignores transfers without an image file', () => {
    expect(
      pickTransferImage(transfer([stringItem('text/plain')], { 'text/plain': 'x' })),
    ).toBeNull()
    expect(pickTransferImage(transfer([]))).toBeNull()
    expect(transferHasImage(transfer([stringItem('text/html')]))).toBe(false)
  })

  it('reports an image during drag-over even before the file is readable', () => {
    expect(transferHasImage(transfer([imageItem('image/jpeg', false)]))).toBe(true)
    expect(pickTransferImage(transfer([imageItem('image/jpeg', false)]))).toBeNull()
  })
})

describe('pictureFileProblem', () => {
  it('accepts png/jpeg/gif within 20 MB and normalises image/jpg', () => {
    expect(pictureFileProblem({ size: 1024, type: 'image/png' })).toBeNull()
    expect(pictureFileProblem({ size: 1024, type: 'image/jpg' })).toBeNull()
    expect(pictureFileProblem({ size: 20 * 1024 * 1024, type: 'image/gif' })).toBeNull()
  })

  it('rejects oversize and unsupported files', () => {
    expect(pictureFileProblem({ size: 20 * 1024 * 1024 + 1, type: 'image/png' })).toBe('too-large')
    expect(pictureFileProblem({ size: 10, type: 'image/webp' })).toBe('bad-type')
    expect(pictureFileProblem({ size: 10, type: 'image/tiff' })).toBe('bad-type')
  })
})

describe('fitPictureFrame', () => {
  it('maps natural pixels onto the 80x22 cell grid', () => {
    expect(fitPictureFrame(400, 220)).toEqual({ columns: 5, rows: 10 })
  })

  it('scales wide images down to a 480px frame', () => {
    expect(fitPictureFrame(1920, 1080)).toEqual({ columns: 6, rows: 12 })
  })

  it('clamps tiny and tall images', () => {
    expect(fitPictureFrame(1, 1)).toEqual({ columns: 2, rows: 2 })
    expect(fitPictureFrame(100, 5000)).toEqual({ columns: 2, rows: 40 })
  })
})

describe('pointInRect', () => {
  const rect = { left: 10, top: 20, right: 110, bottom: 220 }

  it('accepts points inside and on the leading edges', () => {
    expect(pointInRect(10, 20, rect)).toBe(true)
    expect(pointInRect(50, 100, rect)).toBe(true)
  })

  it('rejects points outside and on the trailing edges', () => {
    expect(pointInRect(9, 100, rect)).toBe(false)
    expect(pointInRect(110, 100, rect)).toBe(false)
    expect(pointInRect(50, 220, rect)).toBe(false)
    expect(pointInRect(50, 19, rect)).toBe(false)
  })
})
