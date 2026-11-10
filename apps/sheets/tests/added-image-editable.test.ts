import { describe, expect, it } from 'vitest'

import {
  isEditableAddedImage,
  isEditableFileVisual,
  isEditableShape,
} from '../src/renderer/WorkbookVisuals'
import type { WorkbookVisualObject } from '../src/shared/desktop-api'

const anchor = {
  fromRow: 1,
  fromColumn: 1,
  toRow: 8,
  toColumn: 5,
  fromRowOffset: 0,
  fromColumnOffset: 0,
  toRowOffset: 0,
  toColumnOffset: 0,
}

const pasted: WorkbookVisualObject = {
  id: 'added-image-abc-1',
  sheetId: 'sheet-1',
  kind: 'image',
  anchor,
  mediaType: 'image/png',
  mediaDataUrl: 'data:image/png;base64,AAAA',
  name: 'screenshot.png',
}

describe('session-added pictures', () => {
  it('get the movable/deletable frame even without a drawing locator', () => {
    expect(isEditableAddedImage(pasted)).toBe(true)
    expect(isEditableFileVisual(pasted)).toBe(false)
    expect(isEditableShape(pasted)).toBe(false)
  })

  it('leaves file pictures and session shapes on their own gates', () => {
    expect(
      isEditableAddedImage({
        ...pasted,
        id: 'image-1',
        drawingPath: 'xl/drawings/drawing1.xml',
        drawingIndex: 0,
      }),
    ).toBe(false)
    expect(isEditableAddedImage({ ...pasted, id: 'added-shape-abc-1', kind: 'shape' })).toBe(false)
  })
})
