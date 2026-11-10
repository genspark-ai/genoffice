import { describe, expect, it } from 'vitest'

import { oleCaption, oleFrameStyle, oleRenderKind } from '../src/renderer/ole-visual'
import { isEditableFileVisual } from '../src/renderer/WorkbookVisuals'
import { workbookFileSchema, type WorkbookVisualObject } from '../src/shared/desktop-api'

const visualSchema = workbookFileSchema.shape.visuals.element

const anchor = {
  fromRow: 1,
  fromColumn: 1,
  toRow: 3,
  toColumn: 4,
  fromRowOffset: 0,
  fromColumnOffset: 0,
  toRowOffset: 156_210,
  toColumnOffset: 384_810,
  explicitTo: true,
}

describe('ole visual wire shape', () => {
  it('accepts the sidecar ole record with its progId and preview media', () => {
    const parsed = visualSchema.parse({
      id: 'ole-1',
      sheetId: 'sheet-1',
      kind: 'ole',
      anchor,
      progId: 'Word.Document.12',
      mediaPath: 'xl/media/image1.emf',
      mediaType: 'image/x-emf',
    })
    expect(parsed.kind).toBe('ole')
    expect(parsed.progId).toBe('Word.Document.12')
    expect(parsed.mediaPath).toBe('xl/media/image1.emf')
  })

  it('is never routed through the drawing edit pipeline', () => {
    const ole: WorkbookVisualObject = {
      id: 'ole-1',
      sheetId: 'sheet-1',
      kind: 'ole',
      anchor,
      progId: 'Package',
    }
    expect(isEditableFileVisual(ole)).toBe(false)
    // Even a stray locator must not make it movable/deletable.
    expect(
      isEditableFileVisual({ ...ole, drawingPath: 'xl/drawings/drawing1.xml', drawingIndex: 0 }),
    ).toBe(false)
  })
})

describe('oleRenderKind', () => {
  it('prefers the cached preview picture when the file has one', () => {
    expect(oleRenderKind({ mediaPath: 'xl/media/image1.emf' }, false)).toBe('preview')
  })

  it('falls back to the icon+caption box without a preview or when it fails', () => {
    expect(oleRenderKind({}, false)).toBe('placeholder')
    expect(oleRenderKind({ mediaPath: 'xl/media/image1.emf' }, true)).toBe('placeholder')
  })
})

describe('oleCaption', () => {
  it('maps well-known progIds to their friendly names, ignoring versions', () => {
    expect(oleCaption('Word.Document.12')).toBe('Microsoft Word Document')
    expect(oleCaption('Word.Document.8')).toBe('Microsoft Word Document')
    expect(oleCaption('Excel.Sheet.12')).toBe('Microsoft Excel Worksheet')
    expect(oleCaption('Worksheet')).toBe('Microsoft Excel Worksheet')
    expect(oleCaption('PowerPoint.Show.12')).toBe('Microsoft PowerPoint Presentation')
    expect(oleCaption('Acrobat Document')).toBe('Adobe Acrobat Document')
    expect(oleCaption('AcroExch.Document.11')).toBe('Adobe Acrobat Document')
    expect(oleCaption('Package')).toBe('Package')
  })

  it('keeps unknown progIds minus the numeric version suffix', () => {
    expect(oleCaption('Vendor.Widget.3')).toBe('Vendor.Widget')
    expect(oleCaption('CorelDRAW.Graphic.21')).toBe('CorelDRAW.Graphic')
    expect(oleCaption('  Custom  ')).toBe('Custom')
  })

  it('reads as a generic embedded object without a progId', () => {
    expect(oleCaption(undefined)).toBe('Embedded Object')
    expect(oleCaption('')).toBe('Embedded Object')
  })
})

describe('oleFrameStyle', () => {
  it("draws Excel's hairline frame and opaque backdrop from the VML stroke/fill", () => {
    expect(oleFrameStyle({ lineColor: '#000000', fillColor: '#FFFFFF' })).toEqual({
      border: '1px solid #000000',
      background: '#FFFFFF',
    })
  })

  it('leaves the frame off for stroked="f" / filled="f" shapes', () => {
    expect(oleFrameStyle({})).toEqual({})
    expect(oleFrameStyle({ lineColor: 'none', fillColor: 'none' })).toEqual({})
    expect(oleFrameStyle({ lineColor: '#1F4E79' })).toEqual({ border: '1px solid #1F4E79' })
  })
})
