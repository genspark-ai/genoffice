import { describe, expect, it } from 'vitest'
import JSZip from 'jszip'
import { generateDocx } from '../src/generate'

// 1x1 transparent PNG
const PNG_1PX = Buffer.from(
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==',
  'base64',
)

describe('zero-area image nodes', () => {
  it('floors inline-image run sizes to a finite >= 1px wp:extent', async () => {
    const ir = [
      {
        type: 'para',
        runs: [
          { text: '', inlineImage: true, shotId: 's1', width: 0, height: 0 },
          { text: '', inlineImage: true, shotId: 's1', width: NaN, height: NaN },
        ],
      },
    ]
    const buf = await generateDocx(ir, { s1: PNG_1PX }, {})
    const zip = await JSZip.loadAsync(buf)
    const xml = await zip.file('word/document.xml')!.async('string')
    expect(xml).not.toContain('NaN')
    expect(xml).not.toContain('Infinity')
    const extents = xml.match(/<wp:extent cx="(\d+)" cy="(\d+)"\/>/g) || []
    expect(extents.length).toBe(2)
    for (const extent of extents) {
      const [, cx, cy] = extent.match(/<wp:extent cx="(\d+)" cy="(\d+)"\/>/)!
      expect(Number(cx)).toBeGreaterThanOrEqual(1)
      expect(Number(cy)).toBeGreaterThanOrEqual(1)
    }
  })
})
