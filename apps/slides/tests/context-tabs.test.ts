import type { RenderNode } from '@genoffice/pptx-render'
import { describe, expect, it } from 'vitest'
import {
  contextElementTypeForNode,
  contextTabForElement,
} from '../src/renderer/components/context-tabs'

describe('slides contextual ribbon tabs', () => {
  it('maps every selection type to its dedicated contextual tab', () => {
    expect(contextTabForElement('picture')).toBe('pictureFormat')
    expect(contextTabForElement('shape')).toBe('shapeFormat')
    expect(contextTabForElement('textShape')).toBe('shapeFormat')
    expect(contextTabForElement('table')).toBe('tableDesign')
    expect(contextTabForElement('chart')).toBe('chartDesign')
  })

  it('keeps picture-format for mixed selections so outline stays available', () => {
    expect(contextTabForElement('mixed')).toBe('pictureFormat')
  })

  it('distinguishes text-bearing shapes and groups from ordinary shapes', () => {
    const visibleText = {
      lines: [{ runs: [{ text: 'Title' }] }],
    }
    const textShape = { type: 'shape', text: visibleText } as unknown as RenderNode
    const emptyTextShape = { type: 'shape', text: { lines: [] } } as unknown as RenderNode
    const plainText = { type: 'text', text: visibleText } as unknown as RenderNode
    const textGroup = { type: 'group', children: [plainText] } as unknown as RenderNode

    expect(contextElementTypeForNode(textShape)).toBe('textShape')
    expect(contextElementTypeForNode(emptyTextShape)).toBe('shape')
    expect(contextElementTypeForNode(textGroup)).toBe('textShape')
  })

  it('does not expose a contextual tab without a supported selection', () => {
    expect(contextTabForElement(null)).toBeNull()
  })
})
