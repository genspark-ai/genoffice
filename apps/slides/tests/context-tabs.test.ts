import type { RenderNode } from '@genoffice/pptx-render'
import { describe, expect, it } from 'vitest'
import {
  autoContextTabForElement,
  contextElementTypeForNode,
  contextTabForElement,
} from '../src/renderer/components/context-tabs'

describe('slides contextual ribbon tabs', () => {
  it('keeps shape-format commands available for text shapes without auto-switching', () => {
    expect(contextTabForElement('textShape')).toBe('shapeFormat')
    expect(autoContextTabForElement('textShape')).toBeNull()
  })

  it('auto-switches for pictures, ordinary shapes, and other dedicated contextual tools', () => {
    expect(autoContextTabForElement('picture')).toBe('pictureFormat')
    expect(autoContextTabForElement('shape')).toBe('shapeFormat')
    expect(autoContextTabForElement('table')).toBe('tableDesign')
    expect(autoContextTabForElement('chart')).toBe('chartDesign')
  })

  it('keeps picture-format for mixed selections so outline stays available', () => {
    expect(contextTabForElement('mixed')).toBe('pictureFormat')
    expect(autoContextTabForElement('mixed')).toBe('pictureFormat')
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
    expect(autoContextTabForElement(null)).toBeNull()
  })
})
