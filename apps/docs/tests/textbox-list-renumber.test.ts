import { describe, expect, it } from 'vitest'
import { textboxOrderedMarkerTexts } from '../src/renderer/editor/extensions'
import type { TextboxListMarker } from '@genoffice/docx-engine'

const item = (ilvl: number, numId = '7'): TextboxListMarker => ({
  text: '',
  numId,
  ilvl,
  kind: 'ordered',
})

describe('textboxOrderedMarkerTexts', () => {
  it('restarts the nested level after every parent item, not only the first', () => {
    const texts = textboxOrderedMarkerTexts([
      item(0),
      item(1),
      item(1),
      item(0),
      item(1),
      item(0),
      item(1),
    ])
    expect(texts).toEqual(['1.', '1.', '2.', '2.', '1.', '3.', '1.'])
  })

  it('skips bullets and plain paragraphs and keeps numIds apart', () => {
    const bullet: TextboxListMarker = { text: '•', numId: '9', ilvl: 0, kind: 'bullet' }
    expect(textboxOrderedMarkerTexts([item(0), null, bullet, item(0, '8'), item(0)])).toEqual([
      '1.',
      null,
      null,
      '1.',
      '2.',
    ])
  })
})
