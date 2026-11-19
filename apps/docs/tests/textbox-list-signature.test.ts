/**
 * The textbox paragraph signature must react to list identity alone: toggling
 * a list on an untouched paragraph (same runs, same align) has to reach the
 * byte patch. Marker TEXT stays out of the signature on purpose — per-box
 * counters are recomputed on both sides.
 */
import { describe, expect, it } from 'vitest'
import { textboxParaSignature } from '../src/renderer/editor/convert'
import type { TextboxDisplay } from '@genoffice/docx-engine'

type Para = TextboxDisplay['paras'][number]

const para = (over: Partial<Para> = {}): Para => ({ runs: [], ...over })

describe('textboxParaSignature list identity', () => {
  it('distinguishes the same paragraph with and without a list', () => {
    const plain = para()
    const listed = para({ listMarker: { text: '•', numId: '7', ilvl: 0, kind: 'bullet' } })
    expect(textboxParaSignature(plain)).not.toEqual(textboxParaSignature(listed))
  })

  it('distinguishes numId and ilvl, ignores marker text drift', () => {
    const a = para({ listMarker: { text: '1.', numId: '7', ilvl: 0, kind: 'ordered' } })
    const b = para({ listMarker: { text: '3.', numId: '7', ilvl: 0, kind: 'ordered' } })
    const c = para({ listMarker: { text: '1.', numId: '8', ilvl: 0, kind: 'ordered' } })
    const d = para({ listMarker: { text: '1.', numId: '7', ilvl: 1, kind: 'ordered' } })
    expect(textboxParaSignature(a)).toEqual(textboxParaSignature(b))
    expect(textboxParaSignature(a)).not.toEqual(textboxParaSignature(c))
    expect(textboxParaSignature(a)).not.toEqual(textboxParaSignature(d))
  })
})

describe('textboxParaSignature indent', () => {
  it('reacts to a moved left indent on otherwise identical paragraphs', () => {
    expect(textboxParaSignature(para())).not.toEqual(
      textboxParaSignature(para({ indentLeft: 720 })),
    )
  })
})
