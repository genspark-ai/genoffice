import { describe, expect, it } from 'vitest'
import { deriveAutoFileName, deriveNameFromPrompt } from '../src/renderer/document/auto-name'

describe('deriveAutoFileName', () => {
  it('prefers the document title', () => {
    const html =
      '<html><head><title> Riverside &amp; Books </title></head><body><h1>Hero</h1></body></html>'
    expect(deriveAutoFileName(html)).toBe('Riverside & Books')
  })

  it('falls back to the first h1, stripping inline markup', () => {
    expect(deriveAutoFileName('<body><h1 class="x">A <em>good</em> book</h1></body>')).toBe(
      'A good book',
    )
  })

  it('falls back to the first words of the body', () => {
    const words = Array.from({ length: 12 }, (_, i) => `w${i}`).join(' ')
    expect(deriveAutoFileName(`<body><p>${words}</p></body>`)).toBe('w0 w1 w2 w3 w4 w5 w6 w7')
  })

  it('caps the length', () => {
    expect(deriveAutoFileName(`<title>${'x'.repeat(100)}</title>`)).toHaveLength(60)
  })

  it('leaves out-of-range numeric entities alone instead of throwing', () => {
    expect(deriveAutoFileName('<title>a &#1114112; b &#x110000; c &#65; </title>')).toBe(
      'a &#1114112; b &#x110000; c A',
    )
  })

  it('returns empty for an empty document', () => {
    expect(deriveAutoFileName('')).toBe('')
  })
})

describe('deriveNameFromPrompt', () => {
  it('keeps the first clause of the first sentence and drops the politeness prefix', () => {
    expect(
      deriveNameFromPrompt(
        '\u5e2e\u6211\u8bbe\u8ba1\u4e00\u4e2a\u4e2a\u4eba\u5de5\u4f5c\u53f0\u7684\u754c\u9762\u7a3f\uff0c\u5e38\u7528\u5165\u53e3\u3001\u5f85\u529e\u3001\u65e5\u7a0b\u3001\u6307\u6807\u7b49\u6a21\u5757\u53ef\u81ea\u5b9a\u4e49\u5e03\u5c40\u3002\u4f7f\u7528\u8005\u548c\u573a\u666f\u662f\uff1a\u8bbe\u8ba1/\u521b\u610f',
      ),
    ).toBe('\u8bbe\u8ba1\u4e00\u4e2a\u4e2a\u4eba\u5de5\u4f5c\u53f0\u7684\u754c\u9762\u7a3f')
    expect(deriveNameFromPrompt('Please build a landing page for my bakery, warm tones.')).toBe(
      'Build a landing page for my bakery',
    )
  })

  it('does not break inside the first few characters', () => {
    expect(deriveNameFromPrompt('\u5f85\u529e\uff0c\u65e5\u7a0b\uff0c\u7b14\u8bb0')).toBe(
      '\u5f85\u529e\uff0c\u65e5\u7a0b\uff0c\u7b14\u8bb0',
    )
    expect(
      deriveNameFromPrompt(
        '\u5f85\u529e\uff0c\u65e5\u7a0b\u548c\u7b14\u8bb0\uff0c\u8fd8\u6709\u5929\u6c14',
      ),
    ).toBe('\u5f85\u529e\uff0c\u65e5\u7a0b\u548c\u7b14\u8bb0')
  })

  it('uses only the first line and caps long clauses at a word boundary', () => {
    expect(deriveNameFromPrompt('Dashboard\nwith many details')).toBe('Dashboard')
    const long = 'A single very long request without any punctuation that keeps going on'
    const name = deriveNameFromPrompt(long)
    expect(name.length).toBeLessThanOrEqual(40)
    expect(long.startsWith(name)).toBe(true)
    expect(name.endsWith(' ')).toBe(false)
  })

  it('keeps dotted tokens and returns empty for blank input', () => {
    expect(deriveNameFromPrompt('Rebuild index.html for v2.0 today')).toBe(
      'Rebuild index.html for v2.0 today',
    )
    expect(deriveNameFromPrompt('  \n ')).toBe('')
  })
})
