/**
 * Header/footer strip default font size: Word's Header/Footer styles are based
 * on Normal, so unstyled strip runs take the document default size, both below
 * and above the static 10.5pt guess (a 12pt-default header wraps its long line
 * where Word does).
 */
import { describe, expect, it } from 'vitest'
import type { ParsedDocFull, StyleDisplay, StyleInfo } from '@genoffice/docx-engine'
import { docStyleCss } from '../src/renderer/doc-style-css'

;(globalThis as { CSS?: unknown }).CSS ??= { escape: (s: string) => s }

function parsedWith(opts: {
  ddSizeHalfPoints?: number
  normalSizeHalfPoints?: number
}): ParsedDocFull {
  const styles = new Map<string, StyleInfo>()
  if (opts.normalSizeHalfPoints) {
    styles.set('Normal', {
      styleId: 'Normal',
      name: 'Normal',
      type: 'paragraph',
      isDefault: true,
      display: { sizeHalfPoints: opts.normalSizeHalfPoints } as StyleDisplay,
    } as StyleInfo)
  }
  return {
    styles,
    docDefaults: opts.ddSizeHalfPoints ? { sizeHalfPoints: opts.ddSizeHalfPoints } : {},
    blocks: [],
  } as unknown as ParsedDocFull
}

describe('docStyleCss --hf-default-fs', () => {
  it('emits the strip default size when the document default is below 10.5pt', () => {
    const css = docStyleCss(parsedWith({ ddSizeHalfPoints: 16 }))
    expect(css).toContain('.page-wrap, .doc-page, .pv-page { --hf-default-fs:8pt }')
  })

  it('prefers Normal over docDefaults', () => {
    const css = docStyleCss(parsedWith({ ddSizeHalfPoints: 24, normalSizeHalfPoints: 19 }))
    expect(css).toContain('--hf-default-fs:9.5pt')
  })

  it('grows the strip base above the 10.5pt static guess too', () => {
    expect(docStyleCss(parsedWith({ ddSizeHalfPoints: 21 }))).toContain('--hf-default-fs:10.5pt')
    expect(docStyleCss(parsedWith({ ddSizeHalfPoints: 24 }))).toContain('--hf-default-fs:12pt')
  })

  it("no size anywhere: Word's built-in 10pt default drives body and strips", () => {
    const css = docStyleCss(parsedWith({}))
    expect(css).toContain('font-size:10pt')
    expect(css).toContain('--hf-default-fs:10pt')
  })
})
