/**
 * Header/footer strip default font size: Word's Header/Footer styles are based
 * on Normal, so unstyled strip runs take the document default size. The var is
 * shrink-only — strips price their line boxes with the document-wide script
 * strut, which overshoots Word on Latin-only strip lines, so a default above
 * the static 10.5pt guess must not grow the push-down reserve
 * (prod_082 8pt regression / prod_008+091 guard).
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

  it('stays silent at or above the 10.5pt static guess', () => {
    expect(docStyleCss(parsedWith({ ddSizeHalfPoints: 21 }))).not.toContain('--hf-default-fs')
    expect(docStyleCss(parsedWith({ ddSizeHalfPoints: 22 }))).not.toContain('--hf-default-fs')
    expect(docStyleCss(parsedWith({}))).not.toContain('--hf-default-fs')
  })
})
