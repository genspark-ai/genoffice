/**
 * Word's typed line grid (w:docGrid lines/linesAndChars) governs the body flow
 * only: header/footer strip lines keep their natural heights and the reserved
 * push-down probe must not measure grid-snapped strips (prod-sas 003/004 Word
 * baselines: header pitch 13.5pt under a 21.9pt grid, body top at the raw
 * margin while the snapped strip height would have pushed it down a line+).
 */
import { describe, expect, it } from 'vitest'
import type { ParsedDocFull } from '@genoffice/docx-engine'
import { docStyleCss } from '../src/renderer/doc-style-css'

;(globalThis as { CSS?: unknown }).CSS ??= { escape: (s: string) => s }

function parsedWith(sectPrXml: string | null): ParsedDocFull {
  return {
    styles: new Map(),
    docDefaults: {},
    blocks: sectPrXml ? [{ docxIndex: 0, originalXml: `<w:p>${sectPrXml}</w:p>` }] : [],
  } as unknown as ParsedDocFull
}

const GRID_SECT =
  '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:docGrid w:type="lines" w:linePitch="438"/></w:sectPr>'

describe('docStyleCss header/footer strips under a typed grid', () => {
  it('opts strips out of grid snapping with a re-declared line-height', () => {
    const css = docStyleCss(parsedWith(GRID_SECT))
    const rule = /\.doc-page \.page-hf \{[^}]*\}/.exec(css)?.[0]
    expect(rule).toBeTruthy()
    expect(rule).toContain('--doc-grid-pitch:0.0001px')
    expect(rule).toContain('line-height:')
    // auto multiples survive without the grid arm (mirrors .doc-nosnap)
    expect(rule).toContain(
      '--doc-line-max:calc(var(--doc-line-factor,1.2) * 1em * var(--doc-line-mult,1))',
    )
  })

  it('keeps a document-level exact line height on the strips', () => {
    const parsed = parsedWith(GRID_SECT)
    ;(parsed as unknown as { docDefaults: object }).docDefaults = {
      lineRule: 'exact',
      lineRawTwips: 440,
    }
    const css = docStyleCss(parsed)
    const rule = /\.doc-page \.page-hf \{[^}]*\}/.exec(css)?.[0]
    expect(rule).toContain('line-height:22.0pt')
  })

  it('lets mixed-script CJK stretches re-evaluate an atLeast line under the typed grid', () => {
    const parsed = parsedWith(GRID_SECT)
    ;(parsed as unknown as { docDefaults: object }).docDefaults = {
      lineRule: 'atLeast',
      lineRawTwips: 440,
    }
    const css = docStyleCss(parsed)
    const gridSpan = css.indexOf(':not(.doc-lh-fixed) span { line-height:inherit }')
    const stretch = css.indexOf(':not(.doc-lh-fixed) span.doc-run-lf { line-height:')
    expect(gridSpan).toBeGreaterThan(-1)
    // more specific and later than the grid inherit rule, so the stretch wins
    expect(stretch).toBeGreaterThan(gridSpan)
  })

  it('declares the document line factor on .page-wrap for the canvas strips', () => {
    const css = docStyleCss(parsedWith(null))
    const wrap = /\.page-wrap \{ --doc-line-factor:([0-9.]+) \}/.exec(css)
    const page = /\.doc-page, \.pv-page \{[^}]*--doc-line-factor:([0-9.]+)/.exec(css)
    expect(wrap?.[1]).toBeDefined()
    expect(wrap?.[1]).toBe(page?.[1])
  })

  it('emits no strip override without a typed grid', () => {
    const css = docStyleCss(parsedWith(null))
    expect(css).not.toContain('.doc-page .page-hf')
  })
})
