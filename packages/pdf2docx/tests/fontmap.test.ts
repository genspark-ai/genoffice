/** Output-font substitution unit tests (P21 A): pure mapping, injected installed-check. */
import { describe, expect, it } from 'vitest'
import { applyOutputFontSubstitutions, resolveOutputFamily } from '../src/rebuild/fontmap'
import type { IrPage, Line, Span, TextBlock } from '../src/ir'

const installedSet = (...families: string[]) => {
  const norm = (s: string) => s.toLowerCase().replace(/[\s\-_]/g, '')
  const set = new Set(families.map(norm))
  return (family: string) => set.has(norm(family))
}

describe('resolveOutputFamily', () => {
  it('keeps installed families untouched', () => {
    const installed = installedSet('Helvetica', 'Times New Roman')
    expect(resolveOutputFamily('Helvetica', installed)).toBe('Helvetica')
    expect(resolveOutputFamily('Times New Roman', installed)).toBe('Times New Roman')
  })

  it('maps missing classics to the first installed metric-compatible stand-in', () => {
    const mac = installedSet('Helvetica', 'Arial', 'Times New Roman', 'Courier New', 'Times')
    expect(resolveOutputFamily('Nimbus Roman No9 L', mac)).toBe('Times New Roman')
    expect(resolveOutputFamily('Helvetica World', mac)).toBe('Helvetica')
    expect(resolveOutputFamily('Nimbus Mono L', mac)).toBe('Courier New')
    const linux = installedSet('Liberation Serif', 'Liberation Sans', 'Liberation Mono')
    expect(resolveOutputFamily('Times New Roman', linux)).toBe('Liberation Serif')
    expect(resolveOutputFamily('Helvetica', linux)).toBe('Liberation Sans')
    expect(resolveOutputFamily('Arial', linux)).toBe('Liberation Sans')
  })

  it('resolves styled aliases through their stripped base name', () => {
    const mac = installedSet('Helvetica Neue', 'Helvetica')
    expect(resolveOutputFamily('Helvetica Neue LTStd It', mac)).toBe('Helvetica Neue')
    expect(resolveOutputFamily('Arial Bold', installedSet('Arial'))).toBe('Arial')
  })

  it('leaves unknown missing families alone (CJK included)', () => {
    const mac = installedSet('Helvetica')
    expect(resolveOutputFamily('Universal Std Newswith Comm Pi', mac)).toBe(
      'Universal Std Newswith Comm Pi',
    )
    expect(resolveOutputFamily('SimSun', mac)).toBe('SimSun')
    expect(resolveOutputFamily('', mac)).toBe('')
  })
})

const span = (fontFamily: string): Span => ({
  text: 'x',
  box: { x0: 0, y0: 0, x1: 10, y1: 10 },
  fontSize: 10,
  fontFamily,
  bold: false,
  italic: false,
  color: '000000',
  dir: 'ltr',
  script: 'latin',
})

const lineOf = (s: Span): Line => ({
  spans: [s],
  box: s.box,
  baseline: 8,
  endsWithHyphen: false,
})

const blockOf = (s: Span): TextBlock => ({
  kind: 'text',
  lines: [lineOf(s)],
  box: s.box,
  align: 'left',
  firstLineIndentPt: 0,
  dir: 'ltr',
})

describe('applyOutputFontSubstitutions', () => {
  it('rewrites spans in flow blocks, table cells and footnotes in place', () => {
    const flow = span('Helvetica World')
    const cell = span('Nimbus Roman No9 L')
    const note = span('Helvetica World')
    const page: IrPage = {
      index: 0,
      widthPt: 612,
      heightPt: 792,
      rotation: 0,
      blocks: [
        blockOf(flow),
        {
          kind: 'table',
          box: { x0: 0, y0: 0, x1: 100, y1: 100 },
          colWidthsPt: [100],
          rows: [
            [{ box: { x0: 0, y0: 0, x1: 100, y1: 100 }, gridSpan: 1, blocks: [blockOf(cell)] }],
          ],
        },
      ],
      degraded: false,
      scanned: false,
      hasStructTree: false,
      footnotes: [{ id: 'n1', blocks: [blockOf(note)] }],
    }
    const installed = installedSet('Helvetica', 'Times New Roman')
    applyOutputFontSubstitutions([page], [], installed)
    expect(flow.fontFamily).toBe('Helvetica')
    expect(cell.fontFamily).toBe('Times New Roman')
    expect(note.fontFamily).toBe('Helvetica')
  })
})
