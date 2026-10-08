import { describe, expect, it } from 'vitest'
import type {
  ParaBorderSides,
  ParsedDocFull,
  StyleDisplay,
  StyleInfo,
} from '@genoffice/docx-engine'
import { docStyleCss } from '../src/renderer/doc-style-css'

;(globalThis as { CSS?: unknown }).CSS ??= { escape: (s: string) => s }

const DD: ParaBorderSides = {
  t: { none: true, spacePt: 30 },
  l: { none: true, spacePt: 31 },
  b: { none: true, spacePt: 30 },
  r: { none: true, spacePt: 31 },
}

function parsedWith(
  docDefaultsSides: ParaBorderSides | undefined,
  extra: Record<string, StyleDisplay> = {},
): ParsedDocFull {
  const styles = new Map<string, StyleInfo>()
  const merged = (own?: ParaBorderSides): StyleDisplay => ({
    borderSides: { ...docDefaultsSides, ...own },
  })
  styles.set('Normal', {
    styleId: 'Normal',
    name: 'Normal',
    type: 'paragraph',
    isDefault: true,
    display: merged(),
  } as StyleInfo)
  for (const [id, own] of Object.entries(extra)) {
    styles.set(id, {
      styleId: id,
      name: id,
      type: 'paragraph',
      basedOn: 'Normal',
      display: merged(own.borderSides),
    } as StyleInfo)
  }
  return {
    styles,
    docDefaults: docDefaultsSides ? { borderSides: docDefaultsSides } : {},
    blocks: [],
  } as unknown as ParsedDocFull
}

const CELL_NORMAL = '.doc-page :is(td, th) :is(p, h1, h2, h3, h4, h5, h6):not([data-style])'

describe('docDefaults none-side w:space inside table cells', () => {
  it('unstyled cell paragraphs drop the inherited top/bottom padding, body keeps it', () => {
    const css = docStyleCss(parsedWith(DD))
    expect(css).toContain('p:not([data-style])')
    expect(css).toContain('padding-top:30pt;padding-bottom:30pt')
    expect(css).toContain(`${CELL_NORMAL} { padding-top:0;padding-bottom:0 }`)
  })

  it('a style keeps its own side and drops only the inherited one', () => {
    const css = docStyleCss(
      parsedWith(DD, {
        H2: { borderSides: { b: { color: 'EEEEEE', szPt: 0.75, spacePt: 3 } } },
        Web: { borderSides: { t: { none: true, spacePt: 0 }, b: { none: true, spacePt: 0 } } },
      }),
    )
    expect(css).toContain('.doc-page :is(td, th) [data-style="H2"] { padding-top:0 }')
    expect(css).not.toContain(':is(td, th) [data-style="Web"]')
  })

  it('a Normal-level none space that is not from docDefaults still pads cells', () => {
    const css = docStyleCss(parsedWith(undefined, {}))
    expect(css).not.toContain(':is(td, th)')
    const parsed = parsedWith(undefined)
    parsed.styles.get('Normal')!.display = { borderSides: DD }
    expect(docStyleCss(parsed)).not.toContain(':is(td, th)')
  })
})
