/**
 * A style-level atLeast line is a minimum, not a cap (genoffice#2806). Runs must
 * keep their natural font-relative boxes like they do under a direct atLeast
 * rule; inheriting the paragraph's absolute line-height made a 14pt minimum
 * render shorter than the same rule declared directly.
 */
import type { DocDefaults, ParsedDocFull, StyleInfo } from '@genoffice/docx-engine'
import { describe, expect, it } from 'vitest'
import { docStyleCss } from '../src/renderer/doc-style-css'

;(globalThis as { CSS?: unknown }).CSS ??= { escape: (s: string) => s }

function styleCss(rule: 'atLeast' | 'exact'): string {
  const styles = new Map<string, StyleInfo>()
  styles.set('BigLine', {
    styleId: 'BigLine',
    name: 'Big Line',
    type: 'paragraph',
    display: { lineRule: rule, lineRawTwips: 280 },
  } as StyleInfo)
  const doc = { styles, docDefaults: {} as DocDefaults, blocks: [] } as unknown as ParsedDocFull
  return docStyleCss(doc)
}

const SPAN = /\.doc-page \[data-style="BigLine"\]:not\(\.doc-lh-fixed\) span \{ ([^}]*) \}/

describe('inherited line rules on style runs', () => {
  it('lets atLeast runs keep their natural height', () => {
    const css = styleCss('atLeast')
    const decl = SPAN.exec(css)?.[1]
    expect(decl).toContain('line-height:min(calc(var(--doc-line-factor, 1.2) * 1em)')
    expect(decl).not.toContain('line-height:inherit')
    expect(css).toContain('--doc-line-fixed:1')
  })

  it('still caps exact runs on the fixed line box', () => {
    const css = styleCss('exact')
    expect(SPAN.exec(css)?.[1]).toContain('line-height:inherit')
    expect(css).toContain('--doc-line-fixed:1')
  })
})
