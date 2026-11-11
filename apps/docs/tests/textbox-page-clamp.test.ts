/**
 * Word keeps floating drawing objects on the page: a column-relative posOffset
 * X hanging past a paper edge is pulled back on at display time (e.g.
 * a Word-authored rectangle at -4629150 EMU drew 9.9cm left of the page and
 * cut across the body text). Page-frame offsets stay raw, and the authored
 * offset is never rewritten.
 */
import { describe, expect, it } from 'vitest'
import type { TextboxDisplay } from '@genoffice/docx-engine'
import { textboxBoxStyle } from '../src/renderer/editor/protected-render'

const box = (extra: Partial<TextboxDisplay>): TextboxDisplay => ({
  paras: [],
  floating: true,
  widthPx: 560,
  heightPx: 345,
  ...extra,
})

describe('textboxBoxStyle on-page clamp', () => {
  it('clamps a column-relative X between the page edges', () => {
    const style = textboxBoxStyle(box({ offsetXEmu: -4629150 }))
    expect(style).toContain(
      'left:clamp(calc(0px - var(--doc-margin-left,0px)), -486.0px, ' +
        'calc(var(--doc-content-w,100%) - 560px + var(--doc-margin-right,0px)))',
    )
  })

  it('an on-page offset passes through the clamp unchanged', () => {
    // 853440 EMU = 89.6px, well inside the clamp range: the box position is
    // unaffected, only the CSS carries the guard
    expect(textboxBoxStyle(box({ offsetXEmu: 853440 }))).toContain(', 89.6px,')
  })

  it('page-frame offsets (pagePinned / pageRelX) stay raw', () => {
    expect(textboxBoxStyle(box({ offsetXEmu: -4629150, pagePinned: true }))).toContain(
      'left:-486.0px',
    )
    expect(textboxBoxStyle(box({ offsetXEmu: -4629150, pageRelX: true }))).toContain(
      'left:-486.0px',
    )
  })

  it('cell-anchored boxes skip the clamp (the zero-width .doc-cell-boxes host resolves 100% to 0)', () => {
    const style = textboxBoxStyle(box({ offsetXEmu: 853440 }), { inCell: true })
    expect(style).toContain('left:89.6px')
    expect(style).not.toContain('clamp')
  })

  it('a width-less box cannot compute the right cap and keeps the raw offset', () => {
    expect(textboxBoxStyle(box({ offsetXEmu: -4629150, widthPx: undefined }))).toContain(
      'left:-486.0px',
    )
  })

  it('non-floating boxes get no position at all', () => {
    expect(textboxBoxStyle(box({ floating: false, offsetXEmu: -4629150 }))).not.toContain(
      'position:absolute',
    )
  })
})

describe('textboxBoxStyle vertical anchor and picture geometry', () => {
  it('anchor=ctr/b use safe alignment so overflowing text starts at the box top (Word)', () => {
    expect(textboxBoxStyle(box({ vAlign: 'center' }))).toContain(
      'display:flex;flex-direction:column;justify-content:safe center',
    )
    expect(textboxBoxStyle(box({ vAlign: 'bottom' }))).toContain('justify-content:safe flex-end')
  })

  it('a photo box with a preset geometry clips the picture to the outline', () => {
    const style = textboxBoxStyle(
      box({
        fillImageDataUrl: 'data:image/png;base64,AA==',
        prst: 'ellipse',
        widthPx: 100,
        heightPx: 50,
      }),
    )
    expect(style).toContain('clip-path:path("M 0 25 A 50 25 0 1 1 100 25 A 50 25 0 1 1 0 25 Z")')
    expect(style).not.toContain('background-image:url("data:image/svg')
  })

  it('a photo box without geometry stays rectangular', () => {
    expect(textboxBoxStyle(box({ fillImageDataUrl: 'data:image/png;base64,AA==' }))).not.toContain(
      'clip-path',
    )
  })
})
