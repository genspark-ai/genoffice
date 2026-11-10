import { describe, expect, it } from 'vitest'
import { pinnedFloatPage, type PageSlice } from '../src/renderer/pagination'
import { readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import {
  hoistPaperClamp,
  hoistSlotOf,
  hoistWindowCss,
  stripLeadingTranslate,
  pinnedCloneCss,
} from '../src/renderer/components/PaginationPreview'

const slice = (start: number, end: number): PageSlice => ({ start, end, section: 0 })

describe('pinnedFloatPage', () => {
  const slices = [slice(0, 900), slice(900, 1800), slice(1800, 2700)]

  it('maps an anchor to its containing slice', () => {
    expect(pinnedFloatPage(slices, 0)).toBe(0)
    expect(pinnedFloatPage(slices, 899.9)).toBe(0)
    expect(pinnedFloatPage(slices, 900)).toBe(1)
    expect(pinnedFloatPage(slices, 2000)).toBe(2)
  })

  it('clamps out-of-range anchors', () => {
    expect(pinnedFloatPage(slices, -5)).toBe(0)
    expect(pinnedFloatPage(slices, 2700)).toBe(2)
    expect(pinnedFloatPage(slices, 99999)).toBe(2)
    expect(pinnedFloatPage([], 100)).toBe(0)
  })
})

describe('pinnedCloneCss', () => {
  const page = (index: number, html: string): HTMLElement => {
    const el = document.createElement('div')
    el.className = 'pv-page'
    el.dataset.pvPage = String(index)
    el.innerHTML = html
    return el
  }
  const pinned = (pin?: number) =>
    `<div class="doc-protected doc-protected-floating doc-protected-pagepinned"${
      pin === undefined ? '' : ` data-pin-page="${pin}"`
    }><div class="doc-textbox"></div></div>`
  const boxes = (el: HTMLElement): HTMLElement[] =>
    Array.from(el.querySelectorAll('.doc-protected-pagepinned'))
  const matches = (root: HTMLElement, css: string): boolean[] => {
    const selectors = css
      .split('\n')
      .filter(Boolean)
      .map((r) => r.slice(0, r.indexOf('{')))
    return boxes(root).map((b) => selectors.some((s) => b.matches(s)))
  }

  it('hides clones owned by other pages, keeps the owning page copy', () => {
    const css = pinnedCloneCss(2)
    // visibility, not display: stray-run wrappers keep their measured flow height
    expect(css).toContain('visibility:hidden')
    expect(css).not.toContain('display')
    // full-document clone: every page carries both covers (pin 0 and pin 1)
    const p0 = page(0, pinned(0) + pinned(1) + '<p>body</p>')
    const p1 = page(1, pinned(0) + pinned(1) + '<p>body</p>')
    expect(matches(p0, css)).toEqual([false, true])
    expect(matches(p1, css)).toEqual([true, false])
  })

  it('never targets unstamped wrappers', () => {
    const css = pinnedCloneCss(3)
    const p = page(2, pinned() + pinned(2) + pinned(0))
    expect(matches(p, css)).toEqual([false, false, true])
  })

  it('emits no rules for an empty preview', () => {
    expect(pinnedCloneCss(0)).toBe('')
  })

  it('hides hoisted spill-float boxes on non-owning pages, box only (not the stray line)', () => {
    const css = pinnedCloneCss(2)
    const hoisted = (pin: number) =>
      `<div class="doc-protected doc-protected-floating" data-pv-hoist="1" data-pin-page="${pin}">` +
      `<div class="doc-anchor-strut"><br></div><div class="doc-textbox"></div></div>`
    const p0 = page(0, hoisted(0) + hoisted(1))
    const selectors = css
      .split('\n')
      .filter(Boolean)
      .flatMap((r) => r.slice(0, r.indexOf('{')).split(','))
    const hidden = (root: HTMLElement, sel: string) =>
      Array.from(root.querySelectorAll(sel)).map((b) => selectors.some((s) => b.matches(s)))
    // the escaped boxes hide on the foreign page's clone; the flow strut stays
    expect(hidden(p0, '.doc-textbox')).toEqual([false, true])
    expect(hidden(p0, '.doc-anchor-strut')).toEqual([false, false])
  })
})

describe('hoistSlotOf', () => {
  const box = (style: string) => {
    const el = document.createElement('span')
    el.setAttribute('style', style)
    return { el }
  }

  it('classifies the inline horizontal slot of a floating box', () => {
    expect(hoistSlotOf(box('position:absolute;left:12.5px;top:3px'))).toBe('left')
    expect(hoistSlotOf(box('position:absolute;left:-4px;top:0px'))).toBe('left')
    // a centred cover logo (left:50% + translateX(-50%)) used to be skipped and
    // spilled past the page window, repainting on the next page's top
    expect(
      hoistSlotOf(box('position: absolute; left: 50%; top: 5.8px; transform: translateX(-50%);')),
    ).toBe('center')
    // paper-edge X (relativeFrom="page") pictures render the margin as a calc()
    expect(
      hoistSlotOf(box('position:absolute;left:calc(447.3px - var(--doc-margin-left,0px));top:2px')),
    ).toBe('left')
    expect(hoistSlotOf(box('position:absolute;right:0;top:2px'))).toBe('right')
    expect(hoistSlotOf(box('position:absolute;right:0px;top:2px'))).toBe('right')
    expect(hoistSlotOf(box('position:absolute;right:20px;top:2px'))).toBeNull()
  })

  it('reads the slot off the owning .box-rotator when the box is wrapped', () => {
    // the chrome rotator owns the float slot style; the box itself carries no
    // left/clamp/right any more, so the carrier lookup decides the slot
    const rotator = document.createElement('div')
    rotator.className = 'box-rotator'
    rotator.setAttribute('style', 'position:absolute;left:clamp(0px, 12.5px, 200px);top:3px')
    const inner = document.createElement('span')
    inner.className = 'doc-textbox'
    inner.setAttribute('style', 'background-color:#fff;width:189px;height:40px')
    rotator.appendChild(inner)
    expect(hoistSlotOf({ el: inner })).toBe('left')

    // a box without a rotator parent still classifies from its own style
    expect(hoistSlotOf(box('position:absolute;left:12.5px;top:3px'))).toBe('left')
  })
})

describe('hoistPaperClamp', () => {
  it('pushes a box that would start above the paper down to the page top', () => {
    // anchor at the body top (dy 0, top margin 56.7px), box 94.5px above it
    expect(hoistPaperClamp(0, 56.7, [-94.5])).toBeCloseTo(37.8, 5)
    // the highest box decides: siblings ride the same shift
    expect(hoistPaperClamp(10, 56.7, [-94.5, -20])).toBeCloseTo(37.8, 5)
  })

  it('leaves boxes that stay on the paper alone', () => {
    expect(hoistPaperClamp(0, 56.7, [-56.7])).toBe(0)
    expect(hoistPaperClamp(120, 56.7, [-94.5])).toBe(120)
    expect(hoistPaperClamp(5, 56.7, [])).toBe(5)
  })
})

describe('banded picture width', () => {
  const css = readFileSync(
    join(dirname(fileURLToPath(import.meta.url)), '../src/renderer/styles.css'),
    'utf8',
  )
  it('keeps the wp:extent of a banded picture overhanging the column', () => {
    const rule =
      /\.doc-img-float\.img-wrap-band > \.doc-img-wrap,\s*\.doc-img-float\.img-wrap-band > \.doc-img-wrap > \.doc-protected-img\s*\{([^}]*)\}/.exec(
        css,
      )
    expect(rule).not.toBeNull()
    expect(rule![1]).toMatch(/max-width:\s*none/)
  })
})

describe('stripLeadingTranslate', () => {
  it('drops the anchor translate, nested calc() included, and keeps rotation/flip', () => {
    expect(
      stripLeadingTranslate(
        'translate(calc(447.4px - var(--doc-margin-left,0px) - var(--page-rel-x-undo,0px)), 0.0px) rotate(90deg) scaleX(-1)',
      ),
    ).toBe('rotate(90deg) scaleX(-1)')
    expect(stripLeadingTranslate('translate(96.0px, 0.0px)')).toBe('')
    expect(stripLeadingTranslate('rotate(90deg)')).toBe('rotate(90deg)')
  })
})

describe('hoistWindowCss (prod corpus 20261020 sample 026, regioned page)', () => {
  it("hides a hoisted wrapper's boxes in the body windows that do not own its anchor", () => {
    const css = hoistWindowCss(new Map([[1, new Set(['1040.2'])]]))
    expect(css).toContain(
      '.pv-page[data-pv-page="1"] .pv-content[data-pv-win]:not([data-pv-win="1040.2"]) [data-pv-hoist=\'1\'][data-pin-win="1040.2"] > .doc-textbox',
    )
    expect(css).toContain('> .doc-img-wrap{visibility:hidden;}')
    // the stray anchor line of the wrapper is never targeted
    expect(css).not.toMatch(/data-pin-win="1040\.2"\]\{/)
  })

  it('emits nothing without hoisted wrappers', () => {
    expect(hoistWindowCss(new Map())).toBe('')
  })
})
