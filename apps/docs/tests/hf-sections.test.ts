import { describe, expect, it } from 'vitest'
import type { HeaderFooter, HfPartInfo, SectionInfo } from '@genoffice/docx-engine'
import { EMPTY_HF_VARIANTS } from '../src/renderer/doc-state'
import {
  editKey,
  hfCommitTarget,
  hfLinked,
  hfSaveOptions,
  linkKey,
  ownHf,
  parseEditKey,
  resolveHf,
  withHfEdit,
  withHfLink,
  type HfSectionState,
} from '../src/renderer/hf-sections'

const hf = (text: string): HeaderFooter => ({ text, paras: [{ runs: [{ text }] }] })
const part = (text: string): HfPartInfo => ({
  text,
  hasPageNumber: false,
  paras: [{ runs: [{ text }] }],
})

function section(
  lastBlockIndex: number,
  headerRefs: SectionInfo['headerRefs'] = {},
  footerRefs: SectionInfo['footerRefs'] = {},
): SectionInfo {
  return {
    settings: {} as SectionInfo['settings'],
    startType: 'nextPage',
    firstBlockIndex: 0,
    lastBlockIndex,
    sectPrXml: '',
    titlePg: false,
    headerRefs,
    footerRefs,
  }
}

/** three sections: 1 owns a header (+ first page), 2 inherits, 3 (final) owns another */
function threeSections(): HfSectionState {
  return {
    sections: [
      section(10, { default: 'rH1', first: 'rF1' }),
      section(20),
      section(30, { default: 'rH3' }),
    ],
    edits: {},
    links: {},
    docStrips: { header: hf('doc header'), footer: null, variants: { ...EMPTY_HF_VARIANTS } },
    parsed: {
      hfParts: {
        rH1: { ...part('first section'), images: [{ dataUrl: 'logo' }] },
        rF1: part('first page of section one'),
        rH3: part('third section'),
      },
    },
  }
}

/** one saved section with a header; a break inserted mid-document copies its sectPr */
function pendingBreak(): HfSectionState {
  const saved = section(40, { default: 'rH' })
  return {
    sections: [{ ...saved, lastBlockIndex: 12, pendingBreak: true }, saved],
    edits: {},
    links: {},
    parsed: { hfParts: { rH: { ...part('report header'), images: [{ dataUrl: 'logo' }] } } },
  }
}

const apply = (s: HfSectionState, next: Pick<HfSectionState, 'edits' | 'links'> | null) => {
  expect(next).not.toBeNull()
  return { ...s, ...next! }
}

describe('resolveHf', () => {
  it('a section without its own part inherits the nearest earlier one, images included', () => {
    const s = threeSections()
    expect(resolveHf(s, 1, 'header', 'default')).toMatchObject({
      value: { text: 'first section' },
      owner: 0,
    })
    expect(resolveHf(s, 1, 'header', 'default').images).toEqual([{ dataUrl: 'logo' }])
    expect(resolveHf(s, 2, 'header', 'default')).toMatchObject({
      value: { text: 'third section' },
      owner: 2,
    })
    expect(hfLinked(s, 1, 'header', 'default')).toBe(true)
    expect(hfLinked(s, 2, 'header', 'default')).toBe(false)
    expect(hfLinked(s, 0, 'header', 'default')).toBeNull()
  })

  it('first/even variants inherit too and edit into their own typed slot', () => {
    const s = threeSections()
    expect(resolveHf(s, 2, 'header', 'first')).toMatchObject({
      value: { text: 'first page of section one' },
      owner: 0,
    })
    expect(hfCommitTarget(s, 2, 'header', 'first')).toEqual({ key: '10:header:first', owner: 0 })
    const t = apply(s, withHfEdit(s, 2, 'header', 'first', hf('edited first page')))
    expect(resolveHf(t, 2, 'header', 'first').value?.text).toBe('edited first page')
    expect(resolveHf(t, 0, 'header', 'first').value?.text).toBe('edited first page')
  })

  it('pending edits and Link to Previous win over the parsed references', () => {
    const s = threeSections()
    s.edits[editKey(20, 'header', 'default')] = hf('second, edited')
    expect(resolveHf(s, 1, 'header', 'default')).toMatchObject({
      value: { text: 'second, edited' },
      owner: 1,
    })
    s.links[linkKey(30, 'header', 'default')] = true
    expect(resolveHf(s, 2, 'header', 'default')).toMatchObject({
      value: { text: 'second, edited' },
      owner: 1,
    })
    expect(hfLinked(s, 2, 'header', 'default')).toBe(true)
    expect(ownHf(s, 2, 'header', 'default')).toBeUndefined()
  })

  it('a lone section reads the document-level parse and edits its own slot', () => {
    const s: HfSectionState = { ...threeSections(), sections: [section(5)] }
    expect(resolveHf(s, 0, 'header', 'default')).toMatchObject({
      value: { text: 'doc header' },
      owner: 0,
    })
    expect(hfCommitTarget(s, 0, 'footer', 'even')).toEqual({ key: '5:footer:even', owner: 0 })
    const t = apply(s, withHfEdit(s, 0, 'header', 'default', hf('typed')))
    expect(resolveHf(t, 0, 'header', 'default').value?.text).toBe('typed')
    expect(hfSaveOptions(t)).toEqual({ header: hf('typed') })
  })
})

describe('sections sharing one part (pending breaks, copied sectPr)', () => {
  it('the new first section resolves to the header the whole document had, images included', () => {
    const s = pendingBreak()
    const first = resolveHf(s, 0, 'header', 'default')
    const second = resolveHf(s, 1, 'header', 'default')
    expect(first).toMatchObject({ value: { text: 'report header' }, owner: 0 })
    expect(second.value).toEqual(first.value)
    expect(second.images).toEqual(first.images)
    expect(hfLinked(s, 1, 'header', 'default')).toBe(false)
    // the break paragraph deleted again: the header stays put
    expect(resolveHf({ ...s, sections: [s.sections[1]] }, 0, 'header', 'default').value).toEqual(
      first.value,
    )
  })

  it('editing either half shows on both; the edit is saved through the owner', () => {
    const s = pendingBreak()
    const t = apply(s, withHfEdit(s, 1, 'header', 'default', hf('edited on the last half')))
    expect(resolveHf(t, 0, 'header', 'default').value?.text).toBe('edited on the last half')
    expect(resolveHf(t, 1, 'header', 'default').value?.text).toBe('edited on the last half')
    expect(hfSaveOptions(t)).toEqual({ header: hf('edited on the last half') })
    const u = apply(s, withHfEdit(s, 0, 'header', 'default', hf('edited on the first half')))
    expect(resolveHf(u, 1, 'header', 'default').value?.text).toBe('edited on the first half')
    expect(hfSaveOptions(u)).toEqual({
      sectionHf: [
        {
          lastBlockIndex: 12,
          kind: 'header',
          variant: 'default',
          hf: hf('edited on the first half'),
        },
      ],
    })
  })

  it('edit the shared strip, then Link to Previous on the last section: the edit persists and saves', () => {
    const s = pendingBreak()
    const edited = apply(s, withHfEdit(s, 0, 'header', 'default', hf('shared, edited')))
    const linked = apply(edited, withHfLink(edited, 1, 'header', 'default', true))
    expect(resolveHf(linked, 1, 'header', 'default')).toMatchObject({
      value: { text: 'shared, edited' },
      owner: 0,
    })
    expect(resolveHf(linked, 0, 'header', 'default').value?.text).toBe('shared, edited')
    expect(hfSaveOptions(linked)).toEqual({
      sectionHf: [
        { lastBlockIndex: 12, kind: 'header', variant: 'default', hf: hf('shared, edited') },
      ],
      sectionHfUnlink: [{ lastBlockIndex: 40, kind: 'header', variant: 'default' }],
    })
  })

  it('pending break, then Link to Previous on the last section drops only its own reference', () => {
    const s = pendingBreak()
    const linked = apply(s, withHfLink(s, 1, 'header', 'default', true))
    expect(resolveHf(linked, 1, 'header', 'default')).toMatchObject({
      value: { text: 'report header' },
      owner: 0,
    })
    expect(hfSaveOptions(linked)).toEqual({
      sectionHfUnlink: [{ lastBlockIndex: 40, kind: 'header', variant: 'default' }],
    })
  })
})

describe('Link to Previous', () => {
  it('unlink → edit → relink: the own part goes, the previous section is untouched', () => {
    const s = threeSections()
    const unlinked = apply(s, withHfLink(s, 1, 'header', 'default', false))
    expect(hfLinked(unlinked, 1, 'header', 'default')).toBe(false)
    expect(resolveHf(unlinked, 1, 'header', 'default')).toMatchObject({
      value: { text: 'first section' },
      owner: 1,
    })
    const edited = apply(unlinked, withHfEdit(unlinked, 1, 'header', 'default', hf('mine')))
    expect(resolveHf(edited, 1, 'header', 'default').value?.text).toBe('mine')
    expect(resolveHf(edited, 0, 'header', 'default').value?.text).toBe('first section')
    const relinked = apply(edited, withHfLink(edited, 1, 'header', 'default', true))
    expect(resolveHf(relinked, 1, 'header', 'default')).toMatchObject({
      value: { text: 'first section' },
      owner: 0,
    })
    expect(relinked.edits).toEqual({})
    expect(hfSaveOptions(relinked)).toEqual({
      sectionHfUnlink: [{ lastBlockIndex: 20, kind: 'header', variant: 'default' }],
    })
  })

  it('a commit whose owner moved backwards under the open editor is refused', () => {
    const s = threeSections()
    const before = resolveHf(s, 2, 'header', 'default').owner
    const linked = apply(s, withHfLink(s, 2, 'header', 'default', true))
    expect(withHfEdit(linked, 2, 'header', 'default', hf('leftover'), before)).toBeNull()
    // unlinking moves the owner onto the section itself: that write goes through
    const inherited = resolveHf(s, 1, 'header', 'default').owner
    const unlinked = apply(s, withHfLink(s, 1, 'header', 'default', false))
    expect(hfCommitTarget(unlinked, 1, 'header', 'default', inherited)).toEqual({
      key: '20:header',
      owner: 1,
    })
  })

  it('survives a break inserted before the section (keyed by the break paragraph)', () => {
    const s = threeSections()
    s.links[linkKey(30, 'header', 'default')] = true
    const live: HfSectionState = {
      ...s,
      sections: [
        s.sections[0],
        { ...s.sections[1], lastBlockIndex: 15, pendingBreak: true },
        s.sections[1],
        s.sections[2],
      ],
    }
    expect(hfLinked(live, 3, 'header', 'default')).toBe(true)
    expect(resolveHf(live, 3, 'header', 'default')).toMatchObject({
      value: { text: 'first section' },
      owner: 0,
    })
  })
})

describe('writers compose from the live maps, not a snapshot', () => {
  it('two commits computed from one snapshot both land when merged functionally', () => {
    const s = threeSections()
    // both targets are decided on the same snapshot (one React render)…
    const a = hfCommitTarget(s, 0, 'header', 'default')!
    const b = hfCommitTarget(s, 2, 'footer', 'default')!
    // …and applied as functional updates on whatever the map holds by then
    let edits = s.edits
    for (const [t, v] of [
      [a, hf('header')],
      [b, hf('footer')],
    ] as const) {
      edits = ((m) => ({ ...m, [t.key]: v }))(edits)
    }
    expect(Object.keys(edits).sort()).toEqual(['10:header', '30:footer'])
    // the link writer sees the merged edits map too
    const linked = withHfLink({ ...s, edits }, 2, 'footer', 'default', true)!
    expect(linked.edits).toEqual({ '10:header': hf('header') })
    expect(linked.links).toEqual({ '30:footer': true })
  })

  it('a first-page edit made on the canvas is what an AI read of that variant returns', () => {
    const s: HfSectionState = {
      ...threeSections(),
      sections: [section(5)],
      docStrips: {
        header: hf('doc header'),
        footer: null,
        variants: { ...EMPTY_HF_VARIANTS, headerFirst: hf('parsed first page') },
      },
    }
    expect(resolveHf(s, 0, 'header', 'first').value?.text).toBe('parsed first page')
    const t = apply(s, withHfEdit(s, 0, 'header', 'first', hf('typed on the canvas')))
    expect(resolveHf(t, 0, 'header', 'first').value?.text).toBe('typed on the canvas')
    expect(resolveHf(t, 0, 'header', 'default').value?.text).toBe('doc header')
    expect(hfSaveOptions(t)).toEqual({ headerFirst: hf('typed on the canvas') })
  })
})

describe('hfSaveOptions', () => {
  it('routes the final section to the trailing options and the rest to sectionHf', () => {
    const s = threeSections()
    s.edits[editKey(30, 'header', 'default')] = hf('last')
    s.edits[editKey(30, 'footer', 'even')] = hf('last even footer')
    s.edits[editKey(10, 'footer', 'first')] = hf('first section title footer')
    s.links[linkKey(20, 'header', 'default')] = true
    expect(hfSaveOptions(s)).toEqual({
      header: hf('last'),
      footerEven: hf('last even footer'),
      sectionHf: [
        {
          lastBlockIndex: 10,
          kind: 'footer',
          variant: 'first',
          hf: hf('first section title footer'),
        },
      ],
      sectionHfUnlink: [{ lastBlockIndex: 20, kind: 'header', variant: 'default' }],
    })
  })

  it('keys round-trip with and without a variant', () => {
    expect(parseEditKey(editKey(20, 'footer', 'default'))).toEqual({
      sectionEnd: 20,
      kind: 'footer',
      variant: 'default',
    })
    expect(parseEditKey(editKey(20, 'header', 'even'))).toEqual({
      sectionEnd: 20,
      kind: 'header',
      variant: 'even',
    })
  })
})
