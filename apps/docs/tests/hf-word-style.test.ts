/**
 * A new header or footer starts like Word's: one left-aligned paragraph in the
 * Header / Footer style, with a center and a right tab stop (the document's
 * own style when it has one, else Word's built-in 4680 / 9360 twips), and a
 * line whose right stop lies past the column reaches out to it while editing,
 * as the strip draws it and as Word sets it (probe 2026-11-20).
 */
import { afterEach, describe, expect, it, vi } from 'vitest'
import type { Editor } from '@tiptap/core'
import type { HeaderFooter, ParsedDocFull, StyleInfo } from '@genoffice/docx-engine'
import { docStyleCss } from '../src/renderer/doc-style-css'
import {
  HF_LIVE_COMMIT_MS,
  hfStopReachCss,
  mountHfEditor,
  type HfEditorHandle,
} from '../src/renderer/editor/hf-editor'
import { hfBlankPara, hfIsBlank } from '../src/renderer/editor/hf-text'

const handles: HfEditorHandle[] = []
afterEach(() => {
  for (const h of handles.splice(0)) if (!h.editor.isDestroyed) h.exit()
})

function mount(value: HeaderFooter, blank = hfBlankPara('header')) {
  const host = document.createElement('div')
  document.body.append(host)
  const commits: HeaderFooter[] = []
  const handle = mountHfEditor(host, {
    value,
    pageNo: '1',
    pageTotal: '1',
    spellcheck: false,
    blank,
    onCommit: (next) => commits.push(next),
    onExit: () => host.remove(),
  })
  handles.push(handle)
  return { host, handle, editor: handle.editor as Editor, commits }
}

const style = (over: Partial<StyleInfo>): StyleInfo => ({
  styleId: 'Header',
  name: 'header',
  type: 'paragraph',
  ...over,
})

describe('the paragraph a new header or footer starts with', () => {
  it("is Word's built-in Header / Footer when the document has none", () => {
    for (const kind of ['header', 'footer'] as const) {
      const para = hfBlankPara(kind, new Map())
      expect(para).toMatchObject({
        styleId: kind === 'header' ? 'Header' : 'Footer',
        tabStops: [
          { pos: 4680, val: 'center' },
          { pos: 9360, val: 'right' },
        ],
        spaceAfter: 0,
        lineSpacing: 1,
        runs: [],
      })
      expect(para.align).toBeUndefined()
    }
  })

  it("takes the document's own style, found by id or by its built-in name", () => {
    const tabs = [
      { pos: 4513, val: 'center' as const },
      { pos: 9026, val: 'right' as const },
    ]
    const byId = hfBlankPara(
      'header',
      new Map([['Header', style({ display: { tabStops: tabs, spaceAfterTwips: 0 } })]]),
    )
    expect(byId).toMatchObject({ styleId: 'Header', tabStops: tabs, spaceAfter: 0 })
    // a localized id keeps Word's built-in name
    const byName = hfBlankPara(
      'footer',
      new Map([
        [
          'Fusszeile',
          style({ styleId: 'Fusszeile', name: 'footer', display: { align: 'center' } }),
        ],
      ]),
    )
    expect(byName).toMatchObject({ styleId: 'Fusszeile', align: 'center' })
  })
})

describe('opening a blank part', () => {
  it('starts on the blank paragraph and commits nothing until something is typed', () => {
    expect(hfIsBlank({ text: '' })).toBe(true)
    const { handle, host, commits } = mount({ text: '' })
    const para = host.querySelector<HTMLElement>('.doc-textbox-para')!
    expect(para.dataset.style).toBe('Header')
    expect(para.style.textAlign).toBe('')
    handle.exit()
    expect(commits).toHaveLength(0)
  })

  it("commits the typed line in the blank paragraph's style and stops", () => {
    const { handle, editor, commits } = mount({ text: '' })
    editor.commands.insertContent('Left')
    handle.exit()
    expect(commits).toHaveLength(1)
    expect(commits[0].paras).toHaveLength(1)
    expect(commits[0].paras![0]).toMatchObject({
      styleId: 'Header',
      tabStops: [
        { pos: 4680, val: 'center' },
        { pos: 9360, val: 'right' },
      ],
      runs: [{ text: 'Left' }],
    })
  })

  it('leaves a part that has content on its own paragraphs', () => {
    expect(hfIsBlank({ text: 'Title' })).toBe(false)
    expect(hfIsBlank({ text: '', pageNumber: true })).toBe(false)
    const { host } = mount({ text: '', paras: [{ align: 'right', runs: [{ text: 'Own' }] }] })
    const paras = host.querySelectorAll<HTMLElement>('.doc-textbox-para')
    expect(paras).toHaveLength(1)
    expect(paras[0].dataset.style).toBeUndefined()
  })

  it('lets a line reach a tab stop past the column instead of the column edge', () => {
    const builtIn = JSON.stringify(hfBlankPara('header').tabStops)
    // 9360 twips = 624px: on a narrower column the line's box widens to it
    expect(hfStopReachCss(builtIn, null)).toBe('margin-right:min(0px, calc(100% - 624.0px))')
    // a right indent keeps its own margin; cleared and margin-relative stops reach nowhere
    expect(hfStopReachCss(builtIn, 120)).toBeNull()
    expect(hfStopReachCss(JSON.stringify([{ pos: 9360, val: 'clear' }]), null)).toBeNull()
    expect(
      hfStopReachCss(JSON.stringify([{ pos: 50, val: 'right', rel: 'margin' }]), null),
    ).toBeNull()
    expect(hfStopReachCss(null, null)).toBeNull()
  })
})

describe('typing into an open part', () => {
  afterEach(() => vi.useRealTimers())

  it('commits after a pause, so the page is laid out again while it is open', () => {
    vi.useFakeTimers({ toFake: ['setTimeout', 'clearTimeout'] })
    const { handle, editor, commits } = mount({ text: '' })
    editor.commands.insertContent('Line one')
    vi.advanceTimersByTime(HF_LIVE_COMMIT_MS - 1)
    expect(commits).toHaveLength(0)
    editor.commands.insertContent(' and more')
    vi.advanceTimersByTime(HF_LIVE_COMMIT_MS - 1)
    // each keystroke restarts the pause
    expect(commits).toHaveLength(0)
    vi.advanceTimersByTime(1)
    expect(commits).toHaveLength(1)
    expect(commits[0].paras![0].runs).toEqual([{ text: 'Line one and more' }])
    expect(editor.isDestroyed).toBe(false)
    handle.exit()
    expect(commits.at(-1)!.paras![0].runs).toEqual([{ text: 'Line one and more' }])
  })

  it('drops a pending commit when the part is discarded', () => {
    vi.useFakeTimers({ toFake: ['setTimeout', 'clearTimeout'] })
    const { handle, editor, commits } = mount({ text: '' })
    editor.commands.insertContent('Gone')
    handle.discard()
    vi.advanceTimersByTime(HF_LIVE_COMMIT_MS * 2)
    expect(commits).toHaveLength(0)
  })
})

describe("the first page's header and the last page's footer", () => {
  it("take the document's face, not the app's UI font", () => {
    ;(globalThis as { CSS?: unknown }).CSS ??= { escape: (x: string) => x }
    const parsed = {
      styles: new Map(),
      docDefaults: { asciiFont: 'Aptos', sizeHalfPoints: 24 },
      blocks: [],
    } as unknown as ParsedDocFull
    // they are .doc-page siblings: the body's font-family rule does not reach them
    expect(docStyleCss(parsed)).toMatch(/\.page-wrap > \.page-hf \{ font-family:[^}]*Aptos/)
  })
})
