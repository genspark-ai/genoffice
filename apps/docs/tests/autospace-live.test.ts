/**
 * CJK-Latin autospace pads: Chromium's text-autospace delivers 1/8em of Word's
 * ~1/4em autoSpaceDE/DN gap; the character after each boundary is wrapped in a
 * .doc-autospace-pad span whose ::before collapsible space supplies the rest, and the pads
 * must follow live edits.
 */
import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import { inlineSegments } from '../src/renderer/text-layout/inline-segments'
import { Editor } from '@tiptap/core'
import { afterAll, beforeAll, describe, expect, it } from 'vitest'
import type { Run } from '@genoffice/docx-engine'
import { editorExtensions, paragraphInlineSegments } from '../src/renderer/editor/extensions'
import {
  renderTableSpec,
  renderTextboxSpec,
  runSpanSpecs,
  runSpansWithPads,
  setStaticAsciiResolver,
} from '../src/renderer/editor/protected-render'
import { DOMSerializer } from '@tiptap/pm/model'
import type { TableModel, TextboxDisplay } from '@genoffice/docx-engine'
import {
  autospacePadScale,
  codePointLengthAt,
  balancedSpaceWideningFor,
} from '../src/renderer/line-metrics'
import { latinAutospaceGapEm } from '../src/renderer/text-layout/font-rules'
import { setOpenDocumentText } from '../src/renderer/text-layout/open-document'

const pads = (editor: Editor): string[] =>
  Array.from(editor.view.dom.querySelectorAll('.doc-autospace-pad'), (el) => el.textContent ?? '')
const padCount = (editor: Editor): number => pads(editor).length

const paragraphDoc = (text: string, attrs?: Record<string, unknown>) =>
  ({
    type: 'doc',
    content: [
      {
        type: 'docParagraph',
        ...(attrs ? { attrs } : {}),
        content: text ? [{ type: 'text', text }] : [],
      },
    ],
  }) as never

const makeEditor = (content: unknown) =>
  new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: content as never,
  })

describe('autospace pad decorations', () => {
  it('pads each CJK-Latin/digit boundary, skipping space/punctuation seams', () => {
    const editor = makeEditor(paragraphDoc('テスト17.0km ペン'))
    expect(padCount(editor)).toBe(1)
    editor.destroy()
  })

  // an empty pad span whose margin overflows the line end makes Chromium rewind
  // to the previous pad instead of breaking inside the CJK run (prod 003)
  it('wraps the character after the boundary instead of inserting an empty span', () => {
    const editor = makeEditor(paragraphDoc('累计产生357次表单提交'))
    expect(pads(editor)).toEqual(['3', '次'])
    expect(editor.view.dom.textContent).toBe('累计产生357次表单提交')
    editor.destroy()
  })

  it('covers hangul boundaries', () => {
    const editor = makeEditor(paragraphDoc('한글A와 B'))
    expect(padCount(editor)).toBe(2)
    editor.destroy()
  })

  it('adds no pads when the paragraph turns autoSpace off', () => {
    const editor = makeEditor(paragraphDoc('ペン12', { autoSpace: false }))
    expect(padCount(editor)).toBe(0)
    editor.destroy()
  })

  it('with only autoSpaceDE off the digit seams keep a full-width pad and the native gap goes', () => {
    const editor = makeEditor(paragraphDoc('BIM\u7b4912\u500b', { autoSpace: 'digits' }))
    expect(pads(editor)).toEqual(['1', '\u500b'])
    const padEls = editor.view.dom.querySelectorAll('.doc-autospace-pad-full')
    expect(padEls).toHaveLength(2)
    expect(editor.view.dom.querySelector('p')!.getAttribute('style')).toMatch(
      /text-autospace:\s*no-autospace/,
    )
    editor.destroy()
  })

  it('pads across styled-run boundaries', () => {
    const editor = makeEditor({
      type: 'doc',
      content: [
        {
          type: 'docParagraph',
          content: [
            { type: 'text', text: 'ペン' },
            {
              type: 'text',
              text: '12',
              marks: [{ type: 'docTextStyle', attrs: { sizeHalfPoints: 24 } }],
            },
          ],
        },
      ],
    })
    expect(padCount(editor)).toBe(1)
    editor.destroy()
  })

  it('follows live edits', () => {
    const editor = makeEditor(paragraphDoc('ペン'))
    expect(padCount(editor)).toBe(0)

    editor.commands.insertContentAt(3, '12')
    expect(padCount(editor)).toBe(1)

    editor.commands.deleteRange({ from: 3, to: 5 })
    expect(padCount(editor)).toBe(0)
    editor.destroy()
  })
})

describe('balanced-space decorations', () => {
  const H = '\ud55c\uae00'
  const spaces = (editor: Editor) =>
    Array.from(editor.view.dom.querySelectorAll('.doc-balanced-space'), (el) => el.textContent)
  beforeAll(() => setOpenDocumentText({ balancedSpaces: true }))
  afterAll(() => setOpenDocumentText({ balancedSpaces: false }))

  it('wraps nothing without balanceSingleByteDoubleByteWidth', () => {
    setOpenDocumentText({ balancedSpaces: false })
    const editor = makeEditor(paragraphDoc(`${H} ${H} A B, ${H}`))
    expect(spaces(editor)).toEqual([])
    editor.destroy()
    setOpenDocumentText({ balancedSpaces: true })
  })

  it('wraps nothing under useFELayout alone or the flag with a Latin document East Asian face', () => {
    for (const [settings, face] of [
      [{ useFELayout: true }, '\u5b8b\u4f53'],
      [{ balanceDbcsSpacing: true }, 'Calibri'],
    ] as const) {
      setOpenDocumentText({ balancedSpaces: balancedSpaceWideningFor(settings, face) })
      const editor = makeEditor(paragraphDoc(`${H} abc ${H}`))
      expect(spaces(editor)).toEqual([])
      editor.destroy()
    }
    for (const [settings, face] of [
      [{ balanceDbcsSpacing: true, useFELayout: true }, 'Calibri'],
      [{ balanceDbcsSpacing: true }, '\u5b8b\u4f53'],
    ] as const) {
      setOpenDocumentText({ balancedSpaces: balancedSpaceWideningFor(settings, face) })
      const editor = makeEditor(paragraphDoc(`${H} abc ${H}`))
      expect(spaces(editor)).toEqual([' ', ' '])
      editor.destroy()
    }
  })

  it('wraps spaces with an East Asian neighbour and runs of spaces, not a lone Latin one', () => {
    const editor = makeEditor(paragraphDoc(`${H} ${H} A B, ${H}`))
    expect(spaces(editor)).toEqual([' ', ' ', ' '])
    expect(editor.view.dom.textContent).toBe(`${H} ${H} A B, ${H}`)
    editor.destroy()
    const han = makeEditor(paragraphDoc('\u7532\uff1a  \u4e59 A B  C'))
    expect(spaces(han)).toEqual([' ', ' ', ' ', ' ', ' '])
    han.destroy()
    const kana = makeEditor(paragraphDoc('\u3042 \u3044'))
    expect(spaces(kana)).toEqual([' '])
    kana.destroy()
  })

  it('adds none in Latin text with single spaces', () => {
    for (const text of ['A B', 'a.b, c;d: e', `${H}, A`]) {
      const editor = makeEditor(paragraphDoc(text))
      expect(spaces(editor)).toEqual([])
      editor.destroy()
    }
  })

  it('reads the neighbour across styled-run boundaries', () => {
    const editor = makeEditor({
      type: 'doc',
      content: [
        {
          type: 'docParagraph',
          content: [
            { type: 'text', text: `${H} ` },
            {
              type: 'text',
              text: 'A',
              marks: [{ type: 'docTextStyle', attrs: { sizeHalfPoints: 24 } }],
            },
          ],
        },
      ],
    } as never)
    expect(spaces(editor)).toEqual([' '])
    editor.destroy()
  })

  it('stays on while the paragraph turns autoSpace off', () => {
    const editor = makeEditor(paragraphDoc(`${H} A`, { autoSpace: false }))
    expect(spaces(editor)).toEqual([' '])
    editor.destroy()
  })
})

describe('static-render autospace pads', () => {
  const run = (text: string): Run => ({ text }) as Run

  it('wraps the character after each boundary', () => {
    expect(runSpansWithPads([run('\u7d2f\u8ba1\u4ea7\u751f357\u6b21')])).toEqual([
      [
        'span',
        {},
        '\u7d2f\u8ba1\u4ea7\u751f',
        ['span', { class: 'doc-autospace-pad' }, '3'],
        '57',
        ['span', { class: 'doc-autospace-pad' }, '\u6b21'],
      ],
    ])
  })

  it('a run continuing a CJK-Latin seam pads its first character', () => {
    expect(runSpansWithPads([run('\u7d2f\u8ba1'), run('286\u540d')])).toEqual([
      ['span', {}, '\u7d2f\u8ba1'],
      [
        'span',
        {},
        ['span', { class: 'doc-autospace-pad' }, '2'],
        '86',
        ['span', { class: 'doc-autospace-pad' }, '\u540d'],
      ],
    ])
    expect(runSpansWithPads([run('\u7d2f\u8ba1'), run('286\u540d')], false)).toEqual([
      ['span', { style: 'text-autospace:no-autospace' }, '\u7d2f\u8ba1'],
      ['span', { style: 'text-autospace:no-autospace' }, '286\u540d'],
    ])
  })

  it('a run rendered on its own (field result, stray line) still gets its pads', () => {
    setOpenDocumentText({ balancedSpaces: true })
    expect(runSpanSpecs(run('\u7d2f\u8ba1357\u6b21 A'))).toEqual([
      [
        'span',
        {},
        '\u7d2f\u8ba1',
        ['span', { class: 'doc-autospace-pad' }, '3'],
        '57',
        ['span', { class: 'doc-autospace-pad' }, '\u6b21'],
        ['span', { class: 'doc-balanced-space' }, ' '],
        'A',
      ],
    ])
    setOpenDocumentText({ balancedSpaces: false })
  })

  it('wraps balanced spaces, reading neighbours from the adjacent runs', () => {
    setOpenDocumentText({ balancedSpaces: true })
    const H = '\ud55c\uae00'
    expect(runSpansWithPads([run(`${H} A B`)])).toEqual([
      ['span', {}, H, ['span', { class: 'doc-balanced-space' }, ' '], 'A B'],
    ])
    expect(runSpansWithPads([run('A '), run(H)])[0]).toEqual([
      'span',
      {},
      'A',
      ['span', { class: 'doc-balanced-space' }, ' '],
    ])
    expect(runSpansWithPads([run(H), run(' A')], false)[1]).toEqual([
      'span',
      { style: 'text-autospace:no-autospace' },
      ['span', { class: 'doc-balanced-space' }, ' '],
      'A',
    ])
    expect(runSpansWithPads([run('A '), run('B')])[0]).toEqual(['span', {}, 'A '])
    setOpenDocumentText({ balancedSpaces: false })
  })

  it('measures astral code points as two UTF-16 units', () => {
    expect(codePointLengthAt('A\u{20000}', 1)).toBe(2)
    expect(codePointLengthAt('AB', 1)).toBe(1)
  })
})

describe('balanced-space advance', () => {
  const read = (rel: string) => readFileSync(new URL(rel, import.meta.url), 'utf8')

  it('draws the wrapped space from the fixed half-width (0.5em) space face', () => {
    const rule = /\.doc-balanced-space\s*\{([^}]*)\}/.exec(read('../src/renderer/styles.css'))
    expect(rule?.[1]).toContain("font-family: 'GenOffice Balanced Space'")
    const face = /font-family: 'GenOffice Balanced Space';([^}]*)\}/.exec(
      read('../src/renderer/fonts/fonts.css'),
    )
    expect(face?.[1]).toContain('GenOfficeBalancedSpace.woff2')
    expect(face?.[1]).toContain('unicode-range: U+0020;')
  })
})

describe('Latin-first pad scale follows the inherited ascii face', () => {
  const KR = '\uc6d0'
  const styles = (ascii: string) =>
    new Map([
      ['Normal', { isDefault: true, type: 'paragraph', display: { fontAscii: ascii } }],
    ]) as never
  const scale = (ascii: string) => String(autospacePadScale(latinAutospaceGapEm(ascii)))

  it('paragraph segments recompute when the style table changes the inherited face', () => {
    const editor = makeEditor(paragraphDoc(`C${KR}`))
    const node = editor.state.doc.firstChild!
    const eaOf = () => null
    const scaleOf = (s: { style?: string }) => Number(/:([\d.]+)$/.exec(s.style ?? '')?.[1])
    const first = paragraphInlineSegments(node, eaOf, () => 'Cambria')
    expect(scaleOf(first.segments[0])).toBeCloseTo(
      autospacePadScale(latinAutospaceGapEm('Cambria')),
      3,
    )
    expect(paragraphInlineSegments(node, eaOf, () => 'Cambria')).toBe(first)
    const second = paragraphInlineSegments(node, eaOf, () => 'Times New Roman')
    expect(second).not.toBe(first)
    expect(scaleOf(second.segments[0])).toBeCloseTo(
      autospacePadScale(latinAutospaceGapEm('Times New Roman')),
      3,
    )
    editor.destroy()
  })

  it('a static run without its own face carries the same pad as the editor', () => {
    const editor = makeEditor(paragraphDoc(`C${KR}`))
    editor.storage.listNumbering.styles = styles('Cambria')
    editor.view.dispatch(editor.state.tr.insertText('x', 3))
    const live = editor.view.dom.querySelector('.doc-autospace-pad-le') as HTMLElement
    expect(live.style.getPropertyValue('--doc-autospace-pad-le').trim()).toBe(scale('Cambria'))
    setStaticAsciiResolver(() => editor.storage.listNumbering)
    try {
      const padStyle = (runs: Run[] | null) => {
        const model: TableModel = {
          rows: [[{ paras: [`C${KR}`], ...(runs ? { richParas: [{ runs }] } : {}) }]],
        }
        const dom = DOMSerializer.renderSpec(document, renderTableSpec(model) as never)
          .dom as HTMLElement
        return (dom.querySelector('.doc-autospace-pad-le') as HTMLElement).style
          .getPropertyValue('--doc-autospace-pad-le')
          .trim()
      }
      expect(padStyle([{ text: `C${KR}` }])).toBe(scale('Cambria'))
      expect(padStyle([{ text: `C${KR}`, fontAscii: 'Arial' }])).toBe(scale('Arial'))
      expect(padStyle(null)).toBe(padStyle([{ text: `C${KR}` }]))
    } finally {
      setStaticAsciiResolver(() => undefined)
    }
    editor.destroy()
  })
})

describe('Korean no-break word-joiners', () => {
  const render = (spec: unknown): HTMLElement =>
    DOMSerializer.renderSpec(document, spec as never).dom as HTMLElement
  const WON_CLOSE_I = '5,000\uc6d0)\uc774 \ub2e4'

  it('the static path marks hangul/punctuation seams without touching the text', () => {
    const model: TableModel = { rows: [[{ paras: [WON_CLOSE_I] }]] }
    const dom = render(renderTableSpec(model))
    const wj = dom.querySelectorAll('.doc-kr-wj')
    expect(wj).toHaveLength(1)
    expect(wj[0].textContent).toBe('')
    expect(wj[0].previousSibling?.textContent?.endsWith(')')).toBe(true)
    expect(wj[0].nextSibling?.textContent?.startsWith('\uc774')).toBe(true)
    expect(dom.querySelector('td')!.textContent).toBe(WON_CLOSE_I)
  })

  it('rich cells carry the seam across run boundaries and respect wordWrap off', () => {
    const runs = [{ text: '5,000\uc6d0)' }, { text: '\uc774 \ub2e4', bold: true }]
    const on: TableModel = {
      rows: [[{ paras: [WON_CLOSE_I], richParas: [{ runs }] }]],
    }
    expect(render(renderTableSpec(on)).querySelectorAll('.doc-kr-wj')).toHaveLength(1)
    const off: TableModel = {
      rows: [[{ paras: [WON_CLOSE_I], richParas: [{ runs, wordWrap: false }] }]],
    }
    expect(render(renderTableSpec(off)).querySelectorAll('.doc-kr-wj')).toHaveLength(0)
  })

  it('header/footer table cells follow their own paragraph props, not the row paragraph', () => {
    const storage = {
      styles: new Map([
        ['Normal', { isDefault: true, type: 'paragraph', display: { fontAscii: 'Cambria' } }],
        ['Cell', { type: 'paragraph', display: { fontAscii: 'Arial' } }],
      ]),
    }
    setStaticAsciiResolver(() => storage as never)
    try {
      const KR = '\ud55c'
      const box: TextboxDisplay = {
        paras: [
          {
            runs: [],
            styleId: 'Normal',
            cells: [
              {
                paras: [[{ text: `C${KR}` }], [{ text: WON_CLOSE_I }]],
                paraProps: [{ styleId: 'Cell' }, { wordWrap: false }],
              },
            ],
          },
        ],
      }
      const dom = render(renderTextboxSpec(box))
      const pad = dom.querySelector('.doc-autospace-pad-le') as HTMLElement
      expect(pad.style.getPropertyValue('--doc-autospace-pad-le').trim()).toBe(
        String(autospacePadScale(latinAutospaceGapEm('Arial'))),
      )
      expect(dom.querySelectorAll('.doc-kr-wj')).toHaveLength(0)
    } finally {
      setStaticAsciiResolver(() => undefined)
    }
  })

  it('the editor joins a seam split over two text nodes', () => {
    const editor = makeEditor({
      type: 'doc',
      content: [
        {
          type: 'docParagraph',
          content: [
            { type: 'text', text: '5,000\uc6d0)' },
            { type: 'text', text: '\uc774 \ub2e4', marks: [{ type: 'bold' }] },
          ],
        },
      ],
    })
    expect(editor.view.dom.querySelectorAll('.doc-kr-wj')).toHaveLength(1)
    editor.destroy()
  })

  it('the pad gap is a collapsible ::before space carrying the whole gap; hangul seams keep the margin (Word drops the gap at a line break; corpus 2026-11-10)', () => {
    const css = readFileSync(join(__dirname, '../src/renderer/styles.css'), 'utf8')
    const rule = (sel: string) => {
      const m = new RegExp(`${sel.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}\\s*\\{([^}]*)\\}`).exec(
        css,
      )
      return m?.[1] ?? ''
    }
    expect(rule('@property --doc-autospace-gap')).toMatch(/syntax:\s*'<length>'/)
    expect(rule('.doc-autospace-pad')).toMatch(
      /--doc-autospace-gap:\s*calc\(var\(--doc-autospace-pad[^)]*\)\s*\*\s*var\(--doc-autospace-pad-le,\s*2\)\)/,
    )
    const before = rule('.doc-autospace-pad::before')
    expect(before).toMatch(/content:\s*' '/)
    expect(before).toMatch(/white-space:\s*normal/)
    expect(before).toMatch(/font-size:\s*0\b/)
    // a length line-height inherited by the empty pseudo grew every padded line (gate 2026-11-10)
    expect(before).toMatch(/line-height:\s*0\b/)
    expect(before).toMatch(/word-spacing:\s*var\(--doc-autospace-gap\)/)
    expect(rule('.doc-autospace-pad-m::before')).toMatch(/content:\s*none/)
    expect(rule('.doc-autospace-pad-m')).toMatch(/margin-inline-start:[\s\S]*- 1\s*\)/)
    expect(css).not.toMatch(/(?<![\w.-])\.doc-autospace-pad(?:-full|-le)?\s*\{[^}]*margin/)
    // the scale is the whole gap: Calibri's Latin -> CJK 0.26 em = 2.08 x the 1/8 em base
    expect(autospacePadScale(0.26)).toBe(2.08)
    expect(autospacePadScale(0.25)).toBe(2)
  })

  it('hangul seams take the margin form only under Korean word-wrap; with wordWrap off they take the space', () => {
    const para = (wordWrap: boolean) =>
      ({
        runs: [{ text: 'PG\uc0ac 1\ud68c\uc6a9', offset: 0, size: 11, attrs: null }],
        inheritedAscii: 'Calibri',
        inheritedEa: 'Batang',
        wordWrap,
        balancedSpaces: false,
        justified: false,
        simsunGapLineHeight: null,
        autospace: true,
      }) as never
    const classesOf = (wordWrap: boolean) =>
      inlineSegments(para(wordWrap))
        .segments.filter((s) => s.cls?.includes('doc-autospace-pad'))
        .map((s) => s.cls!)
    const on = classesOf(true)
    expect(on.length).toBeGreaterThan(0)
    for (const c of on) expect(c).toContain('doc-autospace-pad-m')
    const off = classesOf(false)
    expect(off.length).toBe(on.length)
    for (const c of off) expect(c).not.toContain('doc-autospace-pad-m')
  })
})
