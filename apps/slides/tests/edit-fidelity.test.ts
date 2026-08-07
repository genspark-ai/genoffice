/**
 * Text-edit save fidelity regression: the full chain of layout -> editor DOM -> extraction ->
 * main-process mapping -> XML patch. Watches three classes of historical defects: run
 * fragmentation (CJK per-character/Latin per-word commits), run field whitelist truncation
 * (letterSpacing/hyperlink/field etc. lost after one edit), and paragraph attributes shifting
 * by index (wrong inheritance when splitting paragraphs).
 */
import { describe, it, expect } from 'vitest'
import { layoutText, makeViewport, HeuristicMetrics } from '@genoffice/pptx-render'
import {
  generateParagraphXml,
  patchTextElementXml,
  type Paragraph,
  type TextElement,
} from '@genoffice/pptx-engine'
import {
  populateEditorDom,
  extractParagraphs,
  releaseEditorLayoutConstraints,
  releaseFragmentsAtEdit,
  applySelectionParagraphFormat,
} from '../src/renderer/TextEditOverlay'
import { applyEditParagraphs, collectParagraphFormatPatches } from '../src/main/edit-text'

const vp = makeViewport({ cx: 12192000, cy: 6858000 }, 1280) // scale = 1
const metrics = new HeuristicMetrics()
const NO_INSETS = { l: 0, t: 0, r: 0, b: 0 }

function layout(paragraphs: Paragraph[], boxWidthPx = 800) {
  return layoutText({
    body: { paragraphs, insets: NO_INSETS },
    boxWidthPx,
    boxHeightPx: 400,
    metrics,
    vp,
  })
}

/** Layout -> editor DOM -> extraction (simulates entering edit mode and committing immediately with no changes). */
function roundTrip(paragraphs: Paragraph[], boxWidthPx = 800) {
  const div = document.createElement('div')
  document.body.appendChild(div)
  populateEditorDom(div, layout(paragraphs, boxWidthPx).lines)
  const out = extractParagraphs(div, 1)
  div.remove()
  return out
}

describe('run fragmentation: layout fragments merge back into original runs by srcRun', () => {
  it('shows PPT letter spacing in the editor DOM to match canvas', () => {
    const div = document.createElement('div')
    document.body.appendChild(div)
    const lines = layout([{ runs: [{ text: '新能源汽车市场100', letterSpacing: -1.5 }] }]).lines
    populateEditorDom(div, lines)
    expect(div.querySelector<HTMLElement>('[data-layout-fragment]')?.style.letterSpacing).toBe(
      `${lines[0]!.runs[0]!.letterSpacingPx}px`,
    )
    div.remove()
  })

  it('preserves multilingual text order across layout fragments', () => {
    const text = '中文 English 日本語 한글 123'
    expect(
      roundTrip([{ runs: [{ text }] }])[0]!
        .runs.map((r) => r.text)
        .join(''),
    ).toBe(text)
  })

  it('restores logical text order from mixed RTL/LTR canvas runs', () => {
    const text = 'English العربية 123 עברית'
    expect(
      roundTrip([{ runs: [{ text }] }])[0]!
        .runs.map((r) => r.text)
        .join(''),
    ).toBe(text)
  })

  it('releases stale fragment widths after input so text can reflow', () => {
    const div = document.createElement('div')
    populateEditorDom(div, layout([{ runs: [{ text: '可编辑文字' }] }]).lines)
    const fragments = [...div.querySelectorAll<HTMLElement>('[data-layout-fragment]')]
    expect(fragments.some((fragment) => fragment.style.width !== '')).toBe(true)
    releaseEditorLayoutConstraints(div)
    expect(
      fragments.every((fragment) => fragment.style.width === '' && fragment.style.display === ''),
    ).toBe(true)
  })

  it('a collapsed-caret edit releases only the touched fragment and its neighbors', () => {
    const div = document.createElement('div')
    populateEditorDom(div, layout([{ runs: [{ text: '可编辑文字排版' }] }]).lines)
    const fragments = [...div.querySelectorAll<HTMLElement>('[data-layout-fragment]')]
    expect(fragments.length).toBeGreaterThan(4)
    const caret = document.createRange()
    caret.setStart(fragments[2]!.firstChild!, 1)
    caret.collapse(true)
    releaseFragmentsAtEdit(div, [caret])
    const released = fragments.map((f) => f.style.width === '')
    expect(released.slice(1, 4)).toEqual([true, true, true])
    // the rest keep their canvas-measured advances, so the surrounding text doesn't jump
    expect(released[0]).toBe(false)
    expect(released.slice(4).every((r) => !r)).toBe(true)
  })

  it('a ranged edit releases every fragment the range intersects', () => {
    const div = document.createElement('div')
    populateEditorDom(div, layout([{ runs: [{ text: '可编辑文字排版' }] }]).lines)
    const fragments = [...div.querySelectorAll<HTMLElement>('[data-layout-fragment]')]
    const range = document.createRange()
    range.setStart(fragments[1]!.firstChild!, 0)
    range.setEnd(fragments[3]!.firstChild!, 1)
    releaseFragmentsAtEdit(div, [range])
    const released = fragments.map((f) => f.style.width === '')
    expect(released.slice(1, 4)).toEqual([true, true, true])
    expect(released[0]).toBe(false)
    expect(released.slice(4).every((r) => !r)).toBe(true)
  })

  it('CJK paragraph (laid out per character) is still a single run after commit', () => {
    const paras: Paragraph[] = [{ runs: [{ text: '产品路线图规划', fontSize: 18 }] }]
    const out = roundTrip(paras)
    expect(out).toHaveLength(1)
    expect(out[0]!.runs).toHaveLength(1)
    expect(out[0]!.runs[0]!.text).toBe('产品路线图规划')
    expect(out[0]!.runs[0]!.srcRun).toBe(0)
  })

  it('multi-run paragraph keeps original run boundaries and order (with format differences)', () => {
    const paras: Paragraph[] = [
      {
        runs: [
          { text: '你好世界', bold: true, fontSize: 18, color: '#112233' },
          { text: ' plain text', fontSize: 18, color: '#445566' },
        ],
      },
    ]
    const out = roundTrip(paras)
    expect(out[0]!.runs.map((r) => r.text)).toEqual(['你好世界', ' plain text'])
    expect(out[0]!.runs.map((r) => r.srcRun)).toEqual([0, 1])
    expect(out[0]!.runs[0]!.bold).toBe(true)
    expect(out[0]!.runs[1]!.bold).toBe(false)
  })

  it('spaces swallowed by auto-wrap are restored exactly via trailingSpace (lossless text)', () => {
    const paras: Paragraph[] = [{ runs: [{ text: 'Hello World Again', fontSize: 18 }] }]
    const out = roundTrip(paras, 90) // narrow box forces wrapping
    expect(out[0]!.runs).toHaveLength(1)
    expect(out[0]!.runs[0]!.text).toBe('Hello World Again')
  })

  it('hard-wrapped long URL no longer gets spaces injected (historical bug: URL rewritten after one edit round-trip)', () => {
    const url = 'https://example.com/aaaa/bbbb'
    const paras: Paragraph[] = [{ runs: [{ text: url, fontSize: 18 }] }]
    const out = roundTrip(paras, 60)
    expect(out[0]!.runs[0]!.text).toBe(url)
  })

  it('multiple paragraphs → each div carries srcPara; empty paragraphs keep their identity', () => {
    const paras: Paragraph[] = [
      { runs: [{ text: '第一段', fontSize: 18 }] },
      { runs: [{ text: '', fontSize: 18 }] },
      { runs: [{ text: '第三段', fontSize: 18 }] },
    ]
    const out = roundTrip(paras)
    expect(out.map((p) => p.srcPara)).toEqual([0, 1, 2])
  })
})

describe('applyEditParagraphs: srcPara/srcRun tracing + unedited fields preserved', () => {
  it('splitting a paragraph: both halves inherit source paragraph props (bullet/line spacing), later paragraphs stay aligned', () => {
    const oldParas: Paragraph[] = [
      { runs: [{ text: 'AB' }], bullet: { type: 'char', char: '•' }, lineHeight: 90 },
      { runs: [{ text: 'C' }], bullet: { type: 'number' }, level: 1 },
    ]
    const out = applyEditParagraphs(oldParas, [
      { runs: [{ text: 'A', srcRun: 0 }], srcPara: 0 },
      { runs: [{ text: 'B', srcRun: 0 }], srcPara: 0 },
      { runs: [{ text: 'C', srcRun: 0 }], srcPara: 1 },
    ])
    expect(out[0]!.bullet).toEqual({ type: 'char', char: '•' })
    expect(out[1]!.bullet).toEqual({ type: 'char', char: '•' }) // the split-off second half shares the source
    expect(out[1]!.lineHeight).toBe(90)
    expect(out[2]!.bullet).toEqual({ type: 'number' })
    expect(out[2]!.level).toBe(1)
  })

  it('run fields the editor cannot express (letter spacing/hyperlink/field/strikethrough) survive editing', () => {
    const oldParas: Paragraph[] = [
      {
        runs: [
          { text: 'link', hyperlink: 'https://a.com', letterSpacing: -2.25, strike: true },
          { text: '3', field: 'slidenum' },
        ],
      },
    ]
    const out = applyEditParagraphs(oldParas, [
      {
        runs: [
          { text: 'edited link', srcRun: 0, bold: false, italic: false, underline: false },
          { text: '3', srcRun: 1, bold: false, italic: false, underline: false },
        ],
        srcPara: 0,
      },
    ])
    const r0 = out[0]!.runs[0]!
    expect(r0.text).toBe('edited link')
    expect(r0.hyperlink).toBe('https://a.com')
    expect(r0.letterSpacing).toBe(-2.25)
    expect(r0.strike).toBe(true)
    expect(out[0]!.runs[1]!.field).toBe('slidenum')
  })

  it('explicit bold:false is not overridden by the old value via ?? fallback (unbolding must take effect)', () => {
    const oldParas: Paragraph[] = [{ runs: [{ text: 'x', bold: true }] }]
    const out = applyEditParagraphs(oldParas, [
      {
        runs: [{ text: 'x', srcRun: 0, bold: false, italic: false, underline: false }],
        srcPara: 0,
      },
    ])
    expect(out[0]!.runs[0]!.bold).toBe(false)
  })

  it('no srcPara/srcRun (newly typed paragraph) falls back to position without crashing', () => {
    const oldParas: Paragraph[] = [{ runs: [{ text: 'a', fontSize: 20 }] }]
    const out = applyEditParagraphs(oldParas, [
      { runs: [{ text: 'a', srcRun: 0 }], srcPara: 0 },
      { runs: [{ text: 'new paragraph' }] },
    ])
    expect(out[1]!.runs[0]!.text).toBe('new paragraph')
  })
})

describe('full chain: one edit → lossless in-place patch path (run count stable)', () => {
  const ORIGINAL_XML =
    '<p:sp><p:nvSpPr><p:cNvPr id="5" name="T"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>' +
    '<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="3048000" cy="1000000"/></a:xfrm></p:spPr>' +
    '<p:txBody><a:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0"/><a:lstStyle/>' +
    '<a:p><a:pPr algn="l"><a:buNone/></a:pPr>' +
    '<a:r><a:rPr lang="en-US" sz="1800" b="1" spc="-100"><a:solidFill><a:srgbClr val="112233"/></a:solidFill></a:rPr><a:t>你好世界</a:t></a:r>' +
    '<a:r><a:rPr sz="1800"><a:solidFill><a:srgbClr val="445566"/></a:solidFill></a:rPr><a:t> plain text</a:t></a:r>' +
    '</a:p></p:txBody></p:sp>'

  const modelParas: Paragraph[] = [
    {
      runs: [
        { text: '你好世界', bold: true, fontSize: 18, letterSpacing: -1, color: '#112233' },
        { text: ' plain text', fontSize: 18, color: '#445566' },
      ],
      align: 'left',
      bullet: { type: 'none' },
    },
  ]

  function editAndPatch(mutate?: (div: HTMLElement) => void): string {
    const div = document.createElement('div')
    document.body.appendChild(div)
    populateEditorDom(div, layout(modelParas).lines)
    mutate?.(div)
    const edited = extractParagraphs(div, 1)
    div.remove()
    const newParas = applyEditParagraphs(modelParas, edited)
    const el = {
      id: 'e1',
      type: 'text',
      text: { paragraphs: newParas },
    } as unknown as TextElement
    return patchTextElementXml(el, ORIGINAL_XML)
  }

  it('no-op commit: takes patch path, spc/pPr(buNone)/bodyPr/text all preserved', () => {
    const out = editAndPatch()
    expect(out).toContain('spc="-100"') // letter spacing not truncated
    expect(out).toContain('<a:buNone/>') // pPr original bytes kept (the patch path doesn't rebuild paragraphs)
    expect(out).toContain('lIns="0"') // bodyPr untouched
    expect(out).toContain('<a:t>你好世界</a:t>')
    expect(out).toContain('<a:t> plain text</a:t>')
    expect((out.match(/<a:r>/g) ?? []).length + (out.match(/<a:r\s/g) ?? []).length).toBe(2)
  })

  it('text-only edit: patch path replaces text, all other bytes preserved', () => {
    const out = editAndPatch((div) => {
      const span = div.querySelector('span[data-src-run="0"]')!
      span.textContent = '早上好世界'
    })
    expect(out).toContain('<a:t>早上好世界</a:t>')
    expect(out).toContain('spc="-100"')
    expect(out).toContain('<a:buNone/>')
    expect(out).toContain('val="445566"') // unedited run's color as-is
  })
})

describe('theme colors: schemeClr is not needlessly baked into srgbClr by edits', () => {
  const THEME_XML =
    '<p:sp><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/>' +
    '<a:p><a:r><a:rPr sz="1800"><a:solidFill><a:schemeClr val="accent1"><a:lumMod val="75000"/></a:schemeClr></a:solidFill></a:rPr><a:t>主题色文字</a:t></a:r></a:p>' +
    '</p:txBody></p:sp>'
  // Model shape after parse: color is the resolved display value + the colorFollowsTheme flag
  const themeParas: Paragraph[] = [
    { runs: [{ text: '主题色文字', fontSize: 18, color: '#C00000', colorFollowsTheme: true }] },
  ]

  function editThemeAndPatch(mutate?: (div: HTMLElement) => void): string {
    const div = document.createElement('div')
    document.body.appendChild(div)
    populateEditorDom(div, layout(themeParas).lines)
    mutate?.(div)
    const edited = extractParagraphs(div, 1)
    div.remove()
    const el = {
      id: 'e1',
      type: 'text',
      text: { paragraphs: applyEditParagraphs(themeParas, edited) },
    } as unknown as TextElement
    return patchTextElementXml(el, THEME_XML)
  }

  it('text-only edit: schemeClr (with lumMod) bytes preserved, still follows theme switches', () => {
    const out = editThemeAndPatch((div) => {
      div.querySelector('span[data-src-run="0"]')!.textContent = '改过的文字'
    })
    expect(out).toContain('<a:schemeClr val="accent1"><a:lumMod val="75000"/></a:schemeClr>')
    expect(out).not.toContain('srgbClr')
    expect(out).toContain('<a:t>改过的文字</a:t>')
  })

  it('user actually changed the color: baked into srgbClr (correct behavior)', () => {
    const out = editThemeAndPatch((div) => {
      const span = div.querySelector('span[data-src-run="0"]') as HTMLElement
      span.style.color = 'rgb(255, 0, 0)'
    })
    expect(out).toContain('<a:srgbClr val="FF0000"/>')
    expect(out).not.toContain('schemeClr')
  })

  it('applyEditParagraphs: unchanged color keeps colorFollowsTheme, changed color clears it', () => {
    const oldParas: Paragraph[] = [
      { runs: [{ text: 'x', color: '#C00000', colorFollowsTheme: true }] },
    ]
    const same = applyEditParagraphs(oldParas, [
      { runs: [{ text: 'y', srcRun: 0, color: '#C00000' }], srcPara: 0 },
    ])
    expect(same[0]!.runs[0]!.colorFollowsTheme).toBe(true)
    const changed = applyEditParagraphs(oldParas, [
      { runs: [{ text: 'y', srcRun: 0, color: '#00FF00' }], srcPara: 0 },
    ])
    expect(changed[0]!.runs[0]!.colorFollowsTheme).toBeUndefined()
    expect(changed[0]!.runs[0]!.color).toBe('#00FF00')
  })
})

describe('normAutofit fontScale: files already shrunk by PowerPoint render at the stored scale', () => {
  it('stored fontScale is used as-is (no scaling back up even if content fits)', () => {
    const body = {
      paragraphs: [{ runs: [{ text: 'Hello', fontSize: 20 }] }] as Paragraph[],
      insets: NO_INSETS,
      autofit: 'shrink' as const,
      fontScale: 0.625,
    }
    const out = layoutText({ body, boxWidthPx: 800, boxHeightPx: 400, metrics, vp })
    expect(out.fontScale).toBeCloseTo(0.625, 3)
    // 20pt × 96/72 × 0.625 ≈ 16.67px
    expect(out.lines[0]!.runs[0]!.fontSizePx).toBeCloseTo(20 * (96 / 72) * 0.625, 1)
  })

  it('content still overflows at the stored scale: keep bisecting downward, never exceed the stored scale', () => {
    const body = {
      paragraphs: Array.from({ length: 20 }, () => ({
        runs: [{ text: '很长的一行文字内容', fontSize: 20 }],
      })) as Paragraph[],
      insets: NO_INSETS,
      autofit: 'shrink' as const,
      fontScale: 0.8,
    }
    const out = layoutText({ body, boxWidthPx: 400, boxHeightPx: 120, metrics, vp })
    expect(out.fontScale).toBeLessThan(0.8)
  })

  it('lnSpcReduction compresses line spacing (fixed lineExact unaffected)', () => {
    const mk = (red?: number) =>
      layoutText({
        body: {
          paragraphs: [{ runs: [{ text: '甲\n乙\n丙', fontSize: 20 }] }] as Paragraph[],
          insets: NO_INSETS,
          autofit: 'shrink' as const,
          fontScale: 1,
          ...(red != null ? { lnSpcReduction: red } : {}),
        },
        boxWidthPx: 800,
        boxHeightPx: 400,
        metrics,
        vp,
      })
    const plain = mk()
    const reduced = mk(0.2)
    expect(reduced.contentHeight).toBeCloseTo(plain.contentHeight * 0.8, 1)
  })
})

describe('edit-mode vertical metrics: line/paragraph spacing derived back from layout lines onto paragraph divs', () => {
  it('line-height = layout line height (px), paragraph margin-top = space before/after', () => {
    const paras: Paragraph[] = [
      { runs: [{ text: 'first', fontSize: 20 }], spaceAfter: 6 },
      { runs: [{ text: 'second', fontSize: 20 }], lineHeight: 150, spaceBefore: 12 },
    ]
    const out = layout(paras)
    const div = document.createElement('div')
    document.body.appendChild(div)
    populateEditorDom(div, out.lines)
    const [p1, p2] = [...div.children] as HTMLElement[]
    expect(parseFloat(p1!.style.lineHeight)).toBeCloseTo(out.lines[0]!.height, 1)
    expect(parseFloat(p2!.style.lineHeight)).toBeCloseTo(out.lines[1]!.height, 1)
    expect(parseFloat(p2!.style.lineHeight)).toBeGreaterThan(parseFloat(p1!.style.lineHeight) * 1.4)
    // Second paragraph's margin-top = spcAft(6pt) + spcBef(12pt) = 18pt -> 24px
    expect(parseFloat(p2!.style.marginTop)).toBeCloseTo(
      out.lines[1]!.top - out.lines[0]!.top - out.lines[0]!.height,
      1,
    )
    expect(parseFloat(p2!.style.marginTop)).toBeCloseTo(24, 0)
    expect(p1!.style.marginTop).toBe('')
    div.remove()
  })

  it('first paragraph spcBef also applies (measured from root padding)', () => {
    const paras: Paragraph[] = [{ runs: [{ text: 'x', fontSize: 20 }], spaceBefore: 9 }]
    const out = layout(paras)
    const div = document.createElement('div')
    document.body.appendChild(div)
    populateEditorDom(div, out.lines)
    expect(parseFloat((div.firstElementChild as HTMLElement).style.marginTop)).toBeCloseTo(12, 1)
    expect(out.lines[0]!.top).toBeCloseTo(12, 1)
    div.remove()
  })

  it('dy baked into line top by middle anchor is excluded, not miscounted as first-paragraph spacing', () => {
    const body = {
      paragraphs: [{ runs: [{ text: 'x', fontSize: 20 }] }] as Paragraph[],
      insets: NO_INSETS,
      anchor: 'middle' as const,
    }
    const out = layoutText({ body, boxWidthPx: 800, boxHeightPx: 400, metrics, vp })
    const dy = (400 - out.contentHeight) / 2
    expect(out.lines[0]!.top).toBeCloseTo(dy, 1) // precondition: the engine does bake dy into top
    const div = document.createElement('div')
    document.body.appendChild(div)
    populateEditorDom(div, out.lines, dy)
    expect((div.firstElementChild as HTMLElement).style.marginTop).toBe('')
    div.remove()
  })

  it('lnSpcReduction compression is reflected in line-height', () => {
    const mk = (red?: number) => {
      const body = {
        paragraphs: [{ runs: [{ text: '甲乙', fontSize: 20 }] }] as Paragraph[],
        insets: NO_INSETS,
        ...(red != null ? { autofit: 'shrink' as const, fontScale: 1, lnSpcReduction: red } : {}),
      }
      return layoutText({ body, boxWidthPx: 800, boxHeightPx: 400, metrics, vp })
    }
    const div = document.createElement('div')
    document.body.appendChild(div)
    populateEditorDom(div, mk(0.2).lines)
    const reduced = parseFloat((div.firstElementChild as HTMLElement).style.lineHeight)
    populateEditorDom(div, mk().lines)
    const plain = parseFloat((div.firstElementChild as HTMLElement).style.lineHeight)
    expect(reduced).toBeCloseTo(plain * 0.8, 1)
    div.remove()
  })
})

describe('soft line break <a:br/>: visible in edit mode + preserved on commit + written back on save', () => {
  // The parse layer rewrites <a:br/> as a sentinel run with text="\n"; the model looks like this
  const brParas: Paragraph[] = [
    {
      runs: [
        { text: '第一行', fontSize: 18, color: '#112233' },
        { text: '\n', fontSize: 18 },
        { text: '第二行', fontSize: 18, color: '#112233' },
      ],
    },
  ]

  it('layout → editor DOM → extraction: sentinel run round-trips (shown as a line break in edit mode)', () => {
    const out = roundTrip(brParas)
    expect(out).toHaveLength(1) // soft line breaks are no longer treated as paragraph splits
    expect(out[0]!.runs.map((r) => r.text)).toEqual(['第一行', '\n', '第二行'])
    expect(out[0]!.runs.map((r) => r.srcRun)).toEqual([0, 1, 2])
  })

  it('no-op commit takes the patch path: <a:br/> bytes preserved', () => {
    const XML =
      '<p:sp><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/>' +
      '<a:p><a:r><a:rPr sz="1800" spc="-50"/><a:t>第一行</a:t></a:r>' +
      '<a:br/>' +
      '<a:r><a:rPr sz="1800"/><a:t>第二行</a:t></a:r></a:p></p:txBody></p:sp>'
    const div = document.createElement('div')
    document.body.appendChild(div)
    populateEditorDom(div, layout(brParas).lines)
    const edited = extractParagraphs(div, 1)
    div.remove()
    const el = {
      id: 'e1',
      type: 'text',
      text: { paragraphs: applyEditParagraphs(brParas, edited) },
    } as unknown as TextElement
    const out = patchTextElementXml(el, XML)
    expect(out).toContain('<a:br/>')
    expect(out).toContain('spc="-50"') // the patch path was hit (a rebuild would lose spc)
    expect(out).toContain('<a:t>第一行</a:t>')
    expect(out).toContain('<a:t>第二行</a:t>')
  })

  it('user-typed \\n (Shift+Enter) is split into sentinel runs on commit', () => {
    const paras: Paragraph[] = [{ runs: [{ text: '甲乙丙', fontSize: 18 }] }]
    const div = document.createElement('div')
    document.body.appendChild(div)
    populateEditorDom(div, layout(paras).lines)
    div.querySelector('span[data-src-run="0"]')!.textContent = '甲\n乙丙'
    const edited = extractParagraphs(div, 1)
    div.remove()
    expect(edited[0]!.runs.map((r) => r.text)).toEqual(['甲', '\n', '乙丙'])
  })

  it('rebuild path writes sentinel/embedded \\n back as <a:br/>', () => {
    expect(generateParagraphXml(brParas[0]!)).toContain('<a:br/>')
    expect(generateParagraphXml({ runs: [{ text: 'a\nb' }] })).toBe(
      '<a:p><a:r><a:rPr/><a:t>a</a:t></a:r><a:br/><a:r><a:rPr/><a:t>b</a:t></a:r></a:p>',
    )
  })
})

describe('applyEditParagraphs: font/underline fidelity markers', () => {
  const oldParas: Paragraph[] = [
    {
      runs: [
        {
          text: '混排',
          fontFamily: '微软雅黑',
          latinFont: 'Calibri',
          eaFont: '微软雅黑',
          underline: true,
          underlineStyle: 'dbl',
        },
      ],
    },
  ]

  it('font unchanged (display-name alias passed back): latin/ea markers preserved', () => {
    const out = applyEditParagraphs(oldParas, [
      { runs: [{ text: '改了字', srcRun: 0, fontFamily: 'Microsoft YaHei' }], srcPara: 0 },
    ])
    expect(out[0]!.runs[0]!.latinFont).toBe('Calibri')
    expect(out[0]!.runs[0]!.eaFont).toBe('微软雅黑')
  })

  it('font actually changed: markers cleared (new font written back)', () => {
    const out = applyEditParagraphs(oldParas, [
      { runs: [{ text: 'x', srcRun: 0, fontFamily: 'Arial' }], srcPara: 0 },
    ])
    expect(out[0]!.runs[0]!.latinFont).toBeUndefined()
    expect(out[0]!.runs[0]!.eaFont).toBeUndefined()
    expect(out[0]!.runs[0]!.fontFamily).toBe('Arial')
  })

  it('display stack head != model name and not in alias table (JP/KR/TC fonts): data-font returns the original name, markers preserved', () => {
    // The display stack head for the JP Hiragino Kaku Gothic name below is 'Hiragino Sans'; FONT_ALIAS has no such mapping
    const jpParas: Paragraph[] = [
      {
        runs: [
          {
            text: '日本語テキスト',
            fontSize: 18,
            fontFamily: 'ヒラギノ角ゴシック',
            latinFont: 'Calibri',
            eaFont: 'ヒラギノ角ゴシック',
          },
        ],
      },
    ]
    const out = roundTrip(jpParas)
    expect(out[0]!.runs[0]!.fontFamily).toBe('ヒラギノ角ゴシック')
    const applied = applyEditParagraphs(jpParas, out)
    expect(applied[0]!.runs[0]!.latinFont).toBe('Calibri')
    expect(applied[0]!.runs[0]!.eaFont).toBe('ヒラギノ角ゴシック')
  })

  it('font actually changed while editing: extraction returns the new font name (data-font does not hijack)', () => {
    const jpParas: Paragraph[] = [
      { runs: [{ text: 'テキスト', fontSize: 18, fontFamily: 'ヒラギノ角ゴシック' }] },
    ]
    const div = document.createElement('div')
    document.body.appendChild(div)
    populateEditorDom(div, layout(jpParas).lines)
    const span = div.querySelector('span[data-src-run="0"]') as HTMLElement
    span.style.fontFamily = 'Arial, sans-serif' // simulate applySelectionFontFamily output
    const out = extractParagraphs(div, 1)
    div.remove()
    expect(out[0]!.runs[0]!.fontFamily).toBe('Arial')
  })

  it('run without explicit font + bold in edit mode does not commit the layout Arial fallback', () => {
    // No fontFamily on the model run: layout falls back to DEFAULT_FONT 'Arial' for display only
    const paras: Paragraph[] = [{ runs: [{ text: '等线加粗测试 Bold' }] }]
    const div = document.createElement('div')
    document.body.appendChild(div)
    populateEditorDom(div, layout(paras).lines)
    // Simulate Chromium execCommand('bold') on select-all: contents wrapped in <b>
    const container = div.querySelector('span[data-src-run="0"]') as HTMLElement
    const b = document.createElement('b')
    while (container.firstChild) b.appendChild(container.firstChild)
    container.appendChild(b)
    const out = extractParagraphs(div, 1)
    div.remove()
    expect(out[0]!.runs[0]!.bold).toBe(true)
    expect(out[0]!.runs[0]!.fontFamily).toBeUndefined()
    const applied = applyEditParagraphs(paras, out)
    expect(applied[0]!.runs[0]!.fontFamily).toBeUndefined()
    expect(generateParagraphXml(applied[0]!)).not.toContain('Arial')
  })

  it('missing-font substitution (等线 displayed as Arial) still commits the model name', () => {
    const paras: Paragraph[] = [
      { runs: [{ text: '等线文字', fontFamily: '等线', latinFont: '等线', eaFont: '等线' }] },
    ]
    // Simulate a platform without DengXian: metrics substitutes every run's display font with Arial
    const subst = {
      metrics: (s: Parameters<typeof metrics.metrics>[0]) => metrics.metrics(s),
      measure: (t: string, s: Parameters<typeof metrics.metrics>[0]) => metrics.measure(t, s),
      displayFamily: () => 'Arial',
    }
    const laidOut = layoutText({
      body: { paragraphs: paras, insets: NO_INSETS },
      boxWidthPx: 800,
      boxHeightPx: 400,
      metrics: subst,
      vp,
    })
    expect(laidOut.lines[0]!.runs[0]!.fontFamily).toBe('Arial')
    expect(laidOut.lines[0]!.runs[0]!.srcFontFamily).toBe('等线')
    const div = document.createElement('div')
    document.body.appendChild(div)
    populateEditorDom(div, laidOut.lines)
    const container = div.querySelector('span[data-src-run="0"]') as HTMLElement
    const b = document.createElement('b')
    while (container.firstChild) b.appendChild(container.firstChild)
    container.appendChild(b)
    const out = extractParagraphs(div, 1)
    div.remove()
    expect(out[0]!.runs[0]!.bold).toBe(true)
    expect(out[0]!.runs[0]!.fontFamily).toBe('等线')
    // Unchanged font: latin/ea preservation markers survive, nothing rewritten as Arial
    const applied = applyEditParagraphs(paras, out)
    expect(applied[0]!.runs[0]!.latinFont).toBe('等线')
    expect(applied[0]!.runs[0]!.eaFont).toBe('等线')
  })

  it('underline keeps dbl; style cleared after removal', () => {
    const keep = applyEditParagraphs(oldParas, [
      { runs: [{ text: 'x', srcRun: 0, underline: true }], srcPara: 0 },
    ])
    expect(keep[0]!.runs[0]!.underlineStyle).toBe('dbl')
    const off = applyEditParagraphs(oldParas, [
      { runs: [{ text: 'x', srcRun: 0, underline: false }], srcPara: 0 },
    ])
    expect(off[0]!.runs[0]!.underline).toBe(false)
    expect(off[0]!.runs[0]!.underlineStyle).toBeUndefined()
  })
})

describe('vertical text (bodyPr vert): horizontal edit overlay round-trips losslessly', () => {
  const layoutVert = (paragraphs: Paragraph[], boxHeightPx: number) =>
    layoutText({
      body: { paragraphs, insets: NO_INSETS, vert: 'eaVert' },
      boxWidthPx: 800,
      boxHeightPx,
      metrics,
      vp,
    })

  it('eaVert per-character column layout → DOM → extraction: runs merge back intact, paragraph structure preserved', () => {
    const paras: Paragraph[] = [
      { runs: [{ text: '縦書きテキスト', fontSize: 18 }] },
      { runs: [{ text: '第二段 with latin', fontSize: 18 }] },
    ]
    const div = document.createElement('div')
    document.body.appendChild(div)
    // Short box forces over-height column wrap: continued column has paraStart=false; splitting still follows model paragraphs
    populateEditorDom(div, layoutVert(paras, 100).lines)
    const out = extractParagraphs(div, 1)
    div.remove()
    expect(out).toHaveLength(2)
    expect(out[0]!.runs).toHaveLength(1)
    expect(out[0]!.runs[0]!.text).toBe('縦書きテキスト')
    expect(out[0]!.runs[0]!.srcRun).toBe(0)
    expect(out[1]!.runs.map((r) => r.text).join('')).toBe('第二段 with latin')
  })

  it('soft break <a:br/> still round-trips as a \\n sentinel in vertical column layout', () => {
    const paras: Paragraph[] = [
      {
        runs: [
          { text: '上段', fontSize: 18 },
          { text: '\n', fontSize: 18 },
          { text: '下段', fontSize: 18 },
        ],
      },
    ]
    const div = document.createElement('div')
    document.body.appendChild(div)
    populateEditorDom(div, layoutVert(paras, 400).lines)
    const out = extractParagraphs(div, 1)
    div.remove()
    expect(out).toHaveLength(1)
    expect(out[0]!.runs.map((r) => r.text)).toEqual(['上段', '\n', '下段'])
  })
})

describe('per-paragraph format on the editing selection', () => {
  const HEADING_AND_BODY: Paragraph[] = [
    { runs: [{ text: '智能教务', fontSize: 16, bold: true }], pPrExplicit: {} },
    {
      runs: [{ text: '校历管理、招生分班', fontSize: 12 }],
      bullet: { type: 'char', char: '•' },
      marL: 228600,
      indent: -228600,
      pPrExplicit: { bullet: true, marL: true, indent: true },
    },
  ]

  function populate(paragraphs: Paragraph[]) {
    const div = document.createElement('div')
    document.body.appendChild(div)
    populateEditorDom(div, layout(paragraphs).lines)
    return div
  }

  it('populate marks the original bullet kind per paragraph (toggle-off reference)', () => {
    const div = populate(HEADING_AND_BODY)
    const blocks = div.querySelectorAll<HTMLElement>('[data-src-para]')
    expect(blocks[0]!.dataset.hadBullet).toBeUndefined()
    expect(blocks[1]!.dataset.hadBullet).toBe('char')
    div.remove()
  })

  it('applySelectionParagraphFormat marks only the selected paragraph; extraction carries it', () => {
    const div = populate(HEADING_AND_BODY)
    div.contentEditable = 'true'
    div.focus()
    const blocks = div.querySelectorAll<HTMLElement>('[data-src-para]')
    const sel = window.getSelection()!
    const range = document.createRange()
    range.selectNodeContents(blocks[1]!)
    sel.removeAllRanges()
    sel.addRange(range)
    expect(applySelectionParagraphFormat({ bullet: 'char', bulletChar: '○' })).toBe(true)

    const out = extractParagraphs(div, 1)
    expect(out[0]!.bullet).toBeUndefined()
    expect(out[1]!.bullet).toBe('char')
    expect(out[1]!.bulletChar).toBe('○')
    div.remove()
  })

  it('toggle-off: applying the current kind again turns the bullet off', () => {
    const div = populate(HEADING_AND_BODY)
    div.contentEditable = 'true'
    div.focus()
    const blocks = div.querySelectorAll<HTMLElement>('[data-src-para]')
    const sel = window.getSelection()!
    const range = document.createRange()
    range.selectNodeContents(blocks[1]!)
    sel.removeAllRanges()
    sel.addRange(range)
    expect(applySelectionParagraphFormat({ bullet: 'char' })).toBe(true)
    expect(blocks[1]!.dataset.bullet).toBe('none')
    div.remove()
  })

  it('unchanged round-trip carries no per-paragraph format (no spurious commits)', () => {
    const out = roundTrip(HEADING_AND_BODY)
    expect(out.every((p) => p.bullet === undefined && p.lineSpacingPct === undefined)).toBe(true)
  })

  it('collectParagraphFormatPatches maps marks to engine patches by paragraph index', () => {
    const out = collectParagraphFormatPatches([
      { runs: [{ text: 'a' }] },
      { runs: [{ text: 'b' }], bullet: 'char', bulletChar: '○', lineSpacingPct: 150 },
      { runs: [{ text: 'c' }], bullet: 'none', spaceAfterPt: 6 },
    ])
    expect(out).toEqual([
      { index: 1, patch: { bullet: 'char', bulletChar: '○', lineSpacingPct: 150 } },
      { index: 2, patch: { bullet: 'none', spaceAfterPt: 6 } },
    ])
  })
})

describe('run hyperlinks: overlay round-trip + merge semantics', () => {
  it('linked run round-trips through the editor DOM as <a href>', () => {
    const out = roundTrip([
      {
        runs: [
          { text: 'see ' },
          { text: 'docs', hyperlink: 'https://d.dev/x', underline: true, underlineImplicit: true },
        ],
      },
    ])
    const runs = out[0]!.runs
    expect(runs.find((r) => r.text.includes('docs'))?.link).toEqual({
      kind: 'url',
      url: 'https://d.dev/x',
    })
    expect(runs.find((r) => r.text.includes('see'))?.link).toBeNull()
  })

  it('slide-jump encoding round-trips as a slide target', () => {
    const out = roundTrip([{ runs: [{ text: '附录', hyperlink: 'slide:41' }] }])
    expect(out[0]!.runs[0]!.link).toEqual({ kind: 'slide', slideIndex: 41 })
  })

  it('adding a link clears the stale rId so ensureRunLinkRels reallocates', () => {
    const oldParas: Paragraph[] = [{ runs: [{ text: 'plain' }] }]
    const out = applyEditParagraphs(oldParas, [
      {
        runs: [{ text: 'plain', srcRun: 0, link: { kind: 'url', url: 'https://n.dev' } }],
        srcPara: 0,
      },
    ])
    const r = out[0]!.runs[0]!
    expect(r.hyperlink).toBe('https://n.dev')
    expect(r.hyperlinkRId).toBeUndefined()
  })

  it('removing a link drops rId/action and the hlink-derived underline', () => {
    const oldParas: Paragraph[] = [
      {
        runs: [
          {
            text: 'linked',
            hyperlink: 'slide:3',
            hyperlinkRId: 'rId8',
            hyperlinkAction: 'ppaction://hlinksldjump',
            underline: true,
            underlineImplicit: true,
          },
        ],
      },
    ]
    const out = applyEditParagraphs(oldParas, [
      { runs: [{ text: 'linked', srcRun: 0, link: null, underline: true }], srcPara: 0 },
    ])
    const r = out[0]!.runs[0]!
    expect(r.hyperlink).toBeUndefined()
    expect(r.hyperlinkRId).toBeUndefined()
    expect(r.hyperlinkAction).toBeUndefined()
    expect(r.underline).toBe(false)
    expect(r.underlineImplicit).toBeUndefined()
  })

  it('retargeting a link keeps it and reallocates the relationship', () => {
    const oldParas: Paragraph[] = [
      { runs: [{ text: 'go', hyperlink: 'https://old.dev', hyperlinkRId: 'rId4' }] },
    ]
    const out = applyEditParagraphs(oldParas, [
      { runs: [{ text: 'go', srcRun: 0, link: { kind: 'slide', slideIndex: 2 } }], srcPara: 0 },
    ])
    const r = out[0]!.runs[0]!
    expect(r.hyperlink).toBe('slide:2')
    expect(r.hyperlinkRId).toBeUndefined()
  })

  it('an unresolvable rId (not representable in the DOM) survives a commit reporting no link', () => {
    const oldParas: Paragraph[] = [{ runs: [{ text: 'x', hyperlinkRId: 'rId9' }] }]
    const out = applyEditParagraphs(oldParas, [
      { runs: [{ text: 'x edited', srcRun: 0, link: null }], srcPara: 0 },
    ])
    expect(out[0]!.runs[0]!.hyperlinkRId).toBe('rId9')
  })
})
