/**
 * Run-level hyperlinks: parse-time rId resolution + hlink styling, rels allocation
 * for links set in-session, patch-path <a:hlinkClick> surgery, and live lookup for the show.
 */
import { describe, it, expect } from 'vitest'
import { readFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { dirname, join } from 'node:path'
import { parseSlide } from '../src/parse'
import {
  openPptx,
  savePptx,
  duplicateSlide,
  ensureRunLinkRels,
  getRunLinks,
  patchTextElementXml,
  type OpenedPptx,
  type TextElement,
} from '../src/index'
import type { Theme } from '../src/theme'

const here = dirname(fileURLToPath(import.meta.url))
const fx = (name: string) => readFileSync(join(here, 'fixtures', name))

const slideWith = (sp: string) =>
  '<?xml version="1.0"?><p:sld xmlns:p="p" xmlns:a="a" xmlns:r="r"><p:cSld>' +
  `<p:spTree><p:nvGrpSpPr/><p:grpSpPr/>${sp}</p:spTree></p:cSld></p:sld>`
const spWith = (txBodyInner: string) =>
  '<p:sp><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm></p:spPr>' +
  `<p:txBody>${txBodyInner}</p:txBody></p:sp>`

const parseEl = (txBodyInner: string, hlinkRels?: Map<string, string>, theme?: Theme) => {
  const slide = parseSlide({
    path: 'ppt/slides/slide1.xml',
    slideXml: slideWith(spWith(txBodyInner)),
    ctx: { ...(hlinkRels ? { hlinkRels } : {}), ...(theme ? { theme } : {}) },
  })
  return slide.elements[0] as TextElement
}

/** First element whose text has a non-empty run, plus that run's paragraph/run indexes. */
function findTextRun(opened: OpenedPptx, slideIndex = 0) {
  for (const el of opened.deck.slides[slideIndex]!.elements) {
    if (el.type !== 'text' && el.type !== 'shape') continue
    const text = (el as TextElement).text
    if (!text) continue
    for (let pi = 0; pi < text.paragraphs.length; pi++) {
      const runs = text.paragraphs[pi]!.runs
      for (let ri = 0; ri < runs.length; ri++) {
        if (runs[ri]!.text.trim() && !runs[ri]!.field) {
          return { el: el as TextElement, pi, ri }
        }
      }
    }
  }
  throw new Error('fixture has no text run')
}

describe('parse: run hlinkClick resolved via ctx.hlinkRels', () => {
  const LINKED =
    '<a:bodyPr/><a:p><a:r><a:rPr><a:hlinkClick r:id="rId5"/></a:rPr><a:t>go</a:t></a:r></a:p>'

  it('url target + hlink styling (implicit underline, theme hlink color)', () => {
    const theme = { colors: { hlink: '#0563C1' } } as unknown as Theme
    const el = parseEl(LINKED, new Map([['rId5', 'https://x.dev']]), theme)
    const r = el.text!.paragraphs[0]!.runs[0]!
    expect(r.hyperlink).toBe('https://x.dev')
    expect(r.hyperlinkRId).toBe('rId5')
    expect(r.underline).toBe(true)
    expect(r.underlineImplicit).toBe(true)
    expect(r.color).toBe('#0563C1')
    expect(r.colorFollowsTheme).toBe(true)
    expect(r.colorInherited).toBe(true)
  })

  it('slide jump target keeps the slide:N encoding', () => {
    const el = parseEl(LINKED, new Map([['rId5', 'slide:3']]))
    expect(el.text!.paragraphs[0]!.runs[0]!.hyperlink).toBe('slide:3')
  })

  it('explicit u="none" and explicit fill beat hlink styling', () => {
    const el = parseEl(
      '<a:bodyPr/><a:p><a:r><a:rPr u="none"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>' +
        '<a:hlinkClick r:id="rId5"/></a:rPr><a:t>go</a:t></a:r></a:p>',
      new Map([['rId5', 'https://x.dev']]),
    )
    const r = el.text!.paragraphs[0]!.runs[0]!
    expect(r.underline).toBe(false)
    expect(r.underlineImplicit).toBeUndefined()
    expect(r.color).toBe('#FF0000')
    expect(r.colorFollowsTheme).toBeUndefined()
  })

  it('unresolvable rId keeps hyperlinkRId without hyperlink or styling', () => {
    const el = parseEl(LINKED)
    const r = el.text!.paragraphs[0]!.runs[0]!
    expect(r.hyperlinkRId).toBe('rId5')
    expect(r.hyperlink).toBeUndefined()
    expect(r.underline).toBe(false)
  })
})

describe('patch path: hlinkClick surgery without structural change', () => {
  it('adds hlinkClick for a link set this session', () => {
    const el = parseEl('<a:bodyPr/><a:p><a:r><a:rPr b="1"/><a:t>text</a:t></a:r></a:p>')
    const r = el.text!.paragraphs[0]!.runs[0]!
    r.hyperlink = 'https://x.dev'
    r.hyperlinkRId = 'rId7'
    const out = patchTextElementXml(el, el.anchor.originalXml)
    expect(out).toContain('<a:hlinkClick')
    expect(out).toContain('r:id="rId7"')
  })

  it('removes hlinkClick when the link was cleared, keeping other rPr children', () => {
    const el = parseEl(
      '<a:bodyPr/><a:p><a:r><a:rPr><a:latin typeface="Arial"/>' +
        '<a:hlinkClick r:id="rId5"/></a:rPr><a:t>go</a:t></a:r></a:p>',
      new Map([['rId5', 'https://x.dev']]),
    )
    const r = el.text!.paragraphs[0]!.runs[0]!
    delete r.hyperlink
    delete r.hyperlinkRId
    r.underline = false
    delete r.underlineImplicit
    const out = patchTextElementXml(el, el.anchor.originalXml)
    expect(out).not.toContain('hlinkClick')
    expect(out).toContain('<a:latin typeface="Arial"/>')
  })

  it('implicit (hlink-derived) underline never bakes a u attribute', () => {
    const el = parseEl(LINKED_RID9, new Map([['rId9', 'https://x.dev']]))
    el.text!.paragraphs[0]!.runs[0]!.text = 'changed'
    const out = patchTextElementXml(el, el.anchor.originalXml)
    expect(out).not.toMatch(/\su="/)
    expect(out).toContain('r:id="rId9"') // untouched original hlink bytes
  })
})
const LINKED_RID9 =
  '<a:bodyPr/><a:p><a:r><a:rPr><a:hlinkClick r:id="rId9"/></a:rPr><a:t>go</a:t></a:r></a:p>'

describe('ensureRunLinkRels + save round-trip', () => {
  it('url link set in-session survives save → reopen with resolved target', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const { el, pi, ri } = findTextRun(opened)
    const run = el.text!.paragraphs[pi]!.runs[ri]!
    run.hyperlink = 'https://run.example/x?a=1'
    delete run.hyperlinkRId
    el.dirty = true
    expect(ensureRunLinkRels(opened, 0, el.text!.paragraphs)).toBe(true)
    expect(run.hyperlinkRId).toMatch(/^rId\d+$/)

    const reopened = await openPptx(await savePptx(opened))
    const again = findTextRun(reopened)
    const r2 = again.el.text!.paragraphs[again.pi]!.runs[again.ri]!
    expect(r2.hyperlink).toBe('https://run.example/x?a=1')
    expect(r2.underlineImplicit).toBe(true)
  })

  it('slide jump gets the hlinksldjump action and resolves back to slide:N', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    duplicateSlide(opened, 0)
    const { el, pi, ri } = findTextRun(opened)
    const run = el.text!.paragraphs[pi]!.runs[ri]!
    run.hyperlink = 'slide:1'
    delete run.hyperlinkRId
    el.dirty = true
    ensureRunLinkRels(opened, 0, el.text!.paragraphs)
    expect(run.hyperlinkAction).toBe('ppaction://hlinksldjump')

    const reopened = await openPptx(await savePptx(opened))
    const again = findTextRun(reopened)
    expect(again.el.text!.paragraphs[again.pi]!.runs[again.ri]!.hyperlink).toBe('slide:1')
  })

  it('getRunLinks resolves live from the model rIds', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const { el, pi, ri } = findTextRun(opened)
    const run = el.text!.paragraphs[pi]!.runs[ri]!
    run.hyperlink = 'https://live.example/'
    delete run.hyperlinkRId
    ensureRunLinkRels(opened, 0, el.text!.paragraphs)
    const links = getRunLinks(opened, 0)
    expect(links).toContainEqual({
      elementId: el.id,
      paraIndex: pi,
      runIndex: ri,
      target: { kind: 'url', url: 'https://live.example/' },
    })
  })
})
