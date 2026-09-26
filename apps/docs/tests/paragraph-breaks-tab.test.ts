/**
 * Paragraph dialog "Line and Page Breaks" tab: keepNext / keepLines /
 * widowControl / suppressLineNumbers round-trip as tri-state pPr flags and the
 * live pagination reads the edited value off the element, not the parsed block.
 */
import { afterEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { generateParagraphXml, mergePPrFormat, parseDocx } from '@genoffice/docx-engine'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { editorExtensions } from '../src/renderer/editor/extensions'
import {
  blocksToPmDoc,
  pmNodeToGeneratedBlock,
  signatureOfBlock,
  signatureOfGenerated,
} from '../src/renderer/editor/convert'
import { clearParagraphFormatting, setParaAttrs } from '../src/renderer/components/ribbon-tabs'
import { runUiOps } from '../src/renderer/ai/ops'
import {
  directParaFlags,
  effectiveParaFlags,
  paraPaginationMeta,
} from '../src/renderer/editor/para-flags'
import { applyBlockMeta } from '../src/renderer/pagination-measure'
import type { BlockBox } from '../src/renderer/pagination-types'

;(globalThis as { CSS?: unknown }).CSS ??= { escape: (s: string) => s }

const editors = new Set<Editor>()
afterEach(() => {
  for (const editor of editors) editor.destroy()
  editors.clear()
})

const p = (pPr: string, text: string) =>
  `<w:p><w:pPr>${pPr}</w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`

async function openDoc(bodyXml: string): Promise<Editor> {
  const parsed = await parseDocx(await buildDocx({ bodyXml }))
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
  editors.add(editor)
  return editor
}

const CTX = { headingStyleIds: new Map<number, string>(), allocateHyperlinkRel: () => 'rId1' }

function firstNode(editor: Editor) {
  return editor.state.doc.firstChild!
}

describe('pagination flags: parse -> node attrs', () => {
  it('direct pPr flags land on the paragraph node, absent ones stay null', async () => {
    const editor = await openDoc(
      p('<w:keepNext/><w:widowControl w:val="0"/><w:suppressLineNumbers/>', 'a') +
        p('<w:pStyle w:val="Normal"/>', 'b'),
    )
    const [a, b] = [editor.state.doc.child(0), editor.state.doc.child(1)]
    expect(a.attrs).toMatchObject({
      keepNext: true,
      keepLines: null,
      widowControl: false,
      suppressLineNumbers: true,
    })
    expect(b.attrs).toMatchObject({
      keepNext: null,
      keepLines: null,
      widowControl: null,
      suppressLineNumbers: null,
    })
  })
})

describe('pagination flags: node attrs -> saved pPr', () => {
  it('a rebuilt paragraph writes each flag, false as an explicit w:val="0"', async () => {
    const editor = await openDoc(p('', 'body'))
    editor.commands.setTextSelection(2)
    setParaAttrs(editor, {
      keepNext: true,
      keepLines: true,
      widowControl: false,
      suppressLineNumbers: true,
    })
    const block = pmNodeToGeneratedBlock(firstNode(editor).toJSON())
    expect(block.format).toMatchObject({
      keepNext: true,
      keepLines: true,
      widowControl: false,
      suppressLineNumbers: true,
    })
    const xml = generateParagraphXml(block, CTX)
    expect(xml).toContain('<w:keepNext/><w:keepLines/>')
    expect(xml).toContain('<w:widowControl w:val="0"/>')
    expect(xml).toContain('<w:suppressLineNumbers/>')
    // schema order: keepNext/keepLines before widowControl before suppressLineNumbers
    expect(xml.indexOf('<w:keepLines/>')).toBeLessThan(xml.indexOf('<w:widowControl'))
    expect(xml.indexOf('<w:widowControl')).toBeLessThan(xml.indexOf('<w:suppressLineNumbers/>'))
  })

  it('an untouched flag is not written at all', async () => {
    const editor = await openDoc(p('', 'body'))
    const block = pmNodeToGeneratedBlock(firstNode(editor).toJSON())
    const xml = generateParagraphXml(block, CTX)
    expect(xml).not.toMatch(/keepNext|keepLines|widowControl|suppressLineNumbers/)
  })
})

describe('mergePPrFormat pagination flags', () => {
  const RAW = '<w:pPr><w:pStyle w:val="Normal"/><w:keepNext/><w:jc w:val="both"/></w:pPr>'

  it('undefined flag keeps the raw bytes', () => {
    expect(mergePPrFormat(RAW, { align: 'justify' })).toBe(RAW)
  })

  it('same value keeps the raw bytes', () => {
    expect(mergePPrFormat(RAW, { align: 'justify', keepNext: true })).toBe(RAW)
  })

  it('turning a raw flag off writes w:val="0" in place', () => {
    expect(mergePPrFormat(RAW, { align: 'justify', keepNext: false })).toBe(
      '<w:pPr><w:pStyle w:val="Normal"/><w:keepNext w:val="0"/><w:jc w:val="both"/></w:pPr>',
    )
  })

  it('a new flag is inserted at its schema position', () => {
    expect(mergePPrFormat(RAW, { align: 'justify', widowControl: false })).toBe(
      '<w:pPr><w:pStyle w:val="Normal"/><w:keepNext/><w:widowControl w:val="0"/><w:jc w:val="both"/></w:pPr>',
    )
    expect(mergePPrFormat(RAW, { align: 'justify', keepLines: true })).toBe(
      '<w:pPr><w:pStyle w:val="Normal"/><w:keepNext/><w:keepLines/><w:jc w:val="both"/></w:pPr>',
    )
  })

  it('raw w:val="0" turned back on drops the attribute', () => {
    const off = '<w:pPr><w:widowControl w:val="0"/></w:pPr>'
    expect(mergePPrFormat(off, { widowControl: true })).toBe('<w:pPr><w:widowControl/></w:pPr>')
    expect(mergePPrFormat(off, { widowControl: false })).toBe(off)
  })
})

describe('outline level (setOutlineLevel op)', () => {
  it('body paragraph -> outline-only heading -> back to body, style kept', async () => {
    const editor = await openDoc(p('<w:pStyle w:val="Normal"/>', 'body'))
    editor.commands.setTextSelection(2)
    runUiOps(editor, [{ op: 'setOutlineLevel', target: { scope: 'selection' }, level: 3 }])
    let node = firstNode(editor)
    expect(node.type.name).toBe('docHeading')
    expect(node.attrs).toMatchObject({ level: 3, outlineOnly: true, styleId: 'Normal' })
    const xml = generateParagraphXml(pmNodeToGeneratedBlock(node.toJSON()), CTX)
    expect(xml).toContain('<w:pStyle w:val="Normal"/>')
    expect(xml).toContain('<w:outlineLvl w:val="2"/>')

    runUiOps(editor, [{ op: 'setOutlineLevel', target: { scope: 'selection' }, level: 0 }])
    node = firstNode(editor)
    expect(node.type.name).toBe('docParagraph')
    expect(node.attrs.styleId).toBe('Normal')
    expect(generateParagraphXml(pmNodeToGeneratedBlock(node.toJSON()), CTX)).not.toContain(
      'outlineLvl',
    )
  })

  it('a styled heading keeps its style-bound level', async () => {
    const editor = await openDoc(p('<w:pStyle w:val="Heading2"/>', 'styled'))
    editor.commands.setTextSelection(2)
    runUiOps(editor, [{ op: 'setOutlineLevel', target: { scope: 'selection' }, level: 5 }])
    expect(firstNode(editor).attrs).toMatchObject({ level: 2, styleId: 'Heading2' })
    expect(firstNode(editor).attrs.outlineOnly).toBeFalsy()
  })
})

describe('live pagination reads the edited flags off the element', () => {
  it('data-para carries the tri-state flags and overrides the parsed meta', async () => {
    const editor = await openDoc(p('<w:keepNext/>', 'a'))
    editor.commands.setTextSelection(2)
    setParaAttrs(editor, { keepNext: false, widowControl: false })
    const el = editor.view.dom.querySelector('p') as HTMLElement
    expect(directParaFlags(el)).toEqual({ keepNext: false, widowControl: false })

    const block = { top: 0, height: 10, el, docxIndex: 0 } as BlockBox
    applyBlockMeta([block], () => ({ keepNext: true, keepLines: true }))
    expect(block.keepNext).toBeUndefined()
    expect(block.keepLines).toBe(true)
    expect(block.widowControl).toBe(false)
  })

  it('a paragraph without docxIndex still gets its direct flags', async () => {
    const editor = await openDoc(p('', 'a'))
    editor.commands.setTextSelection(2)
    setParaAttrs(editor, { keepNext: true })
    const el = editor.view.dom.querySelector('p') as HTMLElement
    const block = { top: 0, height: 10, el } as BlockBox
    applyBlockMeta([block], () => undefined)
    expect(block.keepNext).toBe(true)
  })
})

describe('review follow-ups', () => {
  it('Ctrl+Q clears the direct pagination flags', async () => {
    const editor = await openDoc(p('<w:keepNext/><w:widowControl w:val="0"/>', 'a'))
    editor.commands.setTextSelection(2)
    clearParagraphFormatting(editor)
    expect(firstNode(editor).attrs).toMatchObject({
      keepNext: null,
      keepLines: null,
      widowControl: null,
      suppressLineNumbers: null,
    })
  })

  it('parse-layer meta never carries a direct page break (the element class does)', () => {
    expect(paraPaginationMeta({ pageBreakBefore: true }, undefined)).toBeUndefined()
    expect(paraPaginationMeta(undefined, { pageBreakBefore: true })).toMatchObject({
      breakBefore: true,
    })
    expect(
      paraPaginationMeta({ pageBreakBefore: false }, { pageBreakBefore: true }),
    ).toBeUndefined()
    expect(paraPaginationMeta({ keepNext: true }, { widowControl: false })).toMatchObject({
      keepNext: true,
      widowControl: false,
    })
  })

  it('unchecking a direct page break drops the class the measurement reads', async () => {
    const editor = await openDoc(p('<w:pageBreakBefore/>', 'a'))
    const el = () => editor.view.dom.querySelector('p') as HTMLElement
    expect(el().classList.contains('page-break-before')).toBe(true)
    editor.commands.setTextSelection(2)
    setParaAttrs(editor, { pageBreakBefore: false })
    expect(el().classList.contains('page-break-before')).toBe(false)
  })
})

describe('Ctrl+Q resets a parsed flag to the style for real', () => {
  const RAW = '<w:pPr><w:pStyle w:val="Normal"/><w:keepNext/><w:jc w:val="both"/></w:pPr>'

  it('the saved pPr drops w:keepNext when the model cleared it', () => {
    const original = { keepNext: true, align: 'justify' as const }
    // untouched model: bytes kept
    expect(mergePPrFormat(RAW, { keepNext: true, align: 'justify' }, original)).toBe(RAW)
    // cleared model: the element goes away, the rest keeps its bytes
    expect(mergePPrFormat(RAW, { align: 'justify' }, original)).toBe(
      '<w:pPr><w:pStyle w:val="Normal"/><w:jc w:val="both"/></w:pPr>',
    )
    // without the original, undefined still means "not edited"
    expect(mergePPrFormat(RAW, { align: 'justify' })).toBe(RAW)
  })

  it('end to end: parsed keepNext -> Ctrl+Q -> pPr without keepNext, pagination no longer keeps', async () => {
    const bytes = await buildDocx({ bodyXml: p('<w:keepNext/>', 'a') + p('', 'b') })
    const parsed = await parseDocx(bytes)
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks) as never,
    })
    editors.add(editor)
    const original = parsed.blocks[0]
    expect(original.format?.keepNext).toBe(true)
    const meta = paraPaginationMeta(original.format, undefined)
    expect(meta).toMatchObject({ keepNext: true })

    editor.commands.setTextSelection(2)
    clearParagraphFormatting(editor)
    const node = firstNode(editor)
    expect(node.attrs.keepNext).toBeNull()

    const generated = pmNodeToGeneratedBlock(node.toJSON())
    expect(generated.format?.keepNext).toBeUndefined()
    const merged = mergePPrFormat(original.rawPPr ?? '', generated.format, original.format)
    expect(merged).not.toContain('keepNext')

    const el = editor.view.dom.querySelector('p') as HTMLElement
    const block = { top: 0, height: 10, el, docxIndex: 0 } as BlockBox
    applyBlockMeta([block], () => meta)
    expect(block.keepNext).toBeUndefined()
    // a block with no element (never measured) still gets the parsed constraint
    const bare = { top: 0, height: 10, docxIndex: 0 } as BlockBox
    applyBlockMeta([bare], () => meta)
    expect(bare.keepNext).toBe(true)
  })

  it('style-inherited flags still reach a measured paragraph', () => {
    const meta = paraPaginationMeta(undefined, { keepLines: true, widowControl: false })
    const el = document.createElement('p')
    el.setAttribute('data-para', '{"styleId":"Normal"}')
    expect(effectiveParaFlags(el, meta)).toEqual({ keepLines: true, widowControl: false })
    el.setAttribute('data-para', '{"styleId":"Normal","keepLines":false}')
    expect(effectiveParaFlags(el, meta)).toEqual({ keepLines: false, widowControl: false })
  })
})

describe('contextualSpacing round-trips like the other flags', () => {
  it('rebuilt pPr writes it at its schema position, false as w:val="0"', async () => {
    const editor = await openDoc(p('<w:ind w:left="720"/>', 'body'))
    editor.commands.setTextSelection(2)
    setParaAttrs(editor, { contextualSpacing: true, align: 'center' })
    let xml = generateParagraphXml(pmNodeToGeneratedBlock(firstNode(editor).toJSON()), CTX)
    expect(xml).toContain('<w:ind w:left="720"/><w:contextualSpacing/><w:jc w:val="center"/>')
    setParaAttrs(editor, { contextualSpacing: false })
    xml = generateParagraphXml(pmNodeToGeneratedBlock(firstNode(editor).toJSON()), CTX)
    expect(xml).toContain('<w:contextualSpacing w:val="0"/>')
  })

  it('mergePPrFormat inserts, rewrites and drops it', () => {
    const raw = '<w:pPr><w:pStyle w:val="ListParagraph"/><w:contextualSpacing/></w:pPr>'
    expect(mergePPrFormat(raw, { contextualSpacing: true })).toBe(raw)
    expect(mergePPrFormat(raw, { contextualSpacing: false })).toBe(
      '<w:pPr><w:pStyle w:val="ListParagraph"/><w:contextualSpacing w:val="0"/></w:pPr>',
    )
    expect(mergePPrFormat(raw, {}, { contextualSpacing: true })).toBe(
      '<w:pPr><w:pStyle w:val="ListParagraph"/></w:pPr>',
    )
    expect(
      mergePPrFormat('<w:pPr><w:jc w:val="both"/></w:pPr>', {
        align: 'justify',
        contextualSpacing: true,
      }),
    ).toBe('<w:pPr><w:contextualSpacing/><w:jc w:val="both"/></w:pPr>')
  })

  it('a contextualSpacing-only edit is not a byte-identical pPr reuse', async () => {
    const parsed = await parseDocx(await buildDocx({ bodyXml: p('<w:jc w:val="both"/>', 'a') }))
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks) as never,
    })
    editors.add(editor)
    editor.commands.setTextSelection(2)
    setParaAttrs(editor, { contextualSpacing: true })
    const generated = pmNodeToGeneratedBlock(firstNode(editor).toJSON())
    expect(signatureOfGenerated(generated)).not.toBe(signatureOfBlock(parsed.blocks[0]))
    const merged = mergePPrFormat(
      parsed.blocks[0].rawPPr ?? '',
      generated.format,
      parsed.blocks[0].format,
    )
    expect(merged).toContain('<w:contextualSpacing/>')
  })
})
