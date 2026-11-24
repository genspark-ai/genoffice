/**
 * Paragraph dialog "Asian Typography" tab: hanging punctuation and the two
 * automatic-spacing switches become the paragraph's own w:overflowPunct /
 * w:autoSpaceDE / w:autoSpaceDN; an untouched switch keeps following the style.
 */
import { afterEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { generateParagraphXml, mergePPrFormat, parseDocx } from '@genoffice/docx-engine'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { blocksToPmDoc, pmNodeToGeneratedBlock } from '../src/renderer/editor/convert'
import {
  activeParaAttrs,
  clearParagraphFormatting,
  setParaAttrs,
} from '../src/renderer/components/ribbon-tabs'
import {
  asianTypographyAttrs,
  asianTypographyFromAttrs,
} from '../src/renderer/components/paragraph-asian-typography'

;(globalThis as { CSS?: unknown }).CSS ??= { escape: (s: string) => s }

const editors = new Set<Editor>()
afterEach(() => {
  for (const editor of editors) editor.destroy()
  editors.clear()
})

const p = (pPr: string, text = 'body') =>
  `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}<w:r><w:t>${text}</w:t></w:r></w:p>`

const TIGHT =
  '<w:style w:type="paragraph" w:styleId="Tight"><w:name w:val="Tight"/><w:basedOn w:val="Normal"/>' +
  '<w:pPr><w:overflowPunct w:val="0"/><w:autoSpaceDE w:val="0"/><w:autoSpaceDN w:val="0"/></w:pPr></w:style>'

async function openDoc(bodyXml: string, extraStylesXml?: string) {
  const parsed = await parseDocx(await buildDocx({ bodyXml, extraStylesXml }))
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
  editors.add(editor)
  editor.commands.setTextSelection(2)
  return { editor, parsed }
}

const CTX = { headingStyleIds: new Map<number, string>(), allocateHyperlinkRel: () => 'rId1' }
const ALL_ON = {
  textAlignment: 'auto' as const,
  kinsoku: true,
  topLinePunct: false,
  latinWrap: false,
  hangingPunct: true,
  autoSpaceLatin: true,
  autoSpaceDigits: true,
}

describe('Asian Typography: what the dialog shows', () => {
  it('Word defaults are all on', async () => {
    const { editor } = await openDoc(p(''))
    expect(asianTypographyFromAttrs(activeParaAttrs(editor))).toEqual(ALL_ON)
  })
  it('the paragraph own values and the read-only ones', async () => {
    const { editor } = await openDoc(
      p(
        '<w:kinsoku w:val="0"/><w:overflowPunct w:val="0"/><w:topLinePunct/>' +
          '<w:autoSpaceDN w:val="0"/><w:textAlignment w:val="center"/>',
      ),
    )
    const attrs = activeParaAttrs(editor)
    expect(asianTypographyFromAttrs(attrs)).toEqual({
      textAlignment: 'center',
      kinsoku: false,
      topLinePunct: true,
      latinWrap: false,
      hangingPunct: false,
      autoSpaceLatin: true,
      autoSpaceDigits: false,
    })
    expect(attrs).toMatchObject({ kinsoku: false, topLinePunct: true, textAlignment: 'center' })
  })
  it('values from the paragraph style show as the effective state', async () => {
    const { editor } = await openDoc(p('<w:pStyle w:val="Tight"/>'), TIGHT)
    const attrs = activeParaAttrs(editor)
    expect(asianTypographyFromAttrs(attrs)).toEqual({
      textAlignment: 'auto',
      kinsoku: true,
      topLinePunct: false,
      latinWrap: false,
      hangingPunct: false,
      autoSpaceLatin: false,
      autoSpaceDigits: false,
    })
    expect(attrs).toMatchObject({ overflowPunctDirect: null, autoSpaceDE: null, autoSpaceDN: null })
  })
})

describe('Asian Typography: what OK writes', () => {
  it('nothing for untouched switches', () => {
    expect(asianTypographyAttrs(ALL_ON, ALL_ON)).toEqual({})
  })
  it('only the changed switch becomes direct', () => {
    expect(asianTypographyAttrs({ ...ALL_ON, autoSpaceDigits: false }, ALL_ON)).toEqual({
      autoSpaceDN: false,
      autoSpace: 'latin',
    })
    expect(asianTypographyAttrs({ ...ALL_ON, hangingPunct: false }, ALL_ON)).toEqual({
      overflowPunct: false,
      overflowPunctDirect: false,
    })
  })
  it('a paragraph without pPr gets the tags in schema order', async () => {
    const { editor } = await openDoc(p(''))
    const initial = asianTypographyFromAttrs(activeParaAttrs(editor))
    setParaAttrs(
      editor,
      asianTypographyAttrs({ ...ALL_ON, hangingPunct: false, autoSpaceLatin: false }, initial),
    )
    const xml = generateParagraphXml(
      pmNodeToGeneratedBlock(editor.state.doc.firstChild!.toJSON()),
      CTX,
    )
    expect(xml).toContain('<w:overflowPunct w:val="0"/><w:autoSpaceDE w:val="0"/>')
    expect(xml).not.toContain('autoSpaceDN')
  })
  it('an existing pPr keeps its other bytes and takes the tags in place', async () => {
    const { editor, parsed } = await openDoc(
      p('<w:keepNext/><w:kinsoku w:val="0"/><w:autoSpaceDN w:val="0"/><w:jc w:val="center"/>'),
    )
    const initial = asianTypographyFromAttrs(activeParaAttrs(editor))
    setParaAttrs(
      editor,
      asianTypographyAttrs({ ...initial, hangingPunct: false, autoSpaceDigits: true }, initial),
    )
    const generated = pmNodeToGeneratedBlock(editor.state.doc.firstChild!.toJSON())
    const original = parsed.blocks[0]
    expect(mergePPrFormat(original.rawPPr!, generated.format, original.format)).toBe(
      '<w:pPr><w:keepNext/><w:kinsoku w:val="0"/><w:overflowPunct w:val="0"/>' +
        '<w:autoSpaceDN/><w:jc w:val="center"/></w:pPr>',
    )
  })
  it('turning a style value back on writes an explicit on, and reads back on', async () => {
    const { editor, parsed } = await openDoc(p('<w:pStyle w:val="Tight"/>'), TIGHT)
    const initial = asianTypographyFromAttrs(activeParaAttrs(editor))
    setParaAttrs(editor, asianTypographyAttrs(ALL_ON, initial))
    const generated = pmNodeToGeneratedBlock(editor.state.doc.firstChild!.toJSON())
    const original = parsed.blocks[0]
    const pPr = mergePPrFormat(original.rawPPr!, generated.format, original.format)
    expect(pPr).toBe(
      '<w:pPr><w:pStyle w:val="Tight"/><w:overflowPunct/><w:autoSpaceDE/><w:autoSpaceDN/></w:pPr>',
    )
    const reopened = await parseDocx(
      await buildDocx({
        bodyXml: `<w:p>${pPr}<w:r><w:t>x</w:t></w:r></w:p>`,
        extraStylesXml: TIGHT,
      }),
    )
    expect(reopened.blocks[0].format).toMatchObject({ overflowPunct: true })
    expect(reopened.blocks[0].format?.autoSpace ?? true).toBe(true)
  })
  it('a paragraph nobody touched keeps its raw pPr', async () => {
    const { editor, parsed } = await openDoc(p('<w:pStyle w:val="Tight"/>'), TIGHT)
    const generated = pmNodeToGeneratedBlock(editor.state.doc.firstChild!.toJSON())
    const original = parsed.blocks[0]
    expect(mergePPrFormat(original.rawPPr!, generated.format, original.format)).toBe(
      original.rawPPr,
    )
  })
  it('Allow Latin text to wrap writes w:wordWrap off; a style value is shown, not copied', async () => {
    const { editor, parsed } = await openDoc(p('<w:keepNext/><w:jc w:val="center"/>'))
    const initial = asianTypographyFromAttrs(activeParaAttrs(editor))
    setParaAttrs(editor, asianTypographyAttrs({ ...initial, latinWrap: true }, initial))
    expect(activeParaAttrs(editor)).toMatchObject({ wordWrap: false, wordWrapDirect: false })
    const generated = pmNodeToGeneratedBlock(editor.state.doc.firstChild!.toJSON())
    const original = parsed.blocks[0]
    expect(mergePPrFormat(original.rawPPr!, generated.format, original.format)).toBe(
      '<w:pPr><w:keepNext/><w:wordWrap w:val="0"/><w:jc w:val="center"/></w:pPr>',
    )
    const style =
      '<w:style w:type="paragraph" w:styleId="Split"><w:name w:val="Split"/>' +
      '<w:basedOn w:val="Normal"/><w:pPr><w:wordWrap w:val="0"/></w:pPr></w:style>'
    const styled = await openDoc(p('<w:pStyle w:val="Split"/>'), style)
    const attrs = activeParaAttrs(styled.editor)
    expect(asianTypographyFromAttrs(attrs).latinWrap).toBe(true)
    expect(attrs.wordWrapDirect).toBeNull()
    const block = pmNodeToGeneratedBlock(styled.editor.state.doc.firstChild!.toJSON())
    const source = styled.parsed.blocks[0]
    expect(mergePPrFormat(source.rawPPr!, block.format, source.format)).toBe(source.rawPPr)
  })
  it('the first/last character rules and line-start compression are written when changed', async () => {
    const { editor, parsed } = await openDoc(p('<w:keepNext/><w:jc w:val="center"/>'))
    const initial = asianTypographyFromAttrs(activeParaAttrs(editor))
    setParaAttrs(
      editor,
      asianTypographyAttrs({ ...initial, kinsoku: false, topLinePunct: true }, initial),
    )
    const generated = pmNodeToGeneratedBlock(editor.state.doc.firstChild!.toJSON())
    const original = parsed.blocks[0]
    const pPr = mergePPrFormat(original.rawPPr!, generated.format, original.format)
    expect(pPr).toBe(
      '<w:pPr><w:keepNext/><w:kinsoku w:val="0"/><w:topLinePunct/><w:jc w:val="center"/></w:pPr>',
    )
    const reopened = await openDoc(`<w:p>${pPr}<w:r><w:t>x</w:t></w:r></w:p>`)
    const back = asianTypographyFromAttrs(activeParaAttrs(reopened.editor))
    setParaAttrs(
      reopened.editor,
      asianTypographyAttrs({ ...back, kinsoku: true, topLinePunct: false }, back),
    )
    const again = pmNodeToGeneratedBlock(reopened.editor.state.doc.firstChild!.toJSON())
    const source = reopened.parsed.blocks[0]
    expect(mergePPrFormat(source.rawPPr!, again.format, source.format)).toBe(
      '<w:pPr><w:keepNext/><w:kinsoku/><w:topLinePunct w:val="0"/><w:jc w:val="center"/></w:pPr>',
    )
  })
  it('w:kinsoku off puts a break opportunity before closers and after openers', async () => {
    const { kinsokuOffPluginKey } = await import('../src/renderer/editor/cjk-punct-shrink')
    const { editor } = await openDoc(
      p('<w:kinsoku w:val="0"/>', '\u4e2d\u3002\u4e2d\uff08\u4e2d (a.b), c') +
        p('', '\u4e2d\u3002'),
    )
    const found = kinsokuOffPluginKey.getState(editor.state)!.find()
    expect(found.map((d) => d.from)).toEqual([2, 5])
  })
  it('the Strict level puts a joiner before small kana in Japanese text only', async () => {
    const { kinsokuOffPluginKey, setKinsokuLevel } =
      await import('../src/renderer/editor/cjk-punct-shrink')
    const ja = '<w:rPr><w:lang w:eastAsia="ja-JP"/></w:rPr>'
    const zh = '<w:rPr><w:lang w:eastAsia="zh-CN"/></w:rPr>'
    const text = '\u304d\u3083\u30fc\u3042'
    const run = (rPr: string) => `<w:p><w:r>${rPr}<w:t>${text}</w:t></w:r></w:p>`
    const { editor } = await openDoc(run(ja) + run(zh))
    const marks = () =>
      kinsokuOffPluginKey
        .getState(editor.state)!
        .find()
        .map((d) => d.from)
    expect(marks()).toEqual([])
    setKinsokuLevel(editor, true, null)
    expect(marks()).toEqual([2, 3])
    setKinsokuLevel(editor, false, null)
    expect(marks()).toEqual([])
  })
  it('the Custom level replaces the standard lists for text of its language only', async () => {
    const { kinsokuOffPluginKey, setKinsokuLevel } =
      await import('../src/renderer/editor/cjk-punct-shrink')
    const ja = '<w:rPr><w:lang w:eastAsia="ja-JP"/></w:rPr>'
    const zh = '<w:rPr><w:lang w:eastAsia="zh-CN"/></w:rPr>'
    // kanji, a listed kana, a stop the list drops, an opener the list drops, a listed kana
    const text = '\u56fd\u3042\u3002\uff08\u3044\u56fd'
    const run = (rPr: string) => `<w:p><w:r>${rPr}<w:t>${text}</w:t></w:r></w:p>`
    const { editor } = await openDoc(run(ja) + run(zh))
    const marks = () =>
      kinsokuOffPluginKey
        .getState(editor.state)!
        .find()
        .map((d) => `${d.from}:${(d.spec as { key: string }).key}`)
    const custom = { lang: 'ja-JP', before: '\u3042', after: '\u3044' }
    setKinsokuLevel(editor, false, custom)
    expect(marks()).toEqual([
      '2:kinsoku-custom',
      '3:kinsoku-free',
      '4:kinsoku-free',
      '5:kinsoku-free',
      '6:kinsoku-custom',
    ])
    // Strict wins over a Japanese list
    setKinsokuLevel(editor, true, custom)
    expect(marks()).toEqual([])
    // a list for a language no East Asian text has changes nothing
    setKinsokuLevel(editor, false, { ...custom, lang: 'English' })
    expect(marks()).toEqual([])
    setKinsokuLevel(editor, false, { ...custom, lang: 'zh-CN' })
    expect(marks().map((m) => Number(m.split(':')[0]))).toEqual([10, 11, 12, 13, 14])
  })
  it('a listed character behind an inline object still gets its joiner', async () => {
    const { kinsokuOffPluginKey, setKinsokuLevel } =
      await import('../src/renderer/editor/cjk-punct-shrink')
    const run = '<w:r><w:t>\u56fd</w:t></w:r><w:r><w:br/></w:r><w:r><w:t>\u3042\u56fd</w:t></w:r>'
    const { editor } = await openDoc(`<w:p>${run}</w:p>`)
    const first = editor.state.doc.firstChild!
    expect(first.childCount).toBe(3)
    expect(first.child(1).isText).toBe(false)
    setKinsokuLevel(editor, false, { lang: 'zh-CN', before: '\u3042', after: '' })
    const found = kinsokuOffPluginKey.getState(editor.state)!.find()
    expect(found.map((d) => `${d.from}:${(d.spec as { key: string }).key}`)).toEqual([
      '3:kinsoku-custom',
    ])
  })
  it('a Custom list frees a Latin mark only where it touches East Asian text', async () => {
    const { kinsokuOffPluginKey, setKinsokuLevel } =
      await import('../src/renderer/editor/cjk-punct-shrink')
    const { editor } = await openDoc(p('', '\u56fd) a.b (\u56fd'))
    setKinsokuLevel(editor, false, { lang: 'zh-CN', before: '', after: '' })
    const found = kinsokuOffPluginKey.getState(editor.state)!.find()
    expect(found.map((d) => d.from)).toEqual([2, 9])
  })
  it('Text Alignment writes w:textAlignment, and Auto takes it away', async () => {
    const { editor, parsed } = await openDoc(p('<w:keepNext/><w:jc w:val="center"/>'))
    const initial = asianTypographyFromAttrs(activeParaAttrs(editor))
    setParaAttrs(editor, asianTypographyAttrs({ ...initial, textAlignment: 'top' }, initial))
    const generated = pmNodeToGeneratedBlock(editor.state.doc.firstChild!.toJSON())
    const original = parsed.blocks[0]
    const pPr = mergePPrFormat(original.rawPPr!, generated.format, original.format)
    expect(pPr).toBe(
      '<w:pPr><w:keepNext/><w:jc w:val="center"/><w:textAlignment w:val="top"/></w:pPr>',
    )
    const reopened = await openDoc(`<w:p>${pPr}<w:r><w:t>x</w:t></w:r></w:p>`)
    const back = asianTypographyFromAttrs(activeParaAttrs(reopened.editor))
    expect(back.textAlignment).toBe('top')
    const untouched = pmNodeToGeneratedBlock(reopened.editor.state.doc.firstChild!.toJSON())
    const source = reopened.parsed.blocks[0]
    expect(mergePPrFormat(source.rawPPr!, untouched.format, source.format)).toBe(source.rawPPr)
    setParaAttrs(reopened.editor, asianTypographyAttrs({ ...back, textAlignment: 'auto' }, back))
    const cleared = pmNodeToGeneratedBlock(reopened.editor.state.doc.firstChild!.toJSON())
    expect(mergePPrFormat(source.rawPPr!, cleared.format, source.format)).toBe(
      '<w:pPr><w:keepNext/><w:jc w:val="center"/></w:pPr>',
    )
  })
  it('the offsets follow the Word probe: SimSun 10 pt beside 28 pt', async () => {
    const { textAlignmentShift } = await import('../src/renderer/editor/text-alignment')
    // hhea ascent 0.859, descent 0.141 of the em; sizes in pt
    const box = (size: number) => ({
      size,
      ascent: size * 0.859375,
      descent: size * 0.140625,
      eastAsian: true,
    })
    // Word: 15.6 higher, 6.72 higher, 1.68 lower (0.24 pt grid)
    expect(textAlignmentShift('top', box(28), box(10))).toBeCloseTo(15.47, 1)
    expect(textAlignmentShift('center', box(28), box(10))).toBeCloseTo(6.47, 1)
    expect(textAlignmentShift('bottom', box(28), box(10))).toBeCloseTo(-1.9, 1)
    const latin = { ...box(10), eastAsian: false }
    expect(textAlignmentShift('bottom', { ...box(28), eastAsian: false }, latin)).toBeCloseTo(
      -1.27,
      1,
    )
  })
  it('Clear Paragraph Formatting drops the direct switches', async () => {
    const { editor } = await openDoc(p('<w:overflowPunct w:val="0"/><w:autoSpaceDN w:val="0"/>'))
    clearParagraphFormatting(editor)
    expect(activeParaAttrs(editor)).toMatchObject({
      overflowPunctDirect: null,
      autoSpaceDN: null,
      overflowPunct: null,
      autoSpace: null,
    })
    const format = pmNodeToGeneratedBlock(editor.state.doc.firstChild!.toJSON()).format
    expect(format?.overflowPunctDirect).toBeUndefined()
    expect(format?.autoSpaceDN).toBeUndefined()
  })
})
