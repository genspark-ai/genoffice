import { Editor } from '@tiptap/core'
import { parseDocx, saveDocx } from '@genoffice/docx-engine'
import { describe, expect, it } from 'vitest'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { docStyleCss } from '../src/renderer/doc-style-css'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'

;(globalThis as { CSS?: unknown }).CSS ??= { escape: (s: string) => s }

const STYLES =
  '<w:style w:type="character" w:styleId="Emphasis"><w:name w:val="Emphasis"/>' +
  '<w:rPr><w:i/><w:color w:val="C00000"/></w:rPr></w:style>'

const BODY =
  '<w:p><w:r><w:t xml:space="preserve">before</w:t></w:r>' +
  '<w:r><w:rPr><w:rStyle w:val="Emphasis"/></w:rPr><w:t>emphasized</w:t></w:r>' +
  '<w:r><w:t>after</w:t></w:r></w:p>'

async function open() {
  const source = await buildDocx({ bodyXml: BODY, extraStylesXml: STYLES })
  const parsed = await parseDocx(source)
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
  return { editor, parsed, source }
}

describe('character styles in the editor', () => {
  it('renders rStyle runs as data-style spans', async () => {
    const { editor } = await open()
    const span = editor.view.dom.querySelector('span[data-style="Emphasis"]')
    expect(span?.textContent).toBe('emphasized')
    editor.destroy()
  })

  it('keeps an untouched document with character styles byte-identical', async () => {
    const { editor, parsed, source } = await open()
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    expect(plan.changedCount).toBe(0)
    expect(await saveDocx(parsed, plan.saveBlocks)).toEqual(source)
    editor.destroy()
  })

  it('preserves rStyle when the paragraph is edited', async () => {
    const { editor, parsed } = await open()
    editor.commands.insertContentAt(editor.state.doc.content.size - 1, 'extra')
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    expect(plan.changedCount).toBe(1)
    const bytes = await saveDocx(parsed, plan.saveBlocks)
    const reparsed = await parseDocx(bytes)
    const styled = reparsed.blocks[0].runs!.find((r) => r.styleId === 'Emphasis')
    expect(styled?.text).toBe('emphasized')
    editor.destroy()
  })
})

describe('explicit off (w:val="0") overrides style-inherited formatting', () => {
  const OFF_STYLES =
    '<w:style w:type="paragraph" w:styleId="BoldPara"><w:name w:val="Bold Para"/>' +
    '<w:rPr><w:b/><w:i/><w:smallCaps/></w:rPr></w:style>'

  const OFF_BODY =
    '<w:p><w:pPr><w:pStyle w:val="BoldPara"/></w:pPr>' +
    '<w:r><w:rPr><w:b w:val="0"/><w:i w:val="0"/><w:smallCaps w:val="0"/></w:rPr><w:t>plain</w:t></w:r>' +
    '<w:r><w:t xml:space="preserve"> styled</w:t></w:r></w:p>'

  async function openOff() {
    const source = await buildDocx({ bodyXml: OFF_BODY, extraStylesXml: OFF_STYLES })
    const parsed = await parseDocx(source)
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks) as never,
    })
    return { editor, parsed, source }
  }

  it('run-level off renders counter styles so style CSS cannot re-apply them', async () => {
    const { editor } = await openOff()
    const span = editor.view.dom.querySelector<HTMLElement>('span[data-doc-style]')
    expect(span?.textContent).toBe('plain')
    expect(span!.style.fontWeight).toBe('normal')
    expect(span!.style.fontStyle).toBe('normal')
    expect(span!.style.fontVariantCaps).toBe('normal')
    expect(span!.style.textTransform).toBe('none')
    editor.destroy()
  })

  it('keeps the untouched document byte-identical', async () => {
    const { editor, parsed, source } = await openOff()
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    expect(plan.changedCount).toBe(0)
    expect(await saveDocx(parsed, plan.saveBlocks)).toEqual(source)
    editor.destroy()
  })

  it('style-level off emits reversing declarations in docStyleCss', async () => {
    const source = await buildDocx({
      bodyXml: '<w:p><w:r><w:t>x</w:t></w:r></w:p>',
      extraStylesXml:
        OFF_STYLES +
        '<w:style w:type="paragraph" w:styleId="PlainChild"><w:name w:val="Plain Child"/><w:basedOn w:val="BoldPara"/>' +
        '<w:rPr><w:b w:val="0"/><w:i w:val="0"/><w:smallCaps w:val="0"/></w:rPr></w:style>',
    })
    const css = docStyleCss(await parseDocx(source))
    const rule = css
      .split('\n')
      .find((l) => l.includes('[data-style="PlainChild"]') && l.includes('font-weight'))
    expect(rule).toContain('font-weight:400')
    expect(rule).toContain('font-style:normal')
    expect(rule).toContain('font-variant-caps:normal')
    expect(rule).toContain('text-transform:none')
  })
})
