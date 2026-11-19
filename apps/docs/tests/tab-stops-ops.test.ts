import { afterEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import JSZip from 'jszip'
import { mergeStyleXml, parseDocx, saveDocx } from '@genoffice/docx-engine'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { executeOps } from '../src/renderer/ai/ops'
import { executeTool } from '../src/renderer/ai/tools'
import { parseHtmlFragment } from '../src/renderer/ai/protocol'

const NUM_IDS = { bullet: null, ordered: null }
const editors = new Set<Editor>()
afterEach(() => {
  for (const e of editors) e.destroy()
  editors.clear()
})

const TOA_P =
  '<w:p><w:pPr><w:keepNext/><w:tabs><w:tab w:val="left" w:pos="420"/></w:tabs></w:pPr>' +
  '<w:r><w:t>Smith v. Jones</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>4, 9</w:t></w:r></w:p>'

async function openEditor(bodyXml: string) {
  const parsed = await parseDocx(await buildDocx({ bodyXml }))
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
  })
  editors.add(editor)
  editor.commands.setContent(blocksToPmDoc(parsed.blocks) as never)
  return { editor, parsed }
}

async function saveXml(editor: Editor, parsed: Awaited<ReturnType<typeof parseDocx>>) {
  const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
  const saved = await saveDocx(parsed, plan.saveBlocks)
  const zip = await JSZip.loadAsync(saved)
  return zip.file('word/document.xml')!.async('string')
}

describe('setParagraphFormat tabs', () => {
  it('writes w:tabs with alignment and leader, keeping unmanaged pPr', async () => {
    const { editor, parsed } = await openEditor(TOA_P)
    const out = executeOps(editor, [
      {
        op: 'setParagraphFormat',
        target: { blockIndexes: [0] },
        tabs: [{ pos: '6.5in', align: 'right', leader: 'dot' }],
      },
    ])
    expect(out.ok).toBe(true)
    const xml = await saveXml(editor, parsed)
    expect(xml).toContain('<w:tabs><w:tab w:val="right" w:pos="9360" w:leader="dot"/></w:tabs>')
    expect(xml).not.toContain('w:pos="420"')
    expect(xml).toContain('<w:keepNext/>')
    expect(xml).toContain('<w:tab/>')
  })

  it('tabs: [] removes the paragraph’s own stops', async () => {
    const { editor, parsed } = await openEditor(TOA_P)
    executeOps(editor, [{ op: 'setParagraphFormat', target: { blockIndexes: [0] }, tabs: [] }])
    const xml = await saveXml(editor, parsed)
    expect(xml).not.toContain('<w:tabs>')
    expect(xml).toContain('<w:keepNext/>')
  })

  it('clear: true writes a w:val="clear" stop and keeps display-only inherited stops', () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docParagraph',
            attrs: {
              docxIndex: null,
              tabStops: JSON.stringify([
                { pos: 1000, val: 'center', inherited: true },
                { pos: 4000, val: 'left', inherited: true },
              ]),
            },
            content: [{ type: 'text', text: 'x' }],
          },
        ],
      },
    })
    editors.add(editor)
    const out = executeOps(editor, [
      {
        op: 'setParagraphFormat',
        target: { blockIndexes: [0] },
        tabs: [
          { pos: 4000, clear: true },
          { pos: 2000, align: 'decimal' },
        ],
      },
    ])
    expect(out.ok).toBe(true)
    expect(JSON.parse(editor.state.doc.child(0).attrs.tabStops)).toEqual([
      { pos: 1000, val: 'center', inherited: true },
      { pos: 2000, val: 'decimal' },
      { pos: 4000, val: 'clear' },
    ])
  })

  it('rejects malformed stops', () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: { type: 'doc', content: [{ type: 'docParagraph', attrs: { docxIndex: null } }] },
    })
    editors.add(editor)
    for (const tabs of [
      [{ pos: 'wide' }],
      [{ pos: 100, align: 'middle' }],
      [{ pos: 100, leader: 'stars' }],
      'x',
    ]) {
      const out = executeOps(editor, [
        { op: 'setParagraphFormat', target: { blockIndexes: [0] }, tabs },
      ])
      expect(out.ok).toBe(false)
    }
  })
})

describe('tab character in restricted HTML', () => {
  it('<span data-tab> becomes a tab; a literal tab collapses like any whitespace', async () => {
    const nodes = parseHtmlFragment(
      '<p>Doe v. Roe<span data-tab></span>12</p><p>Lit\t12</p>',
      NUM_IDS,
    )
    expect(nodes[0].content?.map((n) => n.text).join('')).toBe('Doe v. Roe\t12')
    expect(nodes[1].content?.map((n) => n.text).join('')).toBe('Lit 12')
    const selfClosing = parseHtmlFragment('<p>Doe v. Roe<span data-tab />12</p>', NUM_IDS)
    expect(selfClosing[0].content?.map((n) => n.text).join('')).toBe('Doe v. Roe\t12')

    const { editor, parsed } = await openEditor('<w:p><w:r><w:t>Body</w:t></w:r></w:p>')
    const exec = await executeTool(
      editor,
      {
        id: 't',
        name: 'insert_content',
        input: { html: '<p>Doe v. Roe<span data-tab></span>12</p>', afterBlockIndex: 0 },
      },
      NUM_IDS,
    )
    expect(exec.isError).toBeFalsy()
    const xml = await saveXml(editor, parsed)
    expect(xml).toContain('Doe v. Roe</w:t><w:tab/><w:t xml:space="preserve">12</w:t>')
  })
})

describe('define_style tabs', () => {
  it('writes and removes the style’s w:tabs', () => {
    const existing =
      '<w:style w:type="paragraph" w:styleId="TOA"><w:name w:val="TOA"/>' +
      '<w:pPr><w:keepNext/><w:tabs><w:tab w:val="left" w:pos="420"/></w:tabs><w:ind w:left="720"/></w:pPr></w:style>'
    const out = mergeStyleXml(existing, {
      styleId: 'TOA',
      pPr: { tabStops: [{ pos: 9360, val: 'right', leader: 'dot' }] },
    })
    expect(out).toContain(
      '<w:pPr><w:keepNext/><w:tabs><w:tab w:val="right" w:pos="9360" w:leader="dot"/></w:tabs><w:ind w:left="720"/></w:pPr>',
    )
    const cleared = mergeStyleXml(existing, { styleId: 'TOA', pPr: { tabStops: null } })
    expect(cleared).not.toContain('<w:tabs>')
    expect(cleared).toContain('<w:ind w:left="720"/>')
  })
})
