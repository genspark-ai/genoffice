import { afterEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { generateTocFieldXml, parseDocx } from '@genoffice/docx-engine'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { blocksToPmDoc } from '../src/renderer/editor/convert'
import { collectHeadings } from '../src/renderer/editor/headings'

const editors = new Set<Editor>()

afterEach(() => {
  for (const editor of editors) editor.destroy()
  editors.clear()
})

const p = (pPr: string, text: string) =>
  `<w:p><w:pPr>${pPr}</w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`

/** H1/H2/H3 via built-in styles, direct outlineLvl and a basedOn custom style */
async function openMixedHeadingsDoc(): Promise<Editor> {
  const bytes = await buildDocx({
    extraStylesXml:
      '<w:style w:type="paragraph" w:styleId="MySub"><w:name w:val="My Sub"/><w:basedOn w:val="Heading2"/></w:style>',
    bodyXml: [
      p('<w:pStyle w:val="Heading1"/>', 'Chapter 1'),
      p('', 'body text'),
      p('<w:pStyle w:val="Heading2"/>', 'Section 1.1'),
      p('<w:outlineLvl w:val="2"/>', 'Topic 1.1.1'),
      p('<w:pStyle w:val="MySub"/>', 'Section 1.2'),
    ].join(''),
  })
  const parsed = await parseDocx(bytes)
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
  editors.add(editor)
  return editor
}

describe('TOC heading collection', () => {
  it('collects H1-H3 headings with levels in document order', async () => {
    const editor = await openMixedHeadingsDoc()
    const headings = collectHeadings(editor.state.doc)
    expect(headings.map((h) => [h.text, h.level])).toEqual([
      ['Chapter 1', 1],
      ['Section 1.1', 2],
      ['Topic 1.1.1', 3],
      ['Section 1.2', 2],
    ])
  })

  it('generates a level-indented TOC field from the collected headings', async () => {
    const editor = await openMixedHeadingsDoc()
    const entries = collectHeadings(editor.state.doc).map(({ level, text }) => ({ level, text }))
    const frags = generateTocFieldXml(entries)
    expect(frags[0]).toContain(' TOC \\o "1-3" ')
    expect(frags.map((f) => /w:pStyle w:val="(TOC\d)"/.exec(f)?.[1])).toEqual([
      'TOC1',
      'TOC2',
      'TOC3',
      'TOC2',
    ])
  })
})
