/**
 * A direct w:outlineLvl on a non-heading style makes the paragraph an outline /
 * TOC entry but Word keeps its body formatting: no built-in heading size, font
 * or italic may apply. Styled headings keep the built-in look.
 */
import { readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { afterEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { parseDocx } from '@genoffice/docx-engine'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { blocksToPmDoc, pmNodeToGeneratedBlock } from '../src/renderer/editor/convert'
import { docThemeCss } from '../src/renderer/doc-style-css'

;(globalThis as { CSS?: unknown }).CSS ??= { escape: (s: string) => s }

const STYLES_CSS = readFileSync(
  join(dirname(fileURLToPath(import.meta.url)), '../src/renderer/styles.css'),
  'utf8',
)

const editors = new Set<Editor>()
afterEach(() => {
  for (const editor of editors) editor.destroy()
  editors.clear()
})

const p = (pPr: string, text: string) =>
  `<w:p><w:pPr>${pPr}</w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`

async function openDoc(): Promise<Editor> {
  const bytes = await buildDocx({
    bodyXml: [
      p('<w:outlineLvl w:val="1"/>', 'outline-only body paragraph'),
      p('<w:pStyle w:val="Heading2"/>', 'styled heading'),
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

describe('outline-only headings', () => {
  it('renders the class on the outline-only h2 and not on the styled one', async () => {
    const editor = await openDoc()
    const [outline, styled] = Array.from(editor.view.dom.querySelectorAll('h2'))
    expect(outline.classList.contains('doc-outline-only')).toBe(true)
    expect(styled.classList.contains('doc-outline-only')).toBe(false)
    expect(styled.getAttribute('data-style')).toBe('Heading2')
  })

  it('saves back as an outline level, never as a Heading style', async () => {
    const editor = await openDoc()
    const node = editor.state.doc.child(0).toJSON()
    const block = pmNodeToGeneratedBlock(node)
    expect(block).toMatchObject({ type: 'heading', level: 2, outlineOnly: true })
    expect(block.styleId).toBeUndefined()
    const styledBlock = pmNodeToGeneratedBlock(editor.state.doc.child(1).toJSON())
    expect(styledBlock.outlineOnly).toBeUndefined()
  })

  it('built-in heading font rules skip the class without gaining specificity', () => {
    for (const n of [1, 2, 3, 4, 5, 6]) {
      expect(STYLES_CSS).toContain(`.doc-page h${n}:where(:not(.doc-outline-only))`)
      expect(STYLES_CSS).not.toMatch(new RegExp(`\\.doc-page h${n} \\{\\s*font-size`))
    }
    expect(STYLES_CSS).toMatch(
      /\.doc-page :where\(h1, h2, h3, h4, h5, h6\)\.doc-outline-only \{\s*font-size: inherit;/,
    )
    const css = docThemeCss({ major: 'Aptos Display', minor: 'Aptos' }, null)
    expect(css).toContain('.doc-page h1:where(:not(.doc-outline-only))')
    expect(css).not.toMatch(/\.doc-page h1[ ,]/)
  })
})
