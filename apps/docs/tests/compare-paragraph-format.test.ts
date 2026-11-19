import { afterEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { compareDocuments, parseDocx, readSections, saveDocx } from '@genoffice/docx-engine'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { editorExtensions } from '../src/renderer/editor/extensions'
import {
  blocksToPmDoc,
  pmDocOptions,
  pmDocToSavePlan,
  type PmNode,
} from '../src/renderer/editor/convert'
import {
  acceptAllRevisions,
  collectRevisions,
  rejectAllRevisions,
} from '../src/renderer/editor/revisions'

;(globalThis as { CSS?: unknown }).CSS ??= { escape: (s: string) => s }

const AUTHOR = 'Rae Visor'
const DATE = '2026-11-18T00:00:00Z'
const para = (text: string, pPr = '') =>
  `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}<w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`

const editors: Editor[] = []
afterEach(() => {
  for (const editor of editors.splice(0)) editor.destroy()
})

/** the comparison of two bodies, opened in the editor as the redline is */
async function redline(originalXml: string, revisedXml: string) {
  const { bytes } = await compareDocuments(
    await parseDocx(await buildDocx({ bodyXml: originalXml })),
    await parseDocx(await buildDocx({ bodyXml: revisedXml })),
    { author: AUTHOR, date: DATE },
  )
  const parsed = await parseDocx(bytes)
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    editorProps: { handleScrollToSelection: () => true },
    content: blocksToPmDoc(parsed.blocks, readSections(parsed), pmDocOptions(parsed)),
  })
  editors.push(editor)
  const saved = async () => {
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    const xml = (await parseDocx(await saveDocx(parsed, plan.saveBlocks))).internal.documentXml
    return [...xml.matchAll(/<w:p>(<w:pPr>[\s\S]*?<\/w:pPr>)?<w:r>/g)].map((m) => m[1] ?? '')
  }
  return { editor, saved }
}

describe('a paragraph formatting difference found by Compare is a tracked change in the editor', () => {
  const ORIGINAL =
    para('This paragraph stays.') +
    para(
      'Aligned and indented.',
      '<w:spacing w:before="120"/><w:ind w:left="360"/><w:jc w:val="left"/>',
    )
  const REVISED =
    para('This paragraph stays.') +
    para(
      'Aligned and indented.',
      '<w:spacing w:before="480"/><w:ind w:left="1440"/><w:jc w:val="center"/>',
    )

  it('it is listed as a paragraph formatting change under the label', async () => {
    const { editor } = await redline(ORIGINAL, REVISED)
    const revisions = collectRevisions(editor.state.doc)
    expect(revisions.map((r) => r.kind)).toEqual(['pPrChange'])
    expect(revisions[0]).toMatchObject({ author: AUTHOR, date: DATE })
    expect(editor.state.doc.child(1).attrs.align).toBe('center')
  })

  it('accepting keeps the revised formatting and saves no record', async () => {
    const { editor, saved } = await redline(ORIGINAL, REVISED)
    acceptAllRevisions(editor)
    expect(collectRevisions(editor.state.doc)).toHaveLength(0)
    expect(await saved()).toEqual([
      '',
      '<w:pPr><w:spacing w:before="480"/><w:ind w:left="1440"/><w:jc w:val="center"/></w:pPr>',
    ])
  })

  it('rejecting restores the original formatting and saves no record', async () => {
    const { editor, saved } = await redline(ORIGINAL, REVISED)
    rejectAllRevisions(editor)
    expect(collectRevisions(editor.state.doc)).toHaveLength(0)
    expect(editor.state.doc.child(1).attrs.align).toBe('left')
    const pPr = (await saved())[1]!
    expect(pPr).not.toContain('w:pPrChange')
    expect(pPr).toContain('<w:spacing w:before="120"/>')
    expect(pPr).toContain('<w:ind w:left="360"/>')
    expect(pPr).toContain('<w:jc w:val="left"/>')
  })

  it('rejecting a style and page-break change restores them', async () => {
    const { editor, saved } = await redline(
      para('This paragraph stays.') + para('A heading once.', '<w:pStyle w:val="Heading1"/>'),
      para('This paragraph stays.') + para('A heading once.', '<w:pageBreakBefore/>'),
    )
    rejectAllRevisions(editor)
    expect((await saved())[1]).toBe('<w:pPr><w:pStyle w:val="Heading1"/></w:pPr>')
  })

  it('rejecting a list change puts the paragraph back in its list', async () => {
    const list = '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr>'
    const { bytes } = await compareDocuments(
      await parseDocx(
        await buildDocx({
          withNumbering: true,
          bodyXml: para('This paragraph stays.') + para('An item.', list),
        }),
      ),
      await parseDocx(
        await buildDocx({
          withNumbering: true,
          bodyXml: para('This paragraph stays.') + para('An item.'),
        }),
      ),
      { author: AUTHOR, date: DATE },
    )
    const parsed = await parseDocx(bytes)
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      editorProps: { handleScrollToSelection: () => true },
      content: blocksToPmDoc(parsed.blocks, readSections(parsed), pmDocOptions(parsed)),
    })
    editors.push(editor)
    expect(editor.state.doc.child(1).type.name).toBe('docParagraph')
    rejectAllRevisions(editor)
    expect(editor.state.doc.child(1).type.name).toBe('docListItem')
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    const xml = (await parseDocx(await saveDocx(parsed, plan.saveBlocks))).internal.documentXml
    expect(xml).toContain(list)
    expect(xml).not.toContain('w:pPrChange')
  })
})
