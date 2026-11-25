/**
 * Forms protection section by section: with one section left open (w:formProt
 * off) the editor takes edits there and refuses those that reach a protected one.
 */
import { afterEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { parseDocx, readSections } from '@genoffice/docx-engine'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { blocksToPmDoc } from '../src/renderer/editor/convert'
import { toggleCheckboxAt } from '../src/renderer/editor/checkbox-toggle'
import { SECTION_GUARD_PASS } from '../src/renderer/editor/section-edit-guard'
import { showSavedDocument } from '../src/renderer/file-actions'

;(globalThis as { CSS?: unknown }).CSS ??= { escape: (s: string) => s }

const editors = new Set<Editor>()
afterEach(() => {
  for (const editor of editors) editor.destroy()
  editors.clear()
})

const P = (text: string, pPr = '') =>
  `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}<w:r><w:t>${text}</w:t></w:r></w:p>`
const SECT = (inner: string) =>
  `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:type w:val="continuous"/>${inner}</w:sectPr>`
const CHECKBOX =
  '<w:p><w:r><w:fldChar w:fldCharType="begin"><w:ffData><w:name w:val="Check1"/><w:enabled/>' +
  '<w:checkBox><w:sizeAuto/><w:default w:val="0"/></w:checkBox></w:ffData></w:fldChar></w:r>' +
  '<w:r><w:instrText xml:space="preserve"> FORMCHECKBOX </w:instrText></w:r>' +
  '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>'

/** section 1 protected (two paragraphs and a check box), section 2 open, section 3 protected */
const BODY =
  P('locked one') +
  CHECKBOX +
  P('locked two', SECT('')) +
  P('open one') +
  P('open two', SECT('<w:formProt w:val="0"/>')) +
  P('tail')

async function openDoc() {
  const parsed = await parseDocx(await buildDocx({ bodyXml: BODY }))
  const sections = readSections(parsed)
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
  editors.add(editor)
  const guard = editor.storage.sectionEditGuard
  guard.sections = sections
  guard.locked = sections.map((s) => s.settings.formProt !== false)
  let blocked = 0
  guard.onBlocked = () => blocked++
  const at = (text: string) => {
    let found = -1
    editor.state.doc.descendants((node, pos) => {
      if (node.isText && node.text?.includes(text)) found = pos + node.text.indexOf(text)
    })
    expect(found).toBeGreaterThan(-1)
    return found
  }
  return { editor, sections, at, blocked: () => blocked }
}

describe('forms protection section by section', () => {
  it('reads which sections stay open', async () => {
    const { sections } = await openDoc()
    expect(sections.map((s) => s.settings.formProt)).toEqual([undefined, false, undefined])
  })

  it('typing, formatting and deleting work in the open section only', async () => {
    const { editor, at, blocked } = await openDoc()
    const text = () => editor.state.doc.textBetween(0, editor.state.doc.content.size, '|')
    editor.commands.insertContentAt(at('open one'), 'X')
    expect(text()).toContain('Xopen one')
    expect(blocked()).toBe(0)

    editor.commands.insertContentAt(at('locked one'), 'X')
    expect(text()).not.toContain('Xlocked')
    expect(blocked()).toBe(1)

    const from = at('locked two')
    editor
      .chain()
      .setTextSelection({ from, to: from + 6 })
      .toggleMark('bold')
      .run()
    expect(editor.state.doc.rangeHasMark(from, from + 6, editor.schema.marks.bold)).toBe(false)
    editor.commands.deleteRange({ from, to: from + 6 })
    expect(text()).toContain('locked two')

    const open = at('open two')
    editor
      .chain()
      .setTextSelection({ from: open, to: open + 4 })
      .toggleMark('bold')
      .run()
    expect(editor.state.doc.rangeHasMark(open, open + 4, editor.schema.marks.bold)).toBe(true)
  })

  it('an edit across the section boundary is refused', async () => {
    const { editor, at } = await openDoc()
    const size = editor.state.doc.content.size
    editor.commands.deleteRange({ from: at('locked two') + 7, to: at('open one') + 4 })
    expect(editor.state.doc.content.size).toBe(size)
    // joining the first open paragraph onto the protected one above it
    const start = at('open one')
    editor.commands.deleteRange({ from: start - 2, to: start })
    expect(editor.state.doc.content.size).toBe(size)
    editor.commands.insertContentAt(at('tail'), 'X')
    expect(editor.state.doc.content.size).toBe(size)
    // a new paragraph after the open section's last one joins that section
    editor.commands.insertContentAt(at('open two') + 8, { type: 'docParagraph' })
    expect(editor.state.doc.content.size).toBe(size + 2)
  })

  it('a form check box in a protected section still toggles', async () => {
    const { editor } = await openDoc()
    let box = -1
    editor.state.doc.descendants((node, pos) => {
      if (node.isText && /[\u2610\u2612]/.test(node.text ?? '')) box = pos
    })
    expect(box).toBeGreaterThan(-1)
    const tr = toggleCheckboxAt(editor.state, box, 'after')!
    editor.view.dispatch(tr)
    expect(editor.state.doc.textBetween(box, box + 1)).toBe('\u2612')
  })

  it('layout bookkeeping and an unguarded document pass', async () => {
    const { editor, at } = await openDoc()
    const pos = at('locked one')
    editor.view.dispatch(editor.state.tr.insertText('Y', pos).setMeta('addToHistory', false))
    expect(editor.state.doc.textBetween(pos, pos + 1)).toBe('Y')
    editor.storage.sectionEditGuard.locked = null
    editor.commands.insertContentAt(pos, 'Z')
    expect(editor.state.doc.textBetween(pos, pos + 1)).toBe('Z')
  })

  it('two saves keep the open section break on its paragraph', async () => {
    const { editor, at } = await openDoc()
    const guard = editor.storage.sectionEditGuard
    const { pmDocToSavePlan } = await import('../src/renderer/editor/convert')
    const { saveDocx } = await import('@genoffice/docx-engine')
    let parsed = await parseDocx(await buildDocx({ bodyXml: BODY }))
    const save = async () => {
      const plan = pmDocToSavePlan(editor.getJSON() as never, parsed.blocks)
      parsed = await parseDocx(await saveDocx(parsed, plan.saveBlocks, {}))
      const sections = readSections(parsed)
      showSavedDocument(editor, blocksToPmDoc(parsed.blocks, sections) as never)
      guard.sections = sections
      return sections
    }
    editor.commands.insertContentAt(at('open one') + 'open one'.length + 1, {
      type: 'docParagraph',
      content: [{ type: 'text', text: 'added' }],
    })
    await save()
    // the reload gave the new paragraph its place in the file
    let anchored = false
    editor.state.doc.forEach((node) => {
      if (node.textContent === 'added') anchored = node.attrs.docxIndex !== null
    })
    expect(anchored).toBe(true)
    editor.commands.insertContentAt(at('open two'), 'X')
    const sections = await save()
    const ends = sections.map((s) =>
      (parsed.blocks[s.lastBlockIndex]?.runs ?? []).map((r) => r.text).join(''),
    )
    expect(ends.slice(0, 2)).toEqual(['locked two', 'Xopen two'])
    expect(sections.map((s) => s.settings.formProt)).toEqual([undefined, false, undefined])
  })

  it('a whole-document replacement passes only when it is not an edit', async () => {
    const { editor, blocked } = await openDoc()
    const other = { type: 'doc', content: [{ type: 'docParagraph' }] }
    editor.commands.setContent(other as never)
    expect(editor.state.doc.childCount).toBeGreaterThan(1)
    expect(blocked()).toBe(1)
    editor
      .chain()
      .setMeta(SECTION_GUARD_PASS, true)
      .setContent(other as never)
      .run()
    expect(editor.state.doc.childCount).toBe(1)
    expect(blocked()).toBe(1)
  })
})
