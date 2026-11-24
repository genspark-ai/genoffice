/**
 * Lock Fields, Unlock Fields, Unlink Fields, Next / Previous Field and Update
 * Field on a table of contents kept as blocks, and on the live page fields of
 * a header or footer.
 */
import { afterEach, beforeEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { NodeSelection } from '@tiptap/pm/state'
import JSZip from 'jszip'
import { PAGE_MARK, parseDocx, saveDocx, type HeaderFooter } from '@genoffice/docx-engine'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc, pmDocToSavePlan } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { lockedTableLines, tableLocked, updateTocField } from '../src/renderer/editor/toc-update'
import { applyTocPageDisplays } from '../src/renderer/editor/toc-refresh'
import { collectHeadings } from '../src/renderer/editor/headings'
import {
  fieldTableSites,
  setFieldTablesLocked,
  unlinkFieldTablesAtSelection,
  updateFieldTablesAtSelection,
  updateHitsLockedTable,
} from '../src/renderer/commands/reference-commands'
import { goToField, setFieldsLocked, unlinkFields } from '../src/renderer/editor/word-keys'
import {
  hfValueFromDoc,
  mountHfEditor,
  type HfEditorHandle,
} from '../src/renderer/editor/hf-editor'
import type { PmNode } from '../src/renderer/editor/convert'

const editors = new Set<Editor>()
const handles: HfEditorHandle[] = []
const realAlert = window.alert
beforeEach(() => {
  window.alert = () => {}
})
afterEach(() => {
  for (const editor of editors) editor.destroy()
  editors.clear()
  for (const h of handles.splice(0)) if (!h.editor.isDestroyed) h.exit()
  window.alert = realAlert
})

const heading = (text: string) =>
  `<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`
const pageRef = (anchor: string, page: string) =>
  '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
  `<w:r><w:instrText xml:space="preserve"> PAGEREF ${anchor} \\h </w:instrText></w:r>` +
  '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
  `<w:r><w:t>${page}</w:t></w:r>` +
  '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
const line = (text: string, anchor: string, page: string, lead = '', tail = '') =>
  '<w:p><w:pPr><w:pStyle w:val="TOC1"/><w:tabs><w:tab w:val="right" w:leader="dot" w:pos="9000"/></w:tabs></w:pPr>' +
  lead +
  `<w:hyperlink w:anchor="${anchor}" w:history="1"><w:r><w:t>${text}</w:t></w:r><w:r><w:tab/></w:r>${pageRef(anchor, page)}</w:hyperlink>` +
  tail +
  '</w:p>'
/** a table of contents as Word writes the field: begun in its first line, ended in its last */
const toc = (lock = false) =>
  line(
    'Stale One',
    '_Toc1',
    '4',
    `<w:r><w:fldChar w:fldCharType="begin"${lock ? ' w:fldLock="1"' : ''}/></w:r>` +
      '<w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h \\z \\u </w:instrText></w:r>' +
      '<w:r><w:fldChar w:fldCharType="separate"/></w:r>',
  ) + line('Stale Two', '_Toc2', '5', '', '<w:r><w:fldChar w:fldCharType="end"/></w:r>')

const BODY = `${toc()}${heading('Alpha')}<w:p><w:r><w:t>Body</w:t></w:r></w:p>${heading('Beta')}`

async function opened(bodyXml = BODY) {
  const bytes = await buildDocx({
    bodyXml,
    extraStylesXml:
      '<w:style w:type="paragraph" w:styleId="TOC1"><w:name w:val="toc 1"/><w:basedOn w:val="Normal"/></w:style>',
  })
  const parsed = await parseDocx(bytes)
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
  editors.add(editor)
  return { editor, parsed, bytes }
}

async function savedXml(editor: Editor, parsed: Awaited<ReturnType<typeof parseDocx>>) {
  const plan = pmDocToSavePlan(editor.getJSON() as never, parsed.blocks)
  const bytes = await saveDocx(parsed, plan.saveBlocks)
  return (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string')
}

const lines = (editor: Editor): string[] => {
  const out: string[] = []
  editor.state.doc.forEach((node) => {
    const shown = node.attrs.fieldDisplay as { left?: string; right?: string } | null
    out.push(
      node.type.name === 'docProtected'
        ? `[${shown?.left ?? ''}|${shown?.right ?? ''}]`
        : node.textContent,
    )
  })
  return out
}

describe('field commands on a table of contents', () => {
  it('locks the table from anywhere inside it, saves the lock, and unlocks it', async () => {
    const { editor, parsed } = await opened()
    const site = () => fieldTableSites(editor.state.doc, parsed.blocks)[0]
    expect(site().kind).toBe('toc')
    // the caret in the second line of the table
    editor.view.dispatch(
      editor.state.tr.setSelection(
        NodeSelection.create(editor.state.doc, editor.state.doc.firstChild!.nodeSize),
      ),
    )
    expect(setFieldTablesLocked(editor, parsed.blocks, true)).toBe(1)
    expect(tableLocked(editor.state.doc, parsed.blocks, site().range)).toBe(true)
    expect(setFieldTablesLocked(editor, parsed.blocks, true)).toBe(0)
    const xml = await savedXml(editor, parsed)
    expect(xml).toMatch(
      /<w:fldChar w:fldCharType="begin" w:fldLock="true"\/><\/w:r><w:r><w:instrText[^>]*> TOC /,
    )
    expect(xml.match(/w:fldLock/g)).toHaveLength(1)
    expect(xml).toContain('Stale Two')
    expect(setFieldTablesLocked(editor, parsed.blocks, false)).toBe(1)
    expect(await savedXml(editor, parsed)).not.toContain('w:fldLock')
    // one undo each
    editor.commands.undo()
    expect(tableLocked(editor.state.doc, parsed.blocks, site().range)).toBe(true)
  })

  it('a locked table is left alone by Update Table, by Update Field and by the page refresh', async () => {
    const { editor, parsed } = await opened(
      `${toc(true)}${heading('Alpha')}<w:p><w:r><w:t>Body</w:t></w:r></w:p>${heading('Beta')}`,
    )
    const before = lines(editor)
    expect(before.slice(0, 2)).toEqual(['[Stale One|4]', '[Stale Two|5]'])
    expect(updateHitsLockedTable(editor, parsed.blocks, 'toc')).toBe(true)
    expect(updateTocField(editor, parsed.blocks)).toBe(false)
    editor.commands.setNodeSelection(0)
    const host = {
      blocks: parsed.blocks,
      pagesOf: () => null,
      headingPages: () => null,
      citationHeading: () => '',
    }
    expect(updateFieldTablesAtSelection(editor, host)).toEqual({ updated: 0, locked: 1 })
    expect(lines(editor)).toEqual(before)
    // the refresh that follows the layout: both lines are the locked table's
    const locked = lockedTableLines(editor.state.doc, parsed.blocks)
    expect([...locked]).toEqual([0, editor.state.doc.firstChild!.nodeSize])
    const headings = collectHeadings(editor.state.doc)
    const tr = editor.state.tr
    const stale = headings.map((h) => ({
      ...h,
      text: h.text === 'Alpha' ? 'Stale One' : 'Stale Two',
    }))
    expect(applyTocPageDisplays(editor.state.doc, tr, stale, ['9', '9'], locked)).toBe(false)
    expect(applyTocPageDisplays(editor.state.doc, tr, stale, ['9', '9'])).toBe(true)
    // unlocked, Update Field on the selection rebuilds it from the headings
    expect(setFieldTablesLocked(editor, parsed.blocks, false)).toBe(1)
    expect(lockedTableLines(editor.state.doc, parsed.blocks).size).toBe(0)
    expect(updateFieldTablesAtSelection(editor, host)).toEqual({ updated: 1, locked: 0 })
    expect(lines(editor).slice(0, 2)).toEqual(['[Alpha|]', '[Beta|]'])
  })

  it('Next Field and Previous Field stop at the table', async () => {
    const { editor, parsed } = await opened()
    const stops = fieldTableSites(editor.state.doc, parsed.blocks).map((s) => s.range.from)
    editor.commands.setTextSelection(editor.state.doc.content.size - 1)
    expect(goToField(editor, -1, stops)).toBe(true)
    expect(editor.state.selection).toBeInstanceOf(NodeSelection)
    expect(editor.state.selection.from).toBe(0)
    expect(goToField(editor, -1, stops)).toBe(false)
    editor.commands.setTextSelection(editor.state.doc.content.size - 1)
    expect(goToField(editor, 1, stops)).toBe(false)
  })

  it('Unlink Fields takes the field away: the lines are saved as paragraphs and Update finds no table', async () => {
    const { editor, parsed, bytes } = await opened()
    const before = lines(editor)
    editor.commands.setNodeSelection(0)
    expect(await unlinkFieldTablesAtSelection(editor, parsed.blocks, bytes)).toBe(1)
    expect(fieldTableSites(editor.state.doc, parsed.blocks)).toHaveLength(0)
    // a paragraph in a TOC style is drawn as a table line with or without a field
    expect(lines(editor)).toEqual(before)
    const xml = await savedXml(editor, parsed)
    expect(xml).not.toContain('fldChar')
    expect(xml).not.toContain('instrText')
    expect(xml).not.toContain('w:hyperlink')
    expect(xml).toContain('Stale One')
    expect(xml).toContain('Stale Two')
    expect(xml.match(/<w:pStyle w:val="TOC1"\/>/g)).toHaveLength(2)
    editor.commands.undo()
    expect(fieldTableSites(editor.state.doc, parsed.blocks)).toHaveLength(1)
  })

  it('an unlinked index is ordinary paragraphs to type in', async () => {
    const entry = (text: string, page: string, lead = '', tail = '') =>
      '<w:p><w:pPr><w:pStyle w:val="Index1"/><w:tabs><w:tab w:val="right" w:leader="dot" w:pos="4000"/></w:tabs></w:pPr>' +
      `${lead}<w:r><w:t>${text}</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>${page}</w:t></w:r>${tail}</w:p>`
    const index =
      entry(
        'Apple',
        '1',
        '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
          '<w:r><w:instrText xml:space="preserve"> INDEX \\e "\t" \\c "1" </w:instrText></w:r>' +
          '<w:r><w:fldChar w:fldCharType="separate"/></w:r>',
      ) +
      // a line between the two ends of the field holds no field character
      entry('Fig', '1') +
      entry('Pear', '2', '', '<w:r><w:fldChar w:fldCharType="end"/></w:r>')
    const bytes = await buildDocx({
      bodyXml: `<w:p><w:r><w:t>Body</w:t></w:r></w:p>${index}`,
      extraStylesXml:
        '<w:style w:type="paragraph" w:styleId="Index1"><w:name w:val="index 1"/><w:basedOn w:val="Normal"/></w:style>',
    })
    const parsed = await parseDocx(bytes)
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks) as never,
    })
    editors.add(editor)
    const kinds = () => {
      const out: string[] = []
      editor.state.doc.forEach((node) => out.push(node.type.name))
      return out
    }
    expect(kinds()).toEqual(['docParagraph', 'docProtected', 'docProtected', 'docProtected'])
    expect(fieldTableSites(editor.state.doc, parsed.blocks).map((s) => s.kind)).toEqual(['index'])
    editor.commands.setNodeSelection(editor.state.doc.firstChild!.nodeSize)
    expect(await unlinkFieldTablesAtSelection(editor, parsed.blocks, bytes)).toBe(1)
    expect(kinds()).toEqual(['docParagraph', 'docParagraph', 'docParagraph', 'docParagraph'])
    expect(editor.state.doc.child(1).textContent).toBe('Apple\t1')
    expect(editor.state.doc.child(2).textContent).toBe('Fig\t1')
    expect(editor.state.doc.child(1).attrs.styleId).toBe('Index1')
    editor.commands.insertContentAt(editor.state.doc.firstChild!.nodeSize + 1, 'Green ')
    const xml = await savedXml(editor, parsed)
    expect(xml).not.toContain('fldChar')
    expect(xml).toContain('Green ')
    expect(xml).toContain('Pear')
    expect(xml.match(/<w:pStyle w:val="Index1"\/>/g)).toHaveLength(3)
    // the unlink is one step back
    editor.commands.undo()
    editor.commands.undo()
    expect(kinds()).toEqual(['docParagraph', 'docProtected', 'docProtected', 'docProtected'])
  })
})

describe('field commands on an index whose entries are ordinary paragraphs', () => {
  const entry = (text: string, lead = '', tail = '') =>
    `<w:p><w:pPr><w:pStyle w:val="Index1"/></w:pPr>${lead}<w:r><w:t>${text}</w:t></w:r>${tail}</w:p>`
  const BEGIN =
    '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
    '<w:r><w:instrText xml:space="preserve"> INDEX \\c "1" </w:instrText></w:r>' +
    '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
  const END = '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
  // pages after a comma; the field opens in its first entry or in a paragraph of its own
  const SHAPES = {
    'in its first entry':
      entry('Apple, 2', BEGIN) + entry('Banana, 3') + entry('Cherry, 4', '', END),
    'in a paragraph of its own': `<w:p>${BEGIN}</w:p>${entry('Apple, 2')}${entry('Banana, 3')}${entry('Cherry, 4')}<w:p>${END}</w:p>`,
  }
  for (const [shape, index] of Object.entries(SHAPES))
    it(`locks and unlocks the index from an entry that can be typed in (${shape})`, async () => {
      const bytes = await buildDocx({
        bodyXml: `<w:p><w:r><w:t>Body</w:t></w:r></w:p>${index}`,
        extraStylesXml:
          '<w:style w:type="paragraph" w:styleId="Index1"><w:name w:val="index 1"/><w:basedOn w:val="Normal"/></w:style>',
      })
      const parsed = await parseDocx(bytes)
      const editor = new Editor({
        element: document.createElement('div'),
        extensions: editorExtensions,
        content: blocksToPmDoc(parsed.blocks) as never,
      })
      editors.add(editor)
      const site = () => fieldTableSites(editor.state.doc, parsed.blocks)[0]
      let typed = -1
      editor.state.doc.forEach((node, pos) => {
        if (node.type.name === 'docParagraph' && node.textContent === 'Banana, 3') typed = pos
      })
      expect(typed).toBeGreaterThan(0)
      editor.commands.setTextSelection(typed + 2)
      setFieldsLocked(editor, true)
      expect(setFieldTablesLocked(editor, parsed.blocks, true)).toBe(1)
      expect(tableLocked(editor.state.doc, parsed.blocks, site().range)).toBe(true)
      expect(updateHitsLockedTable(editor, parsed.blocks, 'index')).toBe(true)
      const host = {
        blocks: parsed.blocks,
        pagesOf: () => null,
        headingPages: () => null,
        citationHeading: () => '',
      }
      expect(updateFieldTablesAtSelection(editor, host)).toEqual({ updated: 0, locked: 1 })
      const locked = await savedXml(editor, parsed)
      expect(locked).toMatch(/<w:fldChar w:fldCharType="begin" w:fldLock="true"\/>/)
      expect(locked).toContain('Banana, 3')
      expect(setFieldTablesLocked(editor, parsed.blocks, false)).toBe(1)
      expect(await savedXml(editor, parsed)).not.toContain('w:fldLock')
    })
})

describe('field commands on the page fields of a header or footer', () => {
  const FOOTER: HeaderFooter = {
    text: '',
    pageNumber: false,
    paras: [{ runs: [{ text: `Page ${PAGE_MARK} of the report` }] }],
  }
  function mount(value: HeaderFooter) {
    const host = document.createElement('div')
    document.body.append(host)
    const handle = mountHfEditor(host, {
      value,
      pageNo: '3',
      pageTotal: '9',
      spellcheck: false,
      onCommit: () => {},
      onExit: () => host.remove(),
    })
    handles.push(handle)
    return handle
  }
  const fieldPos = (editor: Editor): number => {
    let at = -1
    editor.state.doc.descendants((node, pos) => {
      if (node.type.name === 'hfField') at = pos
    })
    return at
  }

  it('Lock Fields gives the live page field the number it shows; Unlock Fields makes it live again', async () => {
    const { editor, changed } = mount(FOOTER)
    const initial = editor.getJSON() as PmNode
    editor.commands.setNodeSelection(fieldPos(editor))
    expect(setFieldsLocked(editor, true)).toBe(1)
    expect(fieldPos(editor)).toBe(-1)
    expect(editor.state.doc.textContent).toBe('Page 3 of the report')
    const value = hfValueFromDoc(editor.getJSON() as PmNode, FOOTER, initial)
    expect(value.paras![0].runs).toMatchObject([
      { text: 'Page ' },
      { text: '3', instrField: 'PAGE', fldLock: true },
      { text: ' of the report' },
    ])
    expect(changed()).not.toBeNull()
    // written with its lock and its number, and read back as the locked field
    const source = await buildDocx({ bodyXml: '<w:p><w:r><w:t>body</w:t></w:r></w:p>' })
    const bytes = await saveDocx(await parseDocx(source), [], { footer: value })
    const zip = await JSZip.loadAsync(bytes)
    const name = Object.keys(zip.files).find((n) => /^word\/footer\d*\.xml$/.test(n))!
    const xml = await zip.file(name)!.async('string')
    expect(xml).toMatch(/w:fldCharType="begin" w:fldLock="true"\/>/)
    expect((await parseDocx(bytes)).footerParas![0].runs[1]).toMatchObject({
      text: '3',
      instrField: 'PAGE',
      fldLock: true,
    })
    // the caret at the locked number
    editor.commands.setTextSelection(7)
    expect(setFieldsLocked(editor, false)).toBe(1)
    expect(fieldPos(editor)).toBeGreaterThan(0)
    expect(hfValueFromDoc(editor.getJSON() as PmNode, FOOTER, initial).paras![0].runs[0].text).toBe(
      `Page ${PAGE_MARK} of the report`,
    )
  })

  it('Unlink Fields leaves the number as text, and Next Field finds the page field', () => {
    const { editor } = mount(FOOTER)
    editor.commands.setTextSelection(1)
    expect(goToField(editor, 1)).toBe(true)
    expect(editor.state.selection.from).toBe(fieldPos(editor))
    expect(unlinkFields(editor)).toBe(1)
    expect(fieldPos(editor)).toBe(-1)
    expect(editor.state.doc.textContent).toBe('Page 3 of the report')
    expect(editor.state.doc.firstChild!.firstChild!.marks).toEqual([])
  })
})
