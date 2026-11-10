import { Editor } from '@tiptap/core'
import { TextSelection } from '@tiptap/pm/state'
import JSZip from 'jszip'
import { buildBlankDocx, parseDocx, saveDocx, type Block } from '@genoffice/docx-engine'
import { describe, expect, it } from 'vitest'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { ZoteroDocumentController } from '../src/renderer/zotero/controller'

type Parsed = Awaited<ReturnType<typeof parseDocx>>

const BIBLIOGRAPHY_RTF = '{\\rtf1 Alpha, A. (2024).\\par Beta, B. (2023).\\par Gamma, G. (2022).}'

function openEditor(blocks: Block[]): Editor {
  return new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: blocksToPmDoc(blocks) as never,
  })
}

const controllerFor = (editor: Editor) => {
  const controller = new ZoteroDocumentController(editor, { get: () => '', set: () => {} })
  return (command: string, args: unknown[]) =>
    controller.handle({ requestId: 'test', command, args })
}

/** a saved document whose three paragraphs are one Zotero bibliography field */
async function bibliographyDoc(): Promise<Parsed> {
  const blank = await parseDocx(await buildBlankDocx())
  const editor = openEditor(blank.blocks)
  try {
    const call = controllerFor(editor)
    const [id] = (await call('Document_insertField', [1, 'ReferenceMark', 0])) as [number]
    await call('Field_setCode', [1, id, 'BIBL {} CSL_BIBLIOGRAPHY'])
    await call('Field_setText', [1, id, BIBLIOGRAPHY_RTF, true])
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, blank.blocks)
    return parseDocx(await saveDocx(blank, plan.saveBlocks))
  } finally {
    editor.destroy()
  }
}

async function saveAndReopen(editor: Editor, parsed: Parsed) {
  const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
  const bytes = await saveDocx(parsed, plan.saveBlocks)
  const xml = await (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string')
  const blocks = (await parseDocx(bytes)).blocks.filter((block) => !block.hidden)
  return { xml, blocks }
}

const tokenCount = (xml: string, type: string) =>
  (xml.match(new RegExp(`w:fldCharType="${type}"`, 'g')) ?? []).length
const fieldRuns = (blocks: Block[]) =>
  blocks.flatMap((block) => block.runs ?? []).filter((run) => run.instrField !== undefined)
const partsOf = (blocks: Block[]) => fieldRuns(blocks).map((run) => run.zoteroFieldPart)
const idsOf = (blocks: Block[]) => new Set(fieldRuns(blocks).map((run) => run.zoteroFieldId))
const textsOf = (blocks: Block[]) => blocks.map((b) => (b.runs ?? []).map((r) => r.text).join(''))

function blockRange(editor: Editor, index: number): { from: number; to: number } {
  let from = 0
  for (let i = 0; i < index; i++) from += editor.state.doc.child(i).nodeSize
  return { from, to: from + editor.state.doc.child(index).nodeSize }
}

describe('Zotero field parts are recomputed before saving', () => {
  it('round-trips an untouched bibliography as one begin/inside/end field', async () => {
    const parsed = await bibliographyDoc()
    const blocks = parsed.blocks.filter((block) => !block.hidden)
    expect(partsOf(blocks)).toEqual(['begin', 'inside', 'end'])
    expect(idsOf(blocks).size).toBe(1)
  })

  it('deleting the paragraph with the field end keeps one balanced field', async () => {
    const parsed = await bibliographyDoc()
    const editor = openEditor(parsed.blocks)
    try {
      const { from, to } = blockRange(editor, 2)
      editor.view.dispatch(editor.state.tr.delete(from, to))
      const { xml, blocks } = await saveAndReopen(editor, parsed)
      expect([tokenCount(xml, 'begin'), tokenCount(xml, 'end')]).toEqual([1, 1])
      expect(textsOf(blocks)).toEqual(['Alpha, A. (2024).', 'Beta, B. (2023).'])
      expect(partsOf(blocks)).toEqual(['begin', 'end'])
      expect(idsOf(blocks).size).toBe(1)
    } finally {
      editor.destroy()
    }
  })

  it('deleting the paragraph with the field begin keeps one balanced field', async () => {
    const parsed = await bibliographyDoc()
    const editor = openEditor(parsed.blocks)
    try {
      const { from, to } = blockRange(editor, 0)
      editor.view.dispatch(editor.state.tr.delete(from, to))
      const { xml, blocks } = await saveAndReopen(editor, parsed)
      expect([tokenCount(xml, 'begin'), tokenCount(xml, 'end')]).toEqual([1, 1])
      expect(textsOf(blocks)).toEqual(['Beta, B. (2023).', 'Gamma, G. (2022).'])
      expect(partsOf(blocks)).toEqual(['begin', 'end'])
    } finally {
      editor.destroy()
    }
  })

  it('splitting the begin run with Enter does not duplicate the field begin', async () => {
    const parsed = await bibliographyDoc()
    const editor = openEditor(parsed.blocks)
    try {
      editor.view.dispatch(editor.state.tr.setSelection(TextSelection.create(editor.state.doc, 7)))
      editor.commands.splitBlock()
      expect(editor.state.doc.childCount).toBe(4)
      const { xml, blocks } = await saveAndReopen(editor, parsed)
      expect([tokenCount(xml, 'begin'), tokenCount(xml, 'end')]).toEqual([1, 1])
      expect(textsOf(blocks)).toEqual([
        'Alpha,',
        ' A. (2024).',
        'Beta, B. (2023).',
        'Gamma, G. (2022).',
      ])
      expect(partsOf(blocks)).toEqual(['begin', 'inside', 'inside', 'end'])
      expect(idsOf(blocks).size).toBe(1)
    } finally {
      editor.destroy()
    }
  })

  it('a field reduced to one run is written as a single-paragraph field', async () => {
    const parsed = await bibliographyDoc()
    const editor = openEditor(parsed.blocks)
    try {
      const { from } = blockRange(editor, 1)
      const { to } = blockRange(editor, 2)
      editor.view.dispatch(editor.state.tr.delete(from, to))
      const { xml, blocks } = await saveAndReopen(editor, parsed)
      expect([tokenCount(xml, 'begin'), tokenCount(xml, 'end')]).toEqual([1, 1])
      expect(partsOf(blocks)).toEqual(['single'])
    } finally {
      editor.destroy()
    }
  })

  it('a citation inserted between bibliography paragraphs splits it into two fields', async () => {
    const parsed = await bibliographyDoc()
    const editor = openEditor(parsed.blocks)
    try {
      const { to } = blockRange(editor, 0)
      editor.view.dispatch(
        editor.state.tr.setSelection(TextSelection.create(editor.state.doc, to - 1)),
      )
      editor.commands.splitBlock()
      const call = controllerFor(editor)
      await call('Document_insertField', [1, 'ReferenceMark', 0])
      const { xml, blocks } = await saveAndReopen(editor, parsed)
      expect([tokenCount(xml, 'begin'), tokenCount(xml, 'end')]).toEqual([3, 3])
      expect(textsOf(blocks)).toEqual([
        'Alpha, A. (2024).',
        '{Citation}',
        'Beta, B. (2023).',
        'Gamma, G. (2022).',
      ])
      expect(partsOf(blocks)).toEqual(['single', 'single', 'begin', 'end'])
      expect(idsOf(blocks).size).toBe(3)
    } finally {
      editor.destroy()
    }
  })

  it('runs that lost their ids on a clipboard round trip save as one field again', async () => {
    const parsed = await bibliographyDoc()
    const editor = openEditor(parsed.blocks)
    try {
      const type = editor.state.schema.marks.instrField
      let transaction = editor.state.tr
      editor.state.doc.descendants((node, pos) => {
        const mark = node.marks.find((candidate) => candidate.type === type)
        if (!node.isText || !mark) return
        transaction = transaction
          .removeMark(pos, pos + node.nodeSize, type)
          .addMark(
            pos,
            pos + node.nodeSize,
            type.create({ ...mark.attrs, fieldId: null, fieldPart: null }),
          )
      })
      editor.view.dispatch(transaction)
      const { xml, blocks } = await saveAndReopen(editor, parsed)
      expect([tokenCount(xml, 'begin'), tokenCount(xml, 'end')]).toEqual([1, 1])
      expect(partsOf(blocks)).toEqual(['begin', 'inside', 'end'])
      expect(idsOf(blocks).size).toBe(1)
    } finally {
      editor.destroy()
    }
  })

  it('two pasted copies of a citation with text between them save as two fields', async () => {
    const parsed = await parseDocx(await buildBlankDocx())
    const editor = openEditor(parsed.blocks)
    try {
      const lost = { type: 'instrField', attrs: { instr: 'ADDIN ZOTERO_ITEM CSL_CITATION {}' } }
      editor.commands.insertContent([
        { type: 'text', text: '(Doe)', marks: [lost] },
        { type: 'text', text: ' and again ' },
        { type: 'text', text: '(Doe)', marks: [lost] },
      ])
      const { xml, blocks } = await saveAndReopen(editor, parsed)
      expect([tokenCount(xml, 'begin'), tokenCount(xml, 'end')]).toEqual([2, 2])
      expect(textsOf(blocks)).toEqual(['(Doe) and again (Doe)'])
      expect(partsOf(blocks)).toEqual(['single', 'single'])
      expect(idsOf(blocks).size).toBe(2)
    } finally {
      editor.destroy()
    }
  })
})
