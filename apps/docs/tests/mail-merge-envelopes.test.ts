import { afterEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import {
  buildBlankDocx,
  mergeFieldInstruction,
  parseDocx,
  readSections,
} from '@genoffice/docx-engine'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { blocksToPmDoc } from '../src/renderer/editor/convert'
import { buildMailMergeFiles } from '../src/renderer/mail-merge/output'
import { AVERY_5160, envelopeLayout } from '../src/renderer/mail-merge/layout'
import type { FileActionContext } from '../src/renderer/file-actions'

/**
 * The envelope and the label sheet, end to end: the same template, the same
 * records, the same request — and the documents must come out on the right
 * paper with the merged values and no field code. genoffice#2097, genoffice#2374.
 */

const live: Editor[] = []
afterEach(() => {
  for (const editor of live.splice(0)) editor.destroy()
})

/** a template with two merge fields in it, and the editor that holds it */
async function template(): Promise<Editor> {
  const bytes = await buildBlankDocx()
  const parsed = await parseDocx(bytes)
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks, readSections(parsed)) as never,
  })
  live.push(editor)
  editor.commands.insertContent({
    type: 'text',
    text: 'AdaFirst',
    marks: [{ type: 'instrField', attrs: { instr: mergeFieldInstruction('FirstName') } }],
  })
  editor.commands.insertContent({ type: 'text', text: ' ' })
  editor.commands.insertContent({
    type: 'text',
    text: 'Last',
    marks: [{ type: 'instrField', attrs: { instr: mergeFieldInstruction('LastName') } }],
  })
  return editor
}

/** the merge context the app hands the merge: the editor, the file, its section */
async function context(editor: Editor, fileName = 'letter.docx'): Promise<FileActionContext> {
  const parsed = await parseDocx(await buildBlankDocx())
  return {
    editor,
    doc: { parsed, fileName },
    section: null,
    sectionDirty: false,
    pageCount: 1,
    inkAnnotations: [],
    sections: [],
    sectionsDirty: [],
    pgNumDirtySections: [],
    sectionHfEdits: {},
    hfLinks: {},
    styleUpserts: {},
  } as unknown as FileActionContext
}

/** the text of a document, as the reader sees it */
async function textOf(bytes: Uint8Array): Promise<string> {
  const parsed = await parseDocx(bytes)
  return parsed.blocks.map((block) => (block.runs ?? []).map((run) => run.text).join('')).join('\n')
}

/** the page size of a document, in twips */
async function pageOf(bytes: Uint8Array): Promise<{ width: number; height: number }> {
  const parsed = await parseDocx(bytes)
  const section = readSections(parsed).at(-1)?.settings
  return { width: section?.pageWidth ?? 0, height: section?.pageHeight ?? 0 }
}

const records = [
  { FirstName: 'Ada', LastName: 'Lovelace' },
  { FirstName: 'Ben', LastName: 'Franklin' },
]

describe('the envelope output', () => {
  it('writes one document holding every envelope, on the envelope paper', async () => {
    const editor = await template()
    const ctx = await context(editor)
    const files = await buildMailMergeFiles(ctx, {
      mode: 'envelopes',
      records,
      fields: ['FirstName', 'LastName'],
      namePattern: 'LastName',
      layout: envelopeLayout('size10'),
    })
    // one document now holds all the envelopes (one per page); not one file per record
    expect(files).toHaveLength(1)
    expect(files[0]!.name).toBe('letter-merged.docx')
    const envelope = envelopeLayout('size10')
    const page = await pageOf(files[0]!.bytes)
    expect(page.width).toBe(envelope.pageWidth)
    expect(page.height).toBe(envelope.pageHeight)
    // every envelope's merged values are in the one document, and no field code is
    const text = await textOf(files[0]!.bytes)
    expect(text).toContain('Ada')
    expect(text).toContain('Lovelace')
    expect(text).toContain('Ben')
    expect(text).toContain('Franklin')
    expect(text).not.toContain('First')
  })
})

describe('the label sheet output', () => {
  it('writes one label sheet holding every record', async () => {
    const editor = await template()
    const ctx = await context(editor)
    const files = await buildMailMergeFiles(ctx, {
      mode: 'labels',
      records,
      fields: ['FirstName', 'LastName'],
      namePattern: '',
      layout: AVERY_5160,
    })
    expect(files).toHaveLength(1)
    const page = await pageOf(files[0]!.bytes)
    expect(page.width).toBe(AVERY_5160.pageWidth)
    expect(page.height).toBe(AVERY_5160.pageHeight)
    // the records sit in the label grid's cells, not in the body paragraphs
    const sheet = await parseDocx(files[0]!.bytes)
    const grid = sheet.blocks.find((block) => block.type === 'table')
    expect(grid?.table?.rows[0]?.[0]?.paras.join('')).toContain('Ada')
  })

  it('writes the merged names into the Avery 5160 grid, with no field code', async () => {
    const editor = await template()
    const ctx = await context(editor)
    const files = await buildMailMergeFiles(ctx, {
      mode: 'labels',
      records,
      fields: ['FirstName', 'LastName'],
      namePattern: '',
      layout: AVERY_5160,
    })
    // one label sheet holds every record, on the Avery 5160 paper
    expect(files).toHaveLength(1)
    expect(files[0]!.name).toBe('letter-merged.docx')
    const page = await pageOf(files[0]!.bytes)
    expect(page.width).toBe(AVERY_5160.pageWidth)
    expect(page.height).toBe(AVERY_5160.pageHeight)
    // the records sit in the label grid's cells, not in the body paragraphs
    const sheet = await parseDocx(files[0]!.bytes)
    const grid = sheet.blocks.find((block) => block.type === 'table')
    const rows = grid?.table?.rows
    // Avery 5160: one row here holds the two records — 3 label columns plus 2 gutters
    expect(rows).toHaveLength(1)
    expect(rows![0]).toHaveLength(5)
    // first label cell carries the first record's full merged name
    expect(rows![0]![0]!.paras.join('')).toContain('Ada Lovelace')
    // the label after the first gutter carries the second record
    expect(rows![0]![2]!.paras.join('')).toContain('Ben Franklin')
    // every merged value is on the sheet, and no field-code placeholder survives
    const text = (rows ?? []).flatMap((row) => row.map((cell) => cell.paras.join(''))).join('\n')
    expect(text).toContain('Ada')
    expect(text).toContain('Lovelace')
    expect(text).toContain('Ben')
    expect(text).toContain('Franklin')
    expect(text).not.toContain('First')
    expect(text).not.toContain('Last')
  })
})
