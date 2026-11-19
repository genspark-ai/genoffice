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
import type { FileActionContext } from '../src/renderer/file-actions'
import { buildMailMergeFiles } from '../src/renderer/mail-merge/output'
import { AVERY_5160, envelopeLayout } from '../src/renderer/mail-merge/layout'

const live: Editor[] = []
afterEach(() => {
  for (const editor of live.splice(0)) editor.destroy()
})

/** the context the app hands the merge: settling the font settings returns the live one */
async function appContext(): Promise<FileActionContext> {
  const parsed = await parseDocx(await buildBlankDocx())
  const sections = readSections(parsed)
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks, sections) as never,
  })
  live.push(editor)
  editor.commands.insertContent({
    type: 'text',
    text: '«Name»',
    marks: [{ type: 'instrField', attrs: { instr: mergeFieldInstruction('Name') } }],
  })
  const ctx = {
    editor,
    doc: { parsed, fileName: 'template.docx', filePath: null },
    section: sections.at(-1)?.settings ?? null,
    sectionDirty: false,
    sections,
    sectionsDirty: [],
    pgNumDirtySections: [],
    sectionHfEdits: {},
    hfLinks: {},
    inksDirty: false,
    inkAnnotations: [],
    styleUpserts: {},
    settleFontSettings: async () => ctx,
  } as unknown as FileActionContext
  return ctx
}

async function paperOf(bytes: Uint8Array) {
  const settings = readSections(await parseDocx(bytes)).at(-1)!.settings
  return {
    width: settings.pageWidth,
    height: settings.pageHeight,
    left: settings.marginLeft,
    top: settings.marginTop,
  }
}

const request = { records: [{ Name: 'Ada' }], fields: ['Name'], namePattern: '' }

describe('the paper a merge prints on', () => {
  it('is the label stock, not the template page', async () => {
    const [file] = await buildMailMergeFiles(await appContext(), {
      ...request,
      mode: 'labels',
      layout: AVERY_5160,
    })
    expect(await paperOf(file!.bytes)).toEqual({
      width: AVERY_5160.pageWidth,
      height: AVERY_5160.pageHeight,
      left: AVERY_5160.marginLeft,
      top: AVERY_5160.marginTop,
    })
  })

  it('is the envelope, not the template page', async () => {
    const layout = envelopeLayout('dl')
    const [file] = await buildMailMergeFiles(await appContext(), {
      ...request,
      mode: 'envelopes',
      layout,
    })
    expect(await paperOf(file!.bytes)).toEqual({
      width: layout.pageWidth,
      height: layout.pageHeight,
      left: layout.marginLeft,
      top: layout.marginTop,
    })
  })

  it('stays the template page for a letter', async () => {
    const ctx = await appContext()
    const [file] = await buildMailMergeFiles(ctx, { ...request, mode: 'perRecord' })
    expect((await paperOf(file!.bytes)).width).toBe(ctx.section!.pageWidth)
  })
})
