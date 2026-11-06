import { describe, expect, it } from 'vitest'
import type { Editor } from '@tiptap/core'
import {
  buildBlankDocx,
  mergeFieldInstruction,
  parseDocx,
  readSections,
} from '@genoffice/docx-engine'
import { blocksToPmDoc, type PmNode } from '../src/renderer/editor/convert'
import { buildMailMergeFiles } from '../src/renderer/mail-merge/output'
import { countRecordSlots, hasNextField } from '../src/renderer/mail-merge/template'
import type { FileActionContext } from '../src/renderer/file-actions'

/**
 * The NEXT field, end to end: a template that advances the record cursor mid
 * page, so one page holds several records. genoffice#2374.
 */

/** the template: the first record on one line, the next one after the NEXT */
function template(): PmNode {
  return {
    type: 'doc',
    content: [
      {
        type: 'docParagraph',
        content: [
          {
            type: 'text',
            text: 'A',
            marks: [{ type: 'instrField', attrs: { instr: mergeFieldInstruction('FirstName') } }],
          },
        ],
      },
      {
        type: 'docParagraph',
        content: [
          { type: 'text', text: 'N', marks: [{ type: 'instrField', attrs: { instr: 'NEXT' } }] },
          {
            type: 'text',
            text: 'B',
            marks: [{ type: 'instrField', attrs: { instr: mergeFieldInstruction('FirstName') } }],
          },
        ],
      },
    ],
  }
}

/** the merge context the app hands the merge */
async function context(doc: PmNode): Promise<FileActionContext> {
  const parsed = await parseDocx(await buildBlankDocx())
  const editor = { getJSON: () => doc, storage: {} } as unknown as Editor
  return {
    editor,
    doc: { parsed, fileName: 'letter.docx' },
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

const records = [{ FirstName: 'Ada' }, { FirstName: 'Ben' }, { FirstName: 'Grace' }]

describe('the NEXT field', () => {
  it('counts the slots a page holds, not the NEXT fields', () => {
    const doc = template()
    expect(hasNextField(doc)).toBe(true)
    // one NEXT, but the page holds two records: the first, then the one the
    // NEXT brings in
    expect(countRecordSlots(doc)).toBe(2)
  })

  it('merges several records into one document, a page at a time', async () => {
    const ctx = await context(template())
    const files = await buildMailMergeFiles(ctx, {
      mode: 'single',
      records,
      fields: ['FirstName'],
      namePattern: '',
    })
    expect(files).toHaveLength(1)
    const text = await textOf(files[0]!.bytes)
    for (const record of records) expect(text).toContain(record.FirstName)
    // the NEXT rule itself leaves nothing behind
    expect(text).not.toContain('N')
  })

  it('writes one file per page when the output is a file per record', async () => {
    const ctx = await context(template())
    const files = await buildMailMergeFiles(ctx, {
      mode: 'perRecord',
      records,
      fields: ['FirstName'],
      namePattern: 'FirstName',
    })
    // three records, two per page: two files
    expect(files).toHaveLength(2)
    expect(files[0]!.name).toBe('01-Ada.docx')
    expect(files[1]!.name).toBe('02-Grace.docx')
  })
})
