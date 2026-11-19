/** AI merge tool glue: spreadsheet attachment selection and tool execution. */
import { describe, expect, it } from 'vitest'
import { createMergeSkill, spreadsheetAttachmentIndexes } from '../src/renderer/ai/merge-skill'
import type { AttachmentMeta } from '../src/shared/desktop-api'

const att = (name: string, ext: string): AttachmentMeta => ({
  path: `/tmp/${name}`,
  name,
  ext,
  sizeBytes: 10,
})

describe('spreadsheetAttachmentIndexes', () => {
  it('selects only spreadsheet attachments', () => {
    const list = [
      att('a.pdf', 'pdf'),
      att('b.xlsx', 'xlsx'),
      att('c.csv', 'csv'),
      att('d.png', 'png'),
    ]
    expect(spreadsheetAttachmentIndexes(list)).toEqual([1, 2])
  })
})

describe('merge_attached_workbooks', () => {
  const run = (list: AttachmentMeta[], input: object, merged: string[][] = []) => {
    const skill = createMergeSkill({
      getAttachments: () => list,
      mergePaths: async (paths) => {
        merged.push(paths)
        return { importedSheets: paths.length, files: paths.length, sheetNames: paths }
      },
    })
    return skill.executeTool!({ name: 'merge_attached_workbooks', input, id: 't1' } as never)
  }

  it('merges every spreadsheet attachment by default and skips others', async () => {
    const merged: string[][] = []
    const result = await run([att('a.pdf', 'pdf'), att('b.xlsx', 'xlsx')], {}, merged)
    expect(merged).toEqual([['/tmp/b.xlsx']])
    expect(result.isError).toBeFalsy()
    expect(result.output).toContain('Merged 1 sheets')
  })

  it('respects explicit indexes and reports skips', async () => {
    const merged: string[][] = []
    const result = await run(
      [att('a.xlsx', 'xlsx'), att('b.pdf', 'pdf'), att('c.csv', 'csv')],
      { indexes: [0, 1, 9] },
      merged,
    )
    expect(merged).toEqual([['/tmp/a.xlsx']])
    expect(result.output).toContain('Skipped: b.pdf (not a spreadsheet); #9 (no such attachment)')
  })

  it('errors cleanly with no spreadsheet sources', async () => {
    const result = await run([att('a.pdf', 'pdf')], {})
    expect(result.isError).toBe(true)
  })
})
