import { describe, expect, it } from 'vitest'

import { createBufferEntrySource, planCellEditsToXlsx } from '../src/gateway/xlsx-gateway'
import type { SheetNoteState } from '../src/gateway/xlsx-gateway'
import { buildEditFixture } from './fixture-builder'

async function planNotes(noteStates: SheetNoteState[]) {
  const source = await createBufferEntrySource(await buildEditFixture())
  return planCellEditsToXlsx(
    source,
    [],
    [],
    [],
    undefined,
    [],
    [],
    [],
    [],
    [],
    null,
    [],
    [],
    noteStates,
  )
}

const NOTES: SheetNoteState[] = [
  {
    sheetName: 'Data',
    notes: [
      { row: 0, column: 1, author: 'Reviewer', text: 'Tax inclusive <confirm>' },
      { row: 4, column: 0, author: '', text: 'second note' },
    ],
  },
]

describe('note snapshots', () => {
  it('creates the comments part with authors, refs, and escaped text', async () => {
    const plan = await planNotes(NOTES)
    const comments = [...plan.added.entries()].find(([path]) => /xl\/comments\d+\.xml/.test(path))
    expect(comments).toBeDefined()
    const xml = comments![1]
    expect(xml).toContain('<author>Reviewer</author>')
    expect(xml).toContain('<comment ref="B1" authorId="0">')
    expect(xml).toContain('Tax inclusive &lt;confirm&gt;')
    expect(xml).toContain('<comment ref="A5" authorId="1">')
  })

  it('creates the VML drawing with one Note shape per comment', async () => {
    const plan = await planNotes(NOTES)
    const vml = [...plan.added.entries()].find(([path]) =>
      /xl\/drawings\/vmlDrawing\d+\.vml/.test(path),
    )
    expect(vml).toBeDefined()
    expect(vml![1].match(/ObjectType="Note"/g)).toHaveLength(2)
    expect(vml![1]).toContain('<x:Row>0</x:Row><x:Column>1</x:Column>')
    expect(vml![1]).toContain('<x:Row>4</x:Row><x:Column>0</x:Column>')
  })

  it('registers rels, content types, and the legacyDrawing element', async () => {
    const plan = await planNotes(NOTES)
    const rels =
      plan.replaced.get('xl/worksheets/_rels/sheet1.xml.rels') ??
      plan.added.get('xl/worksheets/_rels/sheet1.xml.rels')
    expect(rels).toContain('relationships/comments')
    expect(rels).toContain('relationships/vmlDrawing')
    const contentTypes = plan.replaced.get('[Content_Types].xml')
    expect(contentTypes).toContain('spreadsheetml.comments+xml')
    expect(contentTypes).toContain('Extension="vml"')
    const worksheet = plan.replaced.get('xl/worksheets/sheet1.xml')
    expect(worksheet).toContain('<legacyDrawing r:id="')
  })

  it('is a no-op when clearing notes on a sheet that never had any', async () => {
    const plan = await planNotes([{ sheetName: 'Data', notes: [] }])
    expect([...plan.added.keys()].some((path) => path.includes('comments'))).toBe(false)
    expect(plan.replaced.has('xl/worksheets/sheet1.xml')).toBe(false)
  })
})
