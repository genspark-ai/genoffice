import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'

import {
  createBufferEntrySource,
  planCellEditsToXlsx,
} from '@genoffice/xlsx-gateway/gateway/xlsx-gateway'
import type { SheetNoteState } from '@genoffice/xlsx-gateway/gateway/xlsx-gateway'
import { buildEditFixture, buildKitchenSinkFixture } from './fixture-builder'

async function planNotes(noteStates: SheetNoteState[], fixture?: Buffer) {
  const source = await createBufferEntrySource(fixture ?? (await buildEditFixture()))
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

async function kitchenSinkWithControlThenNote(): Promise<Buffer> {
  const zip = await JSZip.loadAsync(await buildKitchenSinkFixture())
  zip.file(
    'xl/drawings/vmlDrawing1.vml',
    '<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">' +
      '<v:shape id="_x0000_s1025" type="#_x0000_t201" style="position:absolute"/>' +
      '<v:shape id="_x0000_s1026" type="#_x0000_t202">' +
      '<x:ClientData ObjectType="Note"><x:Anchor>2, 15, 0, 2, 4, 15, 4, 2</x:Anchor>' +
      '<x:Row>0</x:Row><x:Column>1</x:Column></x:ClientData></v:shape>' +
      '</xml>',
  )
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })
}

describe('note removal', () => {
  it('removes the note shape and keeps the self-closing control before it', async () => {
    const plan = await planNotes(
      [{ sheetName: 'Data', notes: [] }],
      await kitchenSinkWithControlThenNote(),
    )
    const vml = plan.replaced.get('xl/drawings/vmlDrawing1.vml')
    expect(vml).toBeDefined()
    expect(vml).toContain('<v:shape id="_x0000_s1025"')
    expect(vml).not.toContain('ObjectType="Note"')
  })
})

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

  it('encodes XML-forbidden controls in note text and authors', async () => {
    const plan = await planNotes([
      {
        sheetName: 'Data',
        notes: [{ row: 0, column: 0, author: 'A\u0001B', text: 'C\u000bD' }],
      },
    ])
    const comments = [...plan.added.entries()].find(([path]) => /xl\/comments\d+\.xml/.test(path))
    const xml = comments![1]
    expect(xml).toContain('<author>A_x0001_B</author>')
    expect(xml).toContain('C_x000B_D')
    expect(xml).not.toContain('\u0001')
    expect(xml).not.toContain('\u000b')
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

  it('allocates a relationship id without spreading a large id set', async () => {
    const zip = await JSZip.loadAsync(await buildEditFixture())
    const relationships = Array.from(
      { length: 130_000 },
      (_, index) =>
        `<Relationship Id="rId${index + 1}" Type="https://example.test/relationship" Target="part${index + 1}"/>`,
    ).join('')
    zip.file(
      'xl/worksheets/_rels/sheet1.xml.rels',
      `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${relationships}</Relationships>`,
    )
    const fixture = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })
    const plan = await planNotes(NOTES, fixture)
    const rels =
      plan.replaced.get('xl/worksheets/_rels/sheet1.xml.rels') ??
      plan.added.get('xl/worksheets/_rels/sheet1.xml.rels')
    expect(rels).toContain('Id="rId130001"')
  })

  it('starts new note shape ids after the ids the part already holds', async () => {
    const plan = await planNotes(NOTES, await buildKitchenSinkFixture())
    const vml =
      plan.replaced.get('xl/drawings/vmlDrawing1.vml') ??
      plan.added.get('xl/drawings/vmlDrawing1.vml')
    expect(vml).toBeDefined()
    const ids = [...vml!.matchAll(/id="_x0000_s(\d+)"/g)].map((match) => match[1])
    expect(ids).toHaveLength(3)
    expect(new Set(ids).size).toBe(ids.length)
  })

  it('is a no-op when clearing notes on a sheet that never had any', async () => {
    const plan = await planNotes([{ sheetName: 'Data', notes: [] }])
    expect([...plan.added.keys()].some((path) => path.includes('comments'))).toBe(false)
    expect(plan.replaced.has('xl/worksheets/sheet1.xml')).toBe(false)
  })
})
