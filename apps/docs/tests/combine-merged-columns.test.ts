import { afterEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { compareDocuments, parseDocx, readSections, saveDocx } from '@genoffice/docx-engine'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import {
  SHAPES,
  baseGrid,
  gridShape,
  gridXml,
  withColumn,
  withRow,
  withoutColumn,
  withoutRow,
  xmlShape,
  type Grid,
} from '../../../packages/docx-engine/tests/helpers/merged-tables'
import { applyTrackedChangesXml } from '../../../packages/docx-engine/src/revision-xml'
import { editorExtensions } from '../src/renderer/editor/extensions'
import {
  blocksToPmDoc,
  pmDocOptions,
  pmDocToSavePlan,
  type PmNode,
} from '../src/renderer/editor/convert'
import { combineDocuments } from '../src/renderer/editor/combine'
import {
  acceptAllRevisions,
  collectRevisions,
  rejectAllRevisions,
} from '../src/renderer/editor/revisions'

;(globalThis as { CSS?: unknown }).CSS ??= { escape: (s: string) => s }

const para = (text: string) => `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`
const around = (grid: Grid) =>
  para('This agreement is made between the supplier and the customer.') +
  gridXml(grid) +
  para('Signed by both parties.')

const editors: Editor[] = []
afterEach(() => {
  for (const editor of editors.splice(0)) editor.destroy()
})

async function open(bytes: Uint8Array) {
  const parsed = await parseDocx(bytes)
  const json = blocksToPmDoc(parsed.blocks, readSections(parsed), pmDocOptions(parsed))
  return { parsed, json }
}

function edit(json: PmNode): Editor {
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    editorProps: { handleScrollToSelection: () => true },
    content: json,
  })
  editors.push(editor)
  return editor
}

async function saved(editor: Editor, parsed: Awaited<ReturnType<typeof parseDocx>>) {
  const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
  const bytes = await saveDocx(parsed, plan.saveBlocks)
  return { bytes, xml: (await parseDocx(bytes)).internal.documentXml }
}

/** the table of the open document after every change was accepted or rejected there, as the file then saved */
async function settled(bytes: Uint8Array, mode: 'accept' | 'reject') {
  const { parsed, json } = await open(bytes)
  const editor = edit(json)
  if (mode === 'accept') acceptAllRevisions(editor)
  else rejectAllRevisions(editor)
  expect(collectRevisions(editor.state.doc)).toEqual([])
  return xmlShape((await saved(editor, parsed)).xml)
}

async function combine(original: Grid, revised: Grid) {
  const mine = await open(await buildDocx({ bodyXml: around(original) }))
  const theirs = await open(await buildDocx({ bodyXml: around(revised) }))
  const editor = edit(mine.json)
  const summary = combineDocuments(editor, theirs.json.content ?? [], {
    author: 'Rae Visor',
    date: '2026-11-24T00:00:00Z',
    formatting: true,
  })
  const file = await saved(editor, mine.parsed)
  const resolved = (mode: 'accept' | 'reject') => {
    const out = applyTrackedChangesXml(file.xml, mode)
    expect(out.remaining).toBe(0)
    return xmlShape(out.xml)
  }
  return { editor, summary, file, resolved }
}

/** accept all gives the revised table and reject all the original: in the saved file, and in the editor after a reopen */
async function expectBothWays(original: Grid, revised: Grid) {
  const result = await combine(original, revised)
  expect(result.summary.skipped).toBe(0)
  expect(result.resolved('accept')).toEqual(gridShape(revised))
  expect(result.resolved('reject')).toEqual(gridShape(original))
  expect(await settled(result.file.bytes, 'accept')).toEqual(gridShape(revised))
  expect(await settled(result.file.bytes, 'reject')).toEqual(gridShape(original))
  return result
}

describe('Combine into the open document: a column added to or taken out of a table with merged cells', () => {
  for (const [name, merges] of Object.entries(SHAPES)) {
    const original = baseGrid(merges)
    for (let at = 0; at <= 4; at++) {
      const { grid: revised, reaches } = withColumn(original, at)
      it(`${name}; a column comes in at ${at}${reaches ? ' (inside the merged cell)' : ''}`, async () => {
        if (reaches) {
          // a merged cell that reaches across the column: left alone and counted
          const { editor, summary } = await combine(original, revised)
          expect(summary.skipped).toBe(1)
          expect(collectRevisions(editor.state.doc)).toEqual([])
          return
        }
        const { editor, summary, file } = await expectBothWays(original, revised)
        expect(summary).toMatchObject({ insertions: 1, deletions: 0 })
        const kinds = collectRevisions(editor.state.doc).map((rev) => rev.kind)
        expect(new Set(kinds)).toEqual(new Set(['cellIns']))
        expect(file.xml.match(/<w:cellIns /g)).toHaveLength(4)
      })
    }
    for (let at = 0; at < 4; at++) {
      const { grid: revised, reaches } = withoutColumn(original, at)
      it(`${name}; column ${at} goes${reaches ? ' (out of the merged cell)' : ''}`, async () => {
        if (reaches) {
          const { summary, resolved } = await combine(original, revised)
          // replaced by column where both of the cell's columns go, else left alone and counted
          if (summary.skipped === 0) {
            expect(resolved('accept')).toEqual(gridShape(revised))
            expect(resolved('reject')).toEqual(gridShape(original))
          }
          return
        }
        const { summary, file } = await expectBothWays(original, revised)
        expect(summary.deletions).toBeGreaterThanOrEqual(1)
        // the row holding only the lower half of the block reads too differently without the
        // column to be paired with its twin: it goes and comes back, each given a cell of the column
        const unpaired = name.startsWith('a block') && (at === 0 || at === 3)
        expect(file.xml.match(/<w:cellDel /g)).toHaveLength(unpaired ? 5 : 4)
        expect(file.xml).not.toContain('<w:cellIns ')
      })
    }
  }
})

describe('a row that reads the same in both tables, with columns that go and come in', () => {
  const same = ['n/a', 'n/a', 'n/a', 'n/a', 'n/a']
  const original: Grid = {
    texts: [
      ['Item', 'Stock', '', 'Price', 'Origin'],
      same,
      ['Bolt', 'many', 'on hand', 'ten dollars', 'Leeds'],
      ['Nut', 'few', 'on order', 'two dollars', 'Turin'],
    ],
    merges: [{ r0: 0, r1: 0, c0: 1, c1: 2 }],
  }
  const revised: Grid = {
    texts: [
      ['Item', 'Price', 'Origin', 'Maker', 'Since'],
      same,
      ['Bolt', 'ten dollars', 'Leeds', 'Acme', 'last year'],
      ['Nut', 'two dollars', 'Turin', 'Borg', 'this year'],
    ],
    merges: [],
  }

  it('into the open document: the row gets its cells for those columns', async () => {
    const { file } = await expectBothWays(original, revised)
    const widths = xmlShape(file.xml).map((row) => row.reduce((sum, [, span]) => sum + span, 0))
    expect(new Set(widths)).toEqual(new Set([7]))
  })

  it('into a new document, opened in the editor', async () => {
    const { bytes } = await compareDocuments(
      await parseDocx(await buildDocx({ bodyXml: around(original) })),
      await parseDocx(await buildDocx({ bodyXml: around(revised) })),
      { author: 'Rae Visor', date: '2026-11-24T00:00:00Z' },
    )
    expect(await settled(bytes, 'accept')).toEqual(gridShape(revised))
    expect(await settled(bytes, 'reject')).toEqual(gridShape(original))
  })
})

describe('Compare into a new document: the result opened in the editor', () => {
  it('accept all gives the revised table and reject all the original, for every pair compared by column', async () => {
    const pairs: Array<[Grid, Grid]> = []
    for (const merges of Object.values(SHAPES)) {
      const original = baseGrid(merges)
      for (let at = 0; at <= 4; at++) {
        const { grid, reaches } = withColumn(original, at)
        if (!reaches) pairs.push([original, grid])
      }
      for (let at = 0; at < 4; at++) {
        const { grid, reaches } = withoutColumn(original, at)
        if (!reaches) pairs.push([original, grid])
      }
    }
    expect(pairs).toHaveLength(42)
    for (const [original, revised] of pairs) {
      const { bytes } = await compareDocuments(
        await parseDocx(await buildDocx({ bodyXml: around(original) })),
        await parseDocx(await buildDocx({ bodyXml: around(revised) })),
        { author: 'Rae Visor', date: '2026-11-24T00:00:00Z' },
      )
      expect(await settled(bytes, 'accept')).toEqual(gridShape(revised))
      expect(await settled(bytes, 'reject')).toEqual(gridShape(original))
    }
  }, 120_000)
})

describe('rows added or removed among cells merged down', () => {
  const variants = (original: Grid) => {
    const columns: Array<[string, Grid]> = [['the same columns', original]]
    for (let at = 0; at <= 4; at++) {
      const { grid, reaches } = withColumn(original, at)
      if (!reaches) columns.push([`a column in at ${at}`, grid])
    }
    for (let at = 0; at < 4; at++) {
      const { grid, reaches } = withoutColumn(original, at)
      if (!reaches) columns.push([`column ${at} out`, grid])
    }
    const out: Array<[string, Grid]> = []
    for (const [name, grid] of columns) {
      for (let at = 0; at <= 4; at++) out.push([`${name}, a row in at ${at}`, withRow(grid, at)])
      for (let at = 0; at < 4; at++) out.push([`${name}, row ${at} out`, withoutRow(grid, at)])
    }
    return out
  }
  const TALL = Object.entries(SHAPES).filter(([, merges]) => merges.some((m) => m.r1 > m.r0))

  // left alone and counted: paired rows merged differently (a row that started or ended a merge went)
  const SKIPPED: Record<string, number> = {
    'a cell down two rows of the first column': 22,
    'a cell down the whole last column': 16,
    'a block two columns wide and two rows high': 32,
    'a cell across and another down': 24,
  }

  for (const [name, merges] of TALL) {
    it(`into the open document, ${name}: both ways in the file and after a reopen, or left alone`, async () => {
      const original = baseGrid(merges)
      let skipped = 0
      for (const [change, revised] of variants(original)) {
        for (const [from, to] of [
          [original, revised],
          [revised, original],
        ] as const) {
          const result = await combine(from, to)
          if (result.summary.skipped > 0) {
            skipped++
            expect(collectRevisions(result.editor.state.doc), change).toEqual([])
            continue
          }
          expect(result.resolved('accept'), change).toEqual(gridShape(to))
          expect(result.resolved('reject'), change).toEqual(gridShape(from))
          expect(await settled(result.file.bytes, 'accept'), change).toEqual(gridShape(to))
          expect(await settled(result.file.bytes, 'reject'), change).toEqual(gridShape(from))
        }
      }
      expect(skipped).toBe(SKIPPED[name])
    }, 240_000)

    it(`into a new document, ${name}: the result opened in the editor resolves both ways`, async () => {
      const original = baseGrid(merges)
      for (const [change, revised] of variants(original)) {
        const { bytes } = await compareDocuments(
          await parseDocx(await buildDocx({ bodyXml: around(original) })),
          await parseDocx(await buildDocx({ bodyXml: around(revised) })),
          { author: 'Rae Visor', date: '2026-11-24T00:00:00Z' },
        )
        expect(await settled(bytes, 'accept'), change).toEqual(gridShape(revised))
        expect(await settled(bytes, 'reject'), change).toEqual(gridShape(original))
      }
    }, 240_000)
  }

  it('a row that comes in where a column with a cell merged down goes: the merge runs through it until the changes are resolved', async () => {
    const original = baseGrid([{ r0: 0, r1: 3, c0: 3, c1: 3 }])
    const revised = withRow(withoutColumn(original, 3).grid, 2)
    const { editor, file } = await expectBothWays(original, revised)
    const spans: number[] = []
    editor.state.doc.descendants((node) => {
      if (node.type.name.startsWith('docTableCell') && node.textContent === 'Origin')
        spans.push(node.attrs.rowspan as number)
    })
    expect(spans).toEqual([5])
    expect(file.xml.match(/<w:trPr><w:ins /g)).toHaveLength(1)
  })
})
