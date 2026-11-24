import { afterEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { parseDocx, readSections, saveDocx } from '@genoffice/docx-engine'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { editorExtensions } from '../src/renderer/editor/extensions'
import {
  blocksToPmDoc,
  pmDocOptions,
  pmDocToSavePlan,
  type PmNode,
} from '../src/renderer/editor/convert'
import { combineDocuments, type CombineOptions } from '../src/renderer/editor/combine'
import {
  acceptAllRevisions,
  applyRevisions,
  collectRevisions,
  rejectAllRevisions,
} from '../src/renderer/editor/revisions'

;(globalThis as { CSS?: unknown }).CSS ??= { escape: (s: string) => s }

const AUTHOR = 'Rae Visor'
const DATE = '2026-11-17T10:00:00Z'
const para = (text: string) => `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`
const cell = (text: string) =>
  `<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>${para(text)}</w:tc>`
const table = (rows: string[][]) =>
  `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid>${rows[0]!
    .map(() => '<w:gridCol w:w="2000"/>')
    .join(
      '',
    )}</w:tblGrid>${rows.map((r) => `<w:tr>${r.map(cell).join('')}</w:tr>`).join('')}</w:tbl>`

const editors: Editor[] = []
afterEach(() => {
  for (const editor of editors.splice(0)) editor.destroy()
})

async function load(bodyXml: string) {
  const parsed = await parseDocx(await buildDocx({ bodyXml }))
  const json = blocksToPmDoc(parsed.blocks, readSections(parsed), pmDocOptions(parsed))
  return { parsed, json }
}

async function combine(
  originalXml: string,
  revisedXml: string,
  options: Partial<CombineOptions> = {},
) {
  const original = await load(originalXml)
  const revised = await load(revisedXml)
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    editorProps: { handleScrollToSelection: () => true },
    content: original.json,
  })
  editors.push(editor)
  const before = editor.state.doc
  const summary = combineDocuments(editor, revised.json.content ?? [], {
    author: AUTHOR,
    date: DATE,
    ...options,
  })
  return { editor, summary, before, parsed: original.parsed }
}

/** the document as a reader sees it: paragraphs of [text, 'ins' | 'del' | ''] pieces */
function pieces(editor: Editor): Array<Array<[string, string]>> {
  const out: Array<Array<[string, string]>> = []
  editor.state.doc.descendants((node) => {
    if (!node.isTextblock) return true
    const row: Array<[string, string]> = []
    node.forEach((child) => {
      const kind = child.marks.find((m) => m.type.name === 'ins' || m.type.name === 'del')
      const last = row[row.length - 1]
      const name = kind?.type.name ?? ''
      if (last && last[1] === name) last[0] += child.text ?? ''
      else row.push([child.text ?? '', name])
    })
    out.push(row)
    return false
  })
  return out
}

const texts = (editor: Editor): string[] => {
  const out: string[] = []
  editor.state.doc.descendants((node) => {
    if (!node.isTextblock) return true
    out.push(node.textContent)
    return false
  })
  return out
}

const kinds = (editor: Editor) => collectRevisions(editor.state.doc).map((r) => r.kind)

describe('Combine Documents writes the differences as tracked changes', () => {
  it('a word only the revised document has is a tracked insertion', async () => {
    const { editor, summary } = await combine(
      para('The quick fox jumps.'),
      para('The quick brown fox jumps.'),
    )
    expect(pieces(editor)).toEqual([
      [
        ['The quick ', ''],
        ['brown ', 'ins'],
        ['fox jumps.', ''],
      ],
    ])
    expect(summary).toEqual({
      insertions: 1,
      deletions: 0,
      formatChanges: 0,
      skipped: 0,
      related: true,
    })
    const revisions = collectRevisions(editor.state.doc)
    expect(revisions).toHaveLength(1)
    expect(revisions[0]).toMatchObject({ kind: 'ins', author: AUTHOR, date: DATE })
  })

  it('a word only the open document has is a tracked deletion', async () => {
    const { editor, summary } = await combine(
      para('The quick brown fox jumps.'),
      para('The quick fox jumps.'),
    )
    expect(pieces(editor)).toEqual([
      [
        ['The quick ', ''],
        ['brown ', 'del'],
        ['fox jumps.', ''],
      ],
    ])
    expect(summary).toMatchObject({ insertions: 0, deletions: 1 })
    expect(collectRevisions(editor.state.doc)[0]).toMatchObject({ kind: 'del', author: AUTHOR })
  })

  it('a replaced word is struck and the new one follows it', async () => {
    const { editor } = await combine(
      para('Payment is due in thirty days.'),
      para('Payment is due in sixty days.'),
    )
    expect(pieces(editor)).toEqual([
      [
        ['Payment is due in ', ''],
        ['thirty', 'del'],
        ['sixty', 'ins'],
        [' days.', ''],
      ],
    ])
  })

  it('a paragraph that shares no words is deleted whole and its replacement inserted whole', async () => {
    const { editor, summary } = await combine(
      para('Opening line.') + para('Alpha beta gamma delta.') + para('Closing line.'),
      para('Opening line.') + para('Something else entirely here.') + para('Closing line.'),
    )
    expect(pieces(editor)).toEqual([
      [['Opening line.', '']],
      [['Alpha beta gamma delta.', 'del']],
      [['Something else entirely here.', 'ins']],
      [['Closing line.', '']],
    ])
    expect(summary).toMatchObject({ insertions: 1, deletions: 1 })
    expect(kinds(editor)).toEqual(['blockDel', 'blockIns'])
  })

  it('added and removed paragraphs resolve to the revised text on accept and the original on reject', async () => {
    const originalXml = para('One.') + para('Two goes away.') + para('Three.')
    const revisedXml = para('One.') + para('Three.') + para('Four is new.')
    const accepted = await combine(originalXml, revisedXml)
    expect(kinds(accepted.editor)).toEqual(['blockDel', 'blockIns'])
    acceptAllRevisions(accepted.editor)
    expect(texts(accepted.editor)).toEqual(['One.', 'Three.', 'Four is new.'])
    expect(collectRevisions(accepted.editor.state.doc)).toEqual([])

    const rejected = await combine(originalXml, revisedXml)
    rejectAllRevisions(rejected.editor)
    expect(texts(rejected.editor)).toEqual(['One.', 'Two goes away.', 'Three.'])
    expect(collectRevisions(rejected.editor.state.doc)).toEqual([])
  })

  it('each change is accepted or rejected on its own', async () => {
    const { editor } = await combine(
      para('Keep the red door and the old lamp.'),
      para('Keep the blue door and the lamp.'),
    )
    const revisions = collectRevisions(editor.state.doc)
    expect(revisions.map((r) => r.kind)).toEqual(['del', 'ins', 'del'])
    applyRevisions(editor, [revisions[2]!], 'reject')
    applyRevisions(editor, collectRevisions(editor.state.doc).slice(0, 2), 'accept')
    expect(texts(editor)).toEqual(['Keep the blue door and the old lamp.'])
  })

  it('a paragraph that is a pending insertion and absent from the revised document is struck on top of it', async () => {
    const pending =
      '<w:ins w:id="1" w:author="Ann" w:date="2026-01-01T00:00:00Z">' +
      para('Ann added this line.') +
      '</w:ins>'
    const originalXml = para('One.') + pending + para('Three.')
    const revisedXml = para('One.') + para('Three.')
    const { editor, summary } = await combine(originalXml, revisedXml)
    expect(summary).toEqual({
      insertions: 0,
      deletions: 1,
      formatChanges: 0,
      skipped: 0,
      related: true,
    })
    expect(editor.state.doc.child(1).attrs.blockRevision).toMatchObject({
      kind: 'ins',
      author: 'Ann',
    })
    expect(pieces(editor)[1]).toEqual([['Ann added this line.', 'del']])
    expect(collectRevisions(editor.state.doc).map((r) => [r.kind, r.author])).toEqual([
      ['blockIns', 'Ann'],
      ['del', AUTHOR],
    ])
    acceptAllRevisions(editor)
    expect(texts(editor)).toEqual(['One.', 'Three.'])
    expect(collectRevisions(editor.state.doc)).toEqual([])

    // rejecting only the combine's deletion leaves the insertion pending as it was
    const kept = await combine(originalXml, revisedXml)
    applyRevisions(
      kept.editor,
      collectRevisions(kept.editor.state.doc).filter((r) => r.author === AUTHOR),
      'reject',
    )
    expect(pieces(kept.editor)[1]).toEqual([['Ann added this line.', '']])
    expect(collectRevisions(kept.editor.state.doc).map((r) => r.kind)).toEqual(['blockIns'])

    const rejected = await combine(originalXml, revisedXml)
    rejectAllRevisions(rejected.editor)
    expect(texts(rejected.editor)).toEqual(['One.', 'Three.'])
    expect(collectRevisions(rejected.editor.state.doc)).toEqual([])
  })

  it('identical documents leave the document alone', async () => {
    const xml = para('Same here.') + table([['a', 'b']]) + para('And here.')
    const { editor, summary, before } = await combine(xml, xml)
    expect(summary).toEqual({
      insertions: 0,
      deletions: 0,
      formatChanges: 0,
      skipped: 0,
      related: true,
    })
    expect(editor.state.doc).toBe(before)
  })

  it('documents with no paragraph in common are left alone and reported', async () => {
    const { editor, summary, before } = await combine(
      para('Minutes of the board meeting.') + para('Attendance was complete.'),
      para('Recipe for lentil soup.') + para('Simmer gently until soft.'),
    )
    expect(summary.related).toBe(false)
    expect(editor.state.doc).toBe(before)
  })

  it('undo takes the whole combine back in one step', async () => {
    const { editor, before } = await combine(
      para('One.') + para('Two goes away.') + para('Three stays put.'),
      para('One.') + para('Three stays here.') + para('Four is new.'),
    )
    expect(editor.state.doc.eq(before)).toBe(false)
    editor.commands.undo()
    expect(editor.state.doc.eq(before)).toBe(true)
  })

  it('text in the cells of a table with the same shape is compared; a table with one more column takes the column as inserted cells', async () => {
    const same = await combine(
      para('Intro.') +
        table([
          ['Item', 'Price'],
          ['Tea', 'two pounds'],
        ]),
      para('Intro.') +
        table([
          ['Item', 'Price'],
          ['Tea', 'three pounds'],
        ]),
    )
    expect(pieces(same.editor)).toContainEqual([
      ['two', 'del'],
      ['three', 'ins'],
      [' pounds', ''],
    ])
    expect(same.summary).toEqual({
      insertions: 1,
      deletions: 1,
      formatChanges: 0,
      skipped: 0,
      related: true,
    })

    const reshaped = await combine(
      para('Intro.') +
        table([
          ['Item', 'Price'],
          ['Tea', 'two pounds'],
        ]),
      para('Intro.') +
        table([
          ['Item', 'Price', 'Stock'],
          ['Tea', 'two pounds', 'plenty'],
        ]),
    )
    expect(reshaped.summary).toEqual({
      insertions: 1,
      deletions: 0,
      formatChanges: 0,
      skipped: 0,
      related: true,
    })
    const cells: string[] = []
    reshaped.editor.state.doc.descendants((node) => {
      const kind = (node.attrs.cellRevision as { kind?: string } | null)?.kind
      if (node.type.name === 'docTableCell') cells.push(`${node.textContent}:${kind ?? ''}`)
    })
    expect(cells).toEqual(['Item:', 'Price:', 'Stock:ins', 'Tea:', 'two pounds:', 'plenty:ins'])
  })

  it('an inserted paragraph keeps its heading level and bold, and drops a style this document lacks', async () => {
    const heading =
      '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>New section</w:t></w:r></w:p>'
    const original = await load(para('Intro.') + para('Outro.'))
    const revised = await load(para('Intro.') + heading + para('Outro.'))
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: original.json,
    })
    editors.push(editor)
    combineDocuments(editor, revised.json.content ?? [], {
      author: AUTHOR,
      date: DATE,
      hasStyle: () => false,
    })
    const inserted = editor.state.doc.child(1)
    expect(inserted.textContent).toBe('New section')
    expect(inserted.attrs.styleId).toBeNull()
    expect(inserted.attrs.docxIndex).toBeNull()
    expect(inserted.firstChild!.marks.map((m) => m.type.name)).toEqual(
      expect.arrayContaining(['bold', 'ins']),
    )
  })

  it('the combined document saves as w:ins / w:del with the author and reads back', async () => {
    const { editor, parsed } = await combine(
      para('Payment is due in thirty days.') + para('This clause goes.') + para('Signed.'),
      para('Payment is due in sixty days.') + para('Signed.') + para('A new closing clause.'),
    )
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    const saved = await saveDocx(parsed, plan.saveBlocks)
    const reparsed = await parseDocx(saved)
    const xml = reparsed.internal.documentXml
    expect(xml).toMatch(
      new RegExp(
        `<w:del w:id="\\d+" w:author="${AUTHOR}" w:date="${DATE}"><w:r>.*?<w:delText[^>]*>thirty</w:delText>`,
      ),
    )
    expect(xml).toMatch(
      new RegExp(
        `<w:ins w:id="\\d+" w:author="${AUTHOR}" w:date="${DATE}"><w:r>.*?<w:t[^>]*>sixty</w:t>`,
      ),
    )
    expect(xml).toContain('>This clause goes.</w:delText>')
    expect(xml).toMatch(
      /<w:ins [^>]*><w:r>(?:<w:rPr>.*?<\/w:rPr>)?<w:t[^>]*>A new closing clause\.<\/w:t>/,
    )

    const reopened = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(reparsed.blocks, readSections(reparsed), pmDocOptions(reparsed)),
    })
    editors.push(reopened)
    const pending = collectRevisions(reopened.state.doc)
    expect(pending.length).toBeGreaterThanOrEqual(4)
    expect(pending.every((r) => r.author === AUTHOR)).toBe(true)
    acceptAllRevisions(reopened)
    expect(texts(reopened)).toEqual([
      'Payment is due in sixty days.',
      'Signed.',
      'A new closing clause.',
    ])
  })
})

describe('the comparison settings of the dialog switch what is written into the open document', () => {
  const ORIGINAL = para('The Seller  shall deliver.')
  const REVISED = para('The seller shall deliver.')

  it('case and white space differences are changes until they are switched off', async () => {
    const on = await combine(ORIGINAL, REVISED)
    expect(on.summary.insertions + on.summary.deletions).toBeGreaterThan(0)
    const noCase = await combine(ORIGINAL, REVISED, { ignore: { case: true } })
    expect(kinds(noCase.editor)).not.toHaveLength(0)
    const off = await combine(ORIGINAL, REVISED, { ignore: { case: true, whitespace: true } })
    expect(off.summary).toEqual({
      insertions: 0,
      deletions: 0,
      formatChanges: 0,
      skipped: 0,
      related: true,
    })
    expect(off.editor.state.doc.eq(off.before)).toBe(true)
  })

  it('character level marks the letters that differ, word level the whole word', async () => {
    const word = await combine(para('The colour is red.'), para('The color is red.'))
    expect(pieces(word.editor)).toEqual([
      [
        ['The ', ''],
        ['colour', 'del'],
        ['color', 'ins'],
        [' is red.', ''],
      ],
    ])
    const char = await combine(para('The colour is red.'), para('The color is red.'), {
      granularity: 'char',
    })
    expect(pieces(char.editor)).toEqual([
      [
        ['The colo', ''],
        ['u', 'del'],
        ['r is red.', ''],
      ],
    ])
  })

  it('with tables switched off a changed cell is left as it is and not counted', async () => {
    const original = para('Intro stays.') + table([['Price', 'ten dollars']])
    const revised = para('Intro stays.') + table([['Price', 'twelve dollars']])
    const on = await combine(original, revised)
    expect(kinds(on.editor)).toEqual(['del', 'ins'])
    const off = await combine(original, revised, { tables: false })
    expect(off.summary).toEqual({
      insertions: 0,
      deletions: 0,
      formatChanges: 0,
      skipped: 0,
      related: true,
    })
    expect(off.editor.state.doc.eq(off.before)).toBe(true)
  })

  it('shown in the revised document, what only it has is inserted and what only the original has comes in deleted', async () => {
    const revised =
      para('Payment is due in sixty days.') + para('Signed.') + para('A new closing clause.')
    const original =
      para('Payment is due in thirty days.') + para('This clause goes.') + para('Signed.')
    // the open document is the revised one; the other is the original
    const { editor, summary, parsed } = await combine(revised, original, { openIsRevised: true })
    expect(summary).toEqual({
      insertions: 2,
      deletions: 2,
      formatChanges: 0,
      skipped: 0,
      related: true,
    })
    expect(pieces(editor)[0]).toEqual([
      ['Payment is due in ', ''],
      ['sixty', 'ins'],
      ['thirty', 'del'],
      [' days.', ''],
    ])
    const saved = await saveDocx(
      parsed,
      pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks).saveBlocks,
    )
    const reparsed = await parseDocx(saved)
    const reopen = () => {
      const reopened = new Editor({
        element: document.createElement('div'),
        extensions: editorExtensions,
        content: blocksToPmDoc(reparsed.blocks, readSections(reparsed), pmDocOptions(reparsed)),
      })
      editors.push(reopened)
      return reopened
    }
    const accepted = reopen()
    acceptAllRevisions(accepted)
    expect(texts(accepted)).toEqual([
      'Payment is due in sixty days.',
      'Signed.',
      'A new closing clause.',
    ])
    const rejected = reopen()
    rejectAllRevisions(rejected)
    expect(texts(rejected)).toEqual([
      'Payment is due in thirty days.',
      'This clause goes.',
      'Signed.',
    ])
  })
})

describe('Combine keeps the tracked changes both documents already hold', () => {
  const ABE = 'w:author="Abe Origin" w:date="2026-11-01T08:00:00Z"'
  const BEA = 'w:author="Bea Revise" w:date="2026-11-02T09:00:00Z"'
  const run = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`
  const struck = (text: string) => `<w:r><w:delText xml:space="preserve">${text}</w:delText></w:r>`
  const OPEN =
    `<w:p>${run('The ')}<w:ins w:id="5" ${ABE}>${run('old ')}</w:ins>${run('cat sat on the mat.')}</w:p>` +
    para('Signed.')
  const OTHER =
    `<w:p>${run('The cat sat ')}<w:del w:id="3" ${BEA}>${struck('on the mat')}</w:del>` +
    `<w:ins w:id="4" ${BEA}>${run('by the door')}</w:ins>${run('. Twice.')}</w:p>` +
    para('Signed.') +
    `<w:p><w:pPr><w:rPr><w:ins w:id="6" ${BEA}/></w:rPr></w:pPr><w:ins w:id="7" ${BEA}>${run('A clause Bea proposed.')}</w:ins></w:p>`
  const changes = (editor: Editor) =>
    collectRevisions(editor.state.doc)
      .filter((r) => r.kind === 'ins' || r.kind === 'del')
      .map((r) => `${r.kind} ${r.author}: ${editor.state.doc.textBetween(r.from, r.to)}`)

  it("the other document's changes come in under their authors; only the plain difference is new", async () => {
    const { editor, summary } = await combine(OPEN, OTHER, { keepRevisions: true })
    expect(changes(editor)).toEqual([
      'ins Abe Origin: old ',
      'del Bea Revise: on the mat',
      'ins Bea Revise: by the door',
      `ins ${AUTHOR}:  Twice.`,
      'ins Bea Revise: A clause Bea proposed.',
    ])
    expect(summary).toMatchObject({ insertions: 1, deletions: 0, carried: 3 })
    const dated = collectRevisions(editor.state.doc).find((r) => r.author === 'Bea Revise')
    expect(dated?.date).toBe('2026-11-02T09:00:00Z')
  })

  it("without it the other document reads as accepted and everything is the label's", async () => {
    const { editor } = await combine(OPEN, OTHER)
    const authors = new Set(collectRevisions(editor.state.doc).map((r) => r.author))
    expect([...authors].sort()).toEqual(['Abe Origin', AUTHOR])
  })

  it('the result saves each change under its author, and accept all / reject all resolve it', async () => {
    const { editor, parsed } = await combine(OPEN, OTHER, { keepRevisions: true })
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    const reparsed = await parseDocx(await saveDocx(parsed, plan.saveBlocks))
    const xml = reparsed.internal.documentXml
    expect(xml).toMatch(new RegExp(`<w:ins w:id="\\d+" ${ABE}><w:r>.*?<w:t[^>]*>old </w:t>`))
    expect(xml).toMatch(
      new RegExp(`<w:del w:id="\\d+" ${BEA}><w:r>.*?<w:delText[^>]*>on the mat</w:delText>`),
    )
    expect(xml).toMatch(new RegExp(`<w:ins w:id="\\d+" ${BEA}><w:r>.*?<w:t[^>]*>by the door</w:t>`))
    expect(xml).toMatch(
      new RegExp(
        `<w:ins w:id="\\d+" w:author="${AUTHOR}" w:date="${DATE}"><w:r>.*?<w:t[^>]*> Twice\\.</w:t>`,
      ),
    )
    expect(xml).toMatch(new RegExp(`<w:pPr><w:rPr><w:ins w:id="\\d+" w:author="Bea Revise"`))
    const ids = [...xml.matchAll(/<w:(?:ins|del) w:id="(\d+)"/g)].map((m) => m[1])
    expect(new Set(ids).size).toBe(ids.length)

    const reopen = () => {
      const reopened = new Editor({
        element: document.createElement('div'),
        extensions: editorExtensions,
        content: blocksToPmDoc(reparsed.blocks, readSections(reparsed), pmDocOptions(reparsed)),
      })
      editors.push(reopened)
      return reopened
    }
    const accepted = reopen()
    acceptAllRevisions(accepted)
    expect(texts(accepted)).toEqual([
      'The old cat sat by the door. Twice.',
      'Signed.',
      'A clause Bea proposed.',
    ])
    const rejected = reopen()
    rejectAllRevisions(rejected)
    expect(texts(rejected).filter(Boolean)).toEqual(['The cat sat on the mat.', 'Signed.'])
  })

  it('a paragraph the open document inserted as a tracked change is left as it is', async () => {
    const { editor } = await combine(
      para('First clause.') +
        `<w:p><w:pPr><w:rPr><w:ins w:id="8" ${ABE}/></w:rPr></w:pPr><w:ins w:id="9" ${ABE}>${run('A clause Abe proposed.')}</w:ins></w:p>` +
        para('The last clause here.'),
      para('First clause.') + para('The last clause there.'),
      { keepRevisions: true },
    )
    expect(changes(editor)).toEqual([
      'ins Abe Origin: A clause Abe proposed.',
      `del ${AUTHOR}: here.`,
      `ins ${AUTHOR}: there.`,
    ])
    expect(texts(editor)).toContain('A clause Abe proposed.')
  })

  it('an empty paragraph that is only a tracked paragraph mark keeps its author on either side', async () => {
    const mark = (who: string) =>
      `<w:p><w:pPr><w:rPr><w:ins w:id="8" ${who}/></w:rPr></w:pPr></w:p>`
    const { editor, summary } = await combine(
      para('First clause.') + mark(ABE) + para('Last clause.'),
      para('First clause.') + para('Last clause.') + mark(BEA) + para('The end.'),
      { keepRevisions: true },
    )
    expect(summary.carried).toBe(1)
    const marks: string[] = []
    editor.state.doc.forEach((node) => {
      if (node.textContent !== '') return
      const ins = node.attrs.paraMarkIns
        ? JSON.parse(node.attrs.paraMarkIns as string).author
        : null
      marks.push(`${ins} ${node.attrs.paraMarkDel ? 'deleted' : 'kept'}`)
    })
    expect(marks).toEqual(['Abe Origin kept', 'Bea Revise kept'])
  })
})
