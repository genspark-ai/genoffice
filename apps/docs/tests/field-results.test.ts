import { Editor } from '@tiptap/core'
import JSZip from 'jszip'
import {
  formatSeqNumber,
  generateCaptionXml,
  parseDocx,
  saveDocx,
  type FieldDisplay,
} from '@genoffice/docx-engine'
import { afterEach, describe, expect, it } from 'vitest'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { allowDirtyFields } from '../../../packages/docx-engine/tests/helpers/dirty-fields'
import { executeOps } from '../src/renderer/ai/ops'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { collectFields, setFieldInstr } from '../src/renderer/editor/field-codes'
import {
  renumberAfterInsert,
  renumberSeqFields,
  settleFieldResults,
  unreliableSeqLabelsOf,
} from '../src/renderer/editor/field-results'
import { insertTableFormula, refreshTableFormulas } from '../src/renderer/editor/table-formula'
import { xmlOfNode } from '../src/renderer/editor/toc-update'

type Parsed = Awaited<ReturnType<typeof parseDocx>>

const p = (text: string) => `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`

const field = (instr: string, result: string, dirty = false) =>
  `<w:r><w:fldChar w:fldCharType="begin"${dirty ? ' w:dirty="true"' : ''}/></w:r>` +
  `<w:r><w:instrText xml:space="preserve"> ${instr} </w:instrText></w:r>` +
  '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
  `<w:r><w:t>${result}</w:t></w:r>` +
  '<w:r><w:fldChar w:fldCharType="end"/></w:r>'

/** a caption whose whole paragraph is a `_Ref` bookmark, as a cross-reference leaves it */
const bookmarked = (xml: string, name: string, id: number) =>
  xml
    .replace('</w:pPr>', `</w:pPr><w:bookmarkStart w:id="${id}" w:name="${name}"/>`)
    .replace(/<\/w:p>$/, `<w:bookmarkEnd w:id="${id}"/></w:p>`)

const refTo = (name: string, text: string) =>
  `<w:p><w:r><w:t xml:space="preserve">See </w:t></w:r>${field(`REF ${name} \\h`, text)}</w:p>`

const editors = new Set<Editor>()
afterEach(() => {
  for (const editor of editors) editor.destroy()
  editors.clear()
})

async function open(bodyXml: string) {
  const parsed = await parseDocx(await buildDocx({ bodyXml }))
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
  editors.add(editor)
  return { editor, parsed }
}

async function save(editor: Editor, parsed: Parsed) {
  settleFieldResults(editor, parsed.blocks)
  const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
  const bytes = await saveDocx(parsed, plan.saveBlocks)
  const xml = await (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string')
  return { bytes, xml }
}

/** each SEQ field of the saved body: its label, its result, whether it is marked */
function seqOf(xml: string): Array<[string, string, boolean]> {
  const re =
    /<w:fldChar w:fldCharType="begin"( w:dirty="true")?\/><\/w:r><w:r>(?:<w:rPr>.*?<\/w:rPr>)?<w:instrText[^>]*> SEQ (\S+)[^<]*<\/w:instrText><\/w:r><w:r>(?:<w:rPr>.*?<\/w:rPr>)?<w:fldChar w:fldCharType="separate"\/><\/w:r><w:r>(?:<w:rPr>.*?<\/w:rPr>)?<w:t>([^<]*)<\/w:t>/g
  return [...xml.matchAll(re)].map((m) => [m[2], m[3], m[1] !== undefined])
}

/** what the caption blocks show, in order */
function shown(editor: Editor): string[] {
  const out: string[] = []
  editor.state.doc.forEach((node) => {
    const display = node.attrs.fieldDisplay as FieldDisplay | null
    if (node.type.name === 'docProtected' && display?.kind === 'text') out.push(display.left ?? '')
  })
  return out
}

/** puts a caption in as the Caption dialog does: numbered as one more than the label has */
function insertCaption(editor: Editor, parsed: Parsed, at: number, label: string, text: string) {
  const count = shown(editor).filter((left) => left.startsWith(`${label} `)).length
  const dirty = unreliableSeqLabelsOf(parsed.blocks, editor.schema).has(label)
  const node = editor.schema.nodes.docProtected.create({
    docxIndex: null,
    blockType: 'passthrough',
    label: 'Caption',
    genXml: generateCaptionXml(label, count + 1, text, { dirty }),
    fieldDisplay: {
      kind: 'text',
      left: `${label} ${formatSeqNumber(count + 1, 'ARABIC')} ${text}`,
    },
  })
  editor.view.dispatch(editor.state.tr.insert(at, node))
  renumberAfterInsert(editor, parsed.blocks)
}

const topPos = (editor: Editor, index: number): number => {
  let pos = 0
  for (let i = 0; i < index; i++) pos += editor.state.doc.child(i).nodeSize
  return pos
}

const refTexts = (editor: Editor): string[] =>
  collectFields(editor.state.doc, editor.schema.marks.refField).map((r) =>
    editor.state.doc.textBetween(r.from, r.to, ''),
  )

const CAPTIONS =
  p('Intro') +
  generateCaptionXml('Figure', 1, 'One') +
  generateCaptionXml('Table', 1, 'T-one') +
  bookmarked(generateCaptionXml('Figure', 2, 'Two'), '_Ref111', 50) +
  refTo('_Ref111', 'Figure 2 Two') +
  generateCaptionXml('Figure', 3, 'Three')

describe('caption numbers are saved in document order, without w:dirty', () => {
  it('a caption put between two others renumbers the ones after it, and the reference to one', async () => {
    const { editor, parsed } = await open(CAPTIONS)
    insertCaption(editor, parsed, topPos(editor, 1), 'Figure', 'New')
    expect(shown(editor)).toEqual([
      'Figure 1 New',
      'Figure 2 One',
      'Table 1 T-one',
      'Figure 3 Two',
      'Figure 4 Three',
    ])
    expect(refTexts(editor)).toEqual(['Figure 3 Two'])
    // the caption list other features read follows the new numbers
    const two = editor.state.doc.child(4)
    expect(xmlOfNode(two as never, parsed.blocks)).toContain('<w:t>3</w:t>')

    const { xml, bytes } = await save(editor, parsed)
    expect(xml).not.toContain('w:dirty')
    expect(seqOf(xml)).toEqual([
      ['Figure', '1', false],
      ['Figure', '2', false],
      ['Table', '1', false],
      ['Figure', '3', false],
      ['Figure', '4', false],
    ])
    expect(xml).toMatch(/ REF _Ref111 \\h <\/w:instrText>.*?<w:t[^>]*>Figure 3 Two<\/w:t>/)
    expect(xml).toMatch(/w:name="_Ref111"\/>.*?Figure .*?<w:t>3<\/w:t>.*? Two/)

    const again = await parseDocx(bytes)
    expect(
      again.blocks.filter((b) => b.fieldDisplay?.kind === 'text').map((b) => b.fieldDisplay?.left),
    ).toEqual(['Figure 1 New', 'Figure 2 One', 'Table 1 T-one', 'Figure 3 Two', 'Figure 4 Three'])
    expect(unreliableSeqLabelsOf(again.blocks, editor.schema).size).toBe(0)
  })

  it('undo takes the caption and the renumbering back in one step', async () => {
    const { editor, parsed } = await open(CAPTIONS)
    const before = shown(editor)
    insertCaption(editor, parsed, topPos(editor, 1), 'Figure', 'New')
    editor.commands.undo()
    expect(shown(editor)).toEqual(before)
    expect(refTexts(editor)).toEqual(['Figure 2 Two'])
  })

  it('a deleted caption closes the gap on save', async () => {
    const { editor, parsed } = await open(CAPTIONS)
    const first = topPos(editor, 1)
    editor.view.dispatch(editor.state.tr.delete(first, first + editor.state.doc.child(1).nodeSize))
    const { xml } = await save(editor, parsed)
    expect(shown(editor)).toEqual(['Table 1 T-one', 'Figure 1 Two', 'Figure 2 Three'])
    expect(refTexts(editor)).toEqual(['Figure 1 Two'])
    expect(seqOf(xml)).toEqual([
      ['Table', '1', false],
      ['Figure', '1', false],
      ['Figure', '2', false],
    ])
    expect(xml).not.toContain('w:dirty')
    // nothing left to settle: a second save changes nothing
    const tr = editor.state.tr
    expect(renumberSeqFields(tr, parsed.blocks)).toBe(0)
    expect(tr.docChanged).toBe(false)
  })

  it('two labels interleaved keep a count each', async () => {
    const { editor, parsed } = await open(CAPTIONS)
    insertCaption(editor, parsed, topPos(editor, 1), 'Table', 'T-new')
    insertCaption(editor, parsed, editor.state.doc.content.size, 'Equation', 'E')
    expect(shown(editor)).toEqual([
      'Table 1 T-new',
      'Figure 1 One',
      'Table 2 T-one',
      'Figure 2 Two',
      'Figure 3 Three',
      'Equation 1 E',
    ])
    expect(refTexts(editor)).toEqual(['Figure 2 Two'])
    const { xml } = await save(editor, parsed)
    expect(seqOf(xml).map(([label, n]) => `${label} ${n}`)).toEqual([
      'Table 1',
      'Figure 1',
      'Table 2',
      'Figure 2',
      'Figure 3',
      'Equation 1',
    ])
    expect(xml).not.toContain('w:dirty')
  })

  it('a restart switch, a repeat and a hidden field are counted as Word counts them', async () => {
    const seqP = (instr: string, result: string) =>
      `<w:p><w:r><w:t xml:space="preserve">Figure </w:t></w:r>${field(instr, result)}</w:p>`
    const { editor, parsed } = await open(
      p('Intro') +
        seqP('SEQ Figure \\* ARABIC', '1') +
        seqP('SEQ Figure \\r 5', '5') +
        seqP('SEQ Figure \\c', '5') +
        `<w:p>${field('SEQ Figure \\h', '')}</w:p>` +
        seqP('SEQ Figure \\* ROMAN', 'VII'),
    )
    expect(unreliableSeqLabelsOf(parsed.blocks, editor.schema).size).toBe(0)
    insertCaption(editor, parsed, topPos(editor, 1), 'Figure', 'New')
    const { xml } = await save(editor, parsed)
    expect(seqOf(xml).map(([, n]) => n)).toEqual(['1', '2', '5', '5', '', 'VII'])
    // one more after the restart moves the numbers behind it only
    insertCaption(editor, parsed, topPos(editor, 4), 'Figure', 'Mid')
    expect(seqOf((await save(editor, parsed)).xml).map(([, n]) => n)).toEqual([
      '1',
      '2',
      '5',
      '6',
      '6',
      '',
      'VIII',
    ])
  })

  it('text typed into a caption stays when its number changes', async () => {
    const { editor, parsed } = await open(CAPTIONS)
    const pos = topPos(editor, 1)
    const node = editor.state.doc.child(1)
    editor.view.dispatch(
      editor.state.tr.setNodeMarkup(pos, undefined, {
        ...node.attrs,
        fieldDisplay: {
          ...(node.attrs.fieldDisplay as FieldDisplay),
          left: 'Figure 1 One, edited',
        },
      }),
    )
    insertCaption(editor, parsed, topPos(editor, 1), 'Figure', 'New')
    expect(shown(editor).slice(0, 2)).toEqual(['Figure 1 New', 'Figure 2 One, edited'])
    const { bytes, xml } = await save(editor, parsed)
    expect(seqOf(xml).map(([, n]) => n)).toEqual(['1', '2', '1', '3', '4'])
    const again = await parseDocx(bytes)
    expect(again.blocks[2].fieldDisplay?.left).toBe('Figure 2 One, edited')
  })

  it('SEQ fields the assistant inserts are numbered with the captions of the file', async () => {
    const { editor, parsed } = await open(
      p('Figure ') + generateCaptionXml('Figure', 1, 'From the file') + p('Figure '),
    )
    const outcome = executeOps(editor, [
      { op: 'insertField', target: { blockIndexes: [0] }, type: 'SEQ', args: 'Figure \\* ARABIC' },
      { op: 'insertField', target: { blockIndexes: [2] }, type: 'SEQ', args: 'Figure \\* ARABIC' },
    ])
    expect(outcome.ok, outcome.error).toBe(true)
    const { xml } = await save(editor, parsed)
    expect(xml).not.toContain('w:dirty')
    expect(xml.match(/SEQ Figure[^<]*<\/w:instrText>.*?<w:t[^>]*>(\d+)<\/w:t>/g)).toHaveLength(3)
    expect(
      [
        ...xml.matchAll(/SEQ Figure[^<]*<\/w:instrText>(?:(?!<w:t[ >]).)*<w:t[^>]*>(\d+)<\/w:t>/g),
      ].map((m) => m[1]),
    ).toEqual(['1', '2', '3'])
    expect(shown(editor)).toEqual(['Figure 2 From the file'])
  })
})

describe('numbering this editor does not reproduce is left as it came', () => {
  const STALE =
    p('Intro') + generateCaptionXml('Figure', 1, 'One') + generateCaptionXml('Figure', 3, 'Three')

  it('stale numbers stay, and a new caption of that label is left for Word to number', async () => {
    const { editor, parsed } = await open(STALE)
    expect([...unreliableSeqLabelsOf(parsed.blocks, editor.schema)]).toEqual(['Figure'])
    const untouched = await save(editor, parsed)
    expect(seqOf(untouched.xml).map(([, n]) => n)).toEqual(['1', '3'])
    insertCaption(editor, parsed, topPos(editor, 1), 'Figure', 'New')
    insertCaption(editor, parsed, topPos(editor, 1), 'Table', 'T')
    allowDirtyFields()
    const { xml } = await save(editor, parsed)
    expect(seqOf(xml)).toEqual([
      ['Table', '1', false],
      ['Figure', '3', true],
      ['Figure', '1', false],
      ['Figure', '3', false],
    ])
  })

  it('a switch that restarts at a heading level keeps the label out of the renumbering', async () => {
    const { editor, parsed } = await open(
      `<w:p><w:r><w:t xml:space="preserve">Figure </w:t></w:r>${field('SEQ Figure \\* ARABIC \\s 1', '1')}</w:p>` +
        `<w:p><w:r><w:t xml:space="preserve">Figure </w:t></w:r>${field('SEQ Figure \\* ARABIC \\s 1', '1')}</w:p>`,
    )
    expect([...unreliableSeqLabelsOf(parsed.blocks, editor.schema)]).toEqual(['Figure'])
    const tr = editor.state.tr
    expect(renumberSeqFields(tr, parsed.blocks)).toBe(0)
  })

  it('captions inside a table cell are counted, not rewritten', async () => {
    const table =
      '<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/></w:tblPr><w:tblGrid><w:gridCol w:w="6000"/></w:tblGrid>' +
      `<w:tr><w:tc>${generateCaptionXml('Figure', 1, 'In a cell')}</w:tc></w:tr></w:tbl>`
    const { editor, parsed } = await open(table + generateCaptionXml('Figure', 2, 'Below'))
    expect(unreliableSeqLabelsOf(parsed.blocks, editor.schema).size).toBe(0)
    insertCaption(editor, parsed, editor.state.doc.content.size, 'Figure', 'Last')
    expect(shown(editor)).toEqual(['Figure 2 Below', 'Figure 3 Last'])
  })
})

describe('a field that arrives with w:dirty', () => {
  const ARRIVED =
    p('Intro') +
    generateCaptionXml('Figure', 1, 'One', { dirty: true }) +
    '<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/></w:tblPr><w:tblGrid><w:gridCol w:w="6000"/></w:tblGrid>' +
    `<w:tr><w:tc>${p('2')}</w:tc></w:tr><w:tr><w:tc>${p('3')}</w:tc></w:tr>` +
    `<w:tr><w:tc><w:p>${field('=SUM(ABOVE)', '5', true)}</w:p></w:tc></w:tr></w:tbl>` +
    `<w:p><w:r><w:t xml:space="preserve">By </w:t></w:r>${field('AUTHOR', 'Ada', true)}</w:p>`

  it('keeps the mark while it is untouched', async () => {
    const { editor, parsed } = await open(ARRIVED)
    allowDirtyFields()
    const { xml } = await save(editor, parsed)
    expect(xml.match(/w:dirty="true"/g)).toHaveLength(3)
    // nothing was rewritten: the body is the file's own
    const original = await (
      await JSZip.loadAsync(parsed.internal.originalBytes)
    )
      .file('word/document.xml')!
      .async('string')
    expect(xml.slice(xml.indexOf('<w:body>'))).toBe(original.slice(original.indexOf('<w:body>')))
  })

  it('loses it once the field is recomputed, and only that field', async () => {
    const { editor, parsed } = await open(ARRIVED)
    // Update Field: the caption number and the formula are computed here, AUTHOR is not
    const tr = editor.state.tr
    refreshTableFormulas(tr)
    renumberSeqFields(tr, parsed.blocks, { dropDirty: true })
    editor.view.dispatch(tr)
    allowDirtyFields()
    const { xml } = await save(editor, parsed)
    expect(seqOf(xml)).toEqual([['Figure', '1', false]])
    expect(xml).toMatch(
      /<w:fldChar w:fldCharType="begin"\/><\/w:r><w:r>[^/]*<w:instrText[^>]*> =SUM\(ABOVE\) <\/w:instrText>.*?<w:t[^>]*>5<\/w:t>/,
    )
    expect(xml).toMatch(/w:dirty="true"\/><\/w:r><w:r>[^/]*<w:instrText[^>]*> AUTHOR /)
    expect(xml.match(/w:dirty="true"/g)).toHaveLength(1)
  })

  it('a caption renumbered by an insert before it loses the mark with its old number', async () => {
    const { editor, parsed } = await open(ARRIVED)
    insertCaption(editor, parsed, topPos(editor, 1), 'Figure', 'New')
    allowDirtyFields()
    const { xml } = await save(editor, parsed)
    expect(seqOf(xml)).toEqual([
      ['Figure', '1', false],
      ['Figure', '2', false],
    ])
  })
})

describe('formula fields carry the computed result', () => {
  it('in body text and after its code is edited', async () => {
    const { editor, parsed } = await open(p('Sum: '))
    editor.commands.setTextSelection(editor.state.doc.content.size - 1)
    expect(insertTableFormula(editor, '=(2+3)*4')).toBeNull()
    let saved = await save(editor, parsed)
    expect(saved.xml).not.toContain('w:dirty')
    expect(saved.xml).toMatch(/ =\(2\+3\)\*4 <\/w:instrText>.*?<w:t[^>]*>20<\/w:t>/)

    // Edit Field has no result of its own for a formula: the evaluator gives it one
    const [range] = collectFields(editor.state.doc)
    expect(setFieldInstr(editor, range, '=6*7', '')).toBe(true)
    expect(collectFields(editor.state.doc)[0].mark.attrs.dirty).toBe(true)
    const tr = editor.state.tr
    refreshTableFormulas(tr)
    editor.view.dispatch(tr)
    saved = await save(editor, parsed)
    expect(saved.xml).not.toContain('w:dirty')
    expect(saved.xml).toMatch(/ =6\*7 <\/w:instrText>.*?<w:t[^>]*>42<\/w:t>/)
  })

  it('an edited code this editor cannot compute keeps its old result and the mark', async () => {
    const { editor, parsed } = await open(
      `<w:p><w:r><w:t xml:space="preserve">By </w:t></w:r>${field('AUTHOR', 'Ada')}</w:p>`,
    )
    const [range] = collectFields(editor.state.doc)
    expect(setFieldInstr(editor, range, 'LASTSAVEDBY', '')).toBe(true)
    allowDirtyFields()
    const { xml } = await save(editor, parsed)
    expect(xml).toMatch(/w:dirty="true"\/><\/w:r><w:r>[^/]*<w:instrText[^>]*> LASTSAVEDBY /)
    expect(xml).toMatch(/<w:t[^>]*>Ada<\/w:t>/)
  })
})

describe('page fields keep the mark they had before', () => {
  it('PAGE, NUMPAGES and PAGEREF the assistant inserts are left for Word to compute', async () => {
    const { editor, parsed } = await open(
      p('Target paragraph.') + p('Back to ') + p('Page ') + p('Of '),
    )
    const outcome = executeOps(editor, [
      { op: 'insertBookmark', target: { blockIndexes: [0] }, name: 'Target' },
      { op: 'insertField', target: { blockIndexes: [1] }, type: 'PAGEREF', args: 'Target \\h' },
      { op: 'insertField', target: { blockIndexes: [2] }, type: 'PAGE' },
      { op: 'insertField', target: { blockIndexes: [3] }, type: 'NUMPAGES' },
    ])
    expect(outcome.ok, outcome.error).toBe(true)
    allowDirtyFields()
    const { xml } = await save(editor, parsed)
    expect(
      [
        ...xml.matchAll(
          /(w:dirty="true")?\/><\/w:r><w:r><w:instrText[^>]*> (PAGEREF|PAGE|NUMPAGES)[^<]*<\/w:instrText>/g,
        ),
      ].map((m) => `${m[2]}${m[1] ? ' dirty' : ''}`),
    ).toEqual(['PAGEREF dirty', 'PAGE dirty', 'NUMPAGES dirty'])
  })

  it('a page field whose code is edited stays marked, with the result given', async () => {
    const { editor } = await open(
      `<w:p><w:r><w:t xml:space="preserve">Page </w:t></w:r>${field('NUMPAGES', '3')}</w:p>`,
    )
    const [range] = collectFields(editor.state.doc)
    expect(setFieldInstr(editor, range, 'PAGE', '2')).toBe(true)
    expect(collectFields(editor.state.doc)[0].mark.attrs).toMatchObject({
      instr: 'PAGE',
      dirty: true,
    })
    expect(editor.state.doc.textContent).toBe('Page 2')
  })
})

describe('a table of contents built before the pages were known', () => {
  it('is saved with the pages the layout filled in, unmarked; without them it stays marked', async () => {
    const heading = (text: string) =>
      `<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`
    const { editor, parsed } = await open(p('Top') + heading('Alpha') + heading('Beta'))
    const outcome = executeOps(editor, [{ op: 'insertToc', afterBlockIndex: 0 }])
    expect(outcome.ok, outcome.error).toBe(true)
    allowDirtyFields()
    const unpaged = await save(editor, parsed)
    expect(unpaged.xml).toMatch(/w:dirty="true"\/><\/w:r><w:r><w:instrText[^>]*> TOC /)

    // what the pagination pass does once it has laid the headings out
    const tr = editor.state.tr
    let page = 2
    editor.state.doc.forEach((node, pos) => {
      const display = node.attrs.fieldDisplay as FieldDisplay | null
      if (display?.kind === 'tocLine')
        tr.setNodeMarkup(pos, undefined, {
          ...node.attrs,
          fieldDisplay: { ...display, right: String(page++) },
        })
    })
    editor.view.dispatch(tr)
    const paged = await save(editor, parsed)
    expect(paged.xml).not.toContain('w:dirty')
    expect(paged.xml).toMatch(
      /Alpha<\/w:t><\/w:r><w:r>(?:<w:rPr>.*?<\/w:rPr>)?<w:tab\/><\/w:r><w:r>.*?<w:t>2<\/w:t>/,
    )
    expect(paged.xml).toMatch(
      /Beta<\/w:t><\/w:r><w:r>(?:<w:rPr>.*?<\/w:rPr>)?<w:tab\/><\/w:r><w:r>.*?<w:t>3<\/w:t>/,
    )
  })
})
