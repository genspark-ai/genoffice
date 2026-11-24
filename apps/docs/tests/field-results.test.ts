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
import { collectCaptions } from '../src/renderer/commands/reference-commands'
import { insertCrossReference, listRefTargets } from '../src/renderer/editor/cross-reference'
import {
  STYLEREF_LABEL,
  captionPreview,
  numberingContextOf,
  renumberAfterInsert,
  renumberSeqFields,
  settleFieldResults,
  textOfXml,
  unreliableSeqLabelsOf,
} from '../src/renderer/editor/field-results'
import { captionBody } from '../src/renderer/editor/index-tables'
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

  it('a reference in capitals to a caption follows its number in capitals', async () => {
    const { editor, parsed } = await open(
      CAPTIONS.replace(
        refTo('_Ref111', 'Figure 2 Two'),
        `<w:p>${field('REF _Ref111 \\* Upper \\h', 'FIGURE 2 TWO')}${field('REF _Ref111 \\* roman', 'Figure 2 Two')}</w:p>`,
      ),
    )
    insertCaption(editor, parsed, topPos(editor, 1), 'Figure', 'New')
    expect(refTexts(editor)).toEqual(['FIGURE 3 TWO', 'Figure 2 Two'])
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

  it('a count that restarts at a heading level stays out of the renumbering while the heading styles are not known', async () => {
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

describe('page fields saved without a page layout keep their mark', () => {
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

describe('captions numbered by chapter', () => {
  const chapter = (text: string, numbered = true) =>
    '<w:p><w:pPr><w:pStyle w:val="Heading1"/>' +
    (numbered ? '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr>' : '') +
    `</w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`
  const simple = (instr: string, result: string) =>
    `<w:fldSimple w:instr=" ${instr} "><w:r><w:rPr><w:noProof/></w:rPr><w:t>${result}</w:t></w:r></w:fldSimple>`
  /** a caption as Word's Caption dialog writes it with "Include chapter number" */
  const fromWord = (no: string, n: string, text: string) =>
    '<w:p><w:pPr><w:pStyle w:val="Caption"/></w:pPr><w:r><w:t xml:space="preserve">Figure </w:t></w:r>' +
    `${simple('STYLEREF 1 \\s', no)}<w:r><w:noBreakHyphen/></w:r>${simple('SEQ Figure \\* ARABIC \\s 1', n)}` +
    `<w:r><w:t xml:space="preserve"> ${text}</w:t></w:r></w:p>`
  const WORD_FILE =
    chapter('Alpha') +
    fromWord('1', '1', 'One') +
    fromWord('1', '2', 'Two') +
    chapter('Beta') +
    fromWord('2', '1', 'Three') +
    p('End')
  const BY_CHAPTER = { format: 'ARABIC', chapter: { style: 1, sep: 'hyphen' } } as const

  const NUMBERING =
    '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
    '<w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/>' +
    '<w:lvlText w:val="Chapter %1."/></w:lvl></w:abstractNum>' +
    '<w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num></w:numbering>'

  async function openNumbered(bodyXml: string) {
    const parsed = await parseDocx(await buildDocx({ bodyXml, numberingXml: NUMBERING }))
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks) as never,
    })
    editors.add(editor)
    editor.storage.listNumbering.styles = parsed.styles
    editor.storage.listNumbering.defs = parsed.numbering
    return { editor, parsed }
  }

  const unreliable = (editor: Editor, parsed: Parsed) => [
    ...unreliableSeqLabelsOf(parsed.blocks, editor.schema, numberingContextOf(editor)),
  ]

  function insertChaptered(editor: Editor, parsed: Parsed, at: number, text: string) {
    const xml = generateCaptionXml('Figure', 9, text, { chapter: BY_CHAPTER.chapter })
    const node = editor.schema.nodes.docProtected.create({
      docxIndex: null,
      blockType: 'passthrough',
      label: 'Caption',
      genXml: xml,
      fieldDisplay: { kind: 'text', left: textOfXml(xml) },
    })
    editor.view.dispatch(editor.state.tr.insert(at, node))
    renumberAfterInsert(editor, parsed.blocks)
  }

  const insertChapter = (editor: Editor, at: number, text: string) =>
    editor.commands.insertContentAt(at, {
      type: 'docListItem',
      attrs: { kind: 'ordered', numId: '2', ilvl: 0, styleId: 'Heading1' },
      content: [{ type: 'text', text }],
    })

  /** the results of the STYLEREF and SEQ fields of the saved body, caption by caption */
  const numbersOf = (xml: string): string[] =>
    [...xml.matchAll(/<w:p[ >].*?<\/w:p>/g)]
      .filter((m) => / SEQ /.test(m[0]))
      .map((m) =>
        textOfXml(m[0])
          .replace(/^Figure /, '')
          .replace(/ .*$/, ''),
      )

  it('a file from Word opens with its numbers and is reliable', async () => {
    const { editor, parsed } = await openNumbered(WORD_FILE)
    expect(unreliable(editor, parsed)).toEqual([])
    expect(shown(editor)).toEqual([
      'Figure 1\u20111 One',
      'Figure 1\u20112 Two',
      'Figure 2\u20111 Three',
    ])
    const { xml } = await save(editor, parsed)
    expect(xml).toContain(fromWord('1', '2', 'Two'))
    expect(xml).not.toContain('w:dirty')
  })

  it('a label judged before the heading styles arrived is judged again once they are there', async () => {
    const { editor, parsed } = await openNumbered(WORD_FILE)
    const styles = editor.storage.listNumbering.styles
    editor.storage.listNumbering.styles = undefined
    expect(unreliable(editor, parsed)).toContain('Figure')
    editor.storage.listNumbering.styles = styles
    expect(unreliable(editor, parsed)).toEqual([])
  })

  it('a caption put into a chapter takes its number there and moves the ones after it', async () => {
    const { editor, parsed } = await openNumbered(WORD_FILE)
    insertChaptered(editor, parsed, topPos(editor, 2), 'New')
    insertChaptered(editor, parsed, topPos(editor, 5), 'First of two')
    expect(shown(editor)).toEqual([
      'Figure 1\u20111 One',
      'Figure 1\u20112 New',
      'Figure 1\u20113 Two',
      'Figure 2\u20111 First of two',
      'Figure 2\u20112 Three',
    ])
    const { xml, bytes } = await save(editor, parsed)
    expect(xml).not.toContain('w:dirty')
    expect(numbersOf(xml)).toEqual(['1\u20111', '1\u20112', '1\u20113', '2\u20111', '2\u20112'])
    expect(xml).toContain(' STYLEREF 1 \\s </w:instrText>')
    expect(xml).toContain(' SEQ Figure \\* ARABIC \\s 1 </w:instrText>')
    const again = await parseDocx(bytes)
    expect(
      again.blocks.filter((b) => b.fieldDisplay?.kind === 'text').map((b) => b.fieldDisplay?.left),
    ).toEqual(shown(editor))
  })

  it('a chapter put in the middle renumbers the chapters and restarts the count on save', async () => {
    const { editor, parsed } = await openNumbered(WORD_FILE)
    insertChapter(editor, topPos(editor, 2), 'Inserted')
    const { xml } = await save(editor, parsed)
    expect(shown(editor)).toEqual([
      'Figure 1\u20111 One',
      'Figure 2\u20111 Two',
      'Figure 3\u20111 Three',
    ])
    expect(numbersOf(xml)).toEqual(['1\u20111', '2\u20111', '3\u20111'])
    expect(xml).not.toContain('w:dirty')
  })

  it('a heading without a number gives 0, and no heading the text Word leaves', async () => {
    const { editor, parsed } = await openNumbered(
      p('Intro') + chapter('Plain', false) + fromWord('0', '1', 'One') + p('End'),
    )
    expect(unreliable(editor, parsed)).toEqual([])
    insertChaptered(editor, parsed, topPos(editor, 1), 'Before')
    insertChaptered(editor, parsed, topPos(editor, 4), 'After')
    // before any heading the first one after it is taken, as STYLEREF does
    expect(shown(editor)).toEqual([
      'Figure 0\u20111 Before',
      'Figure 0\u20111 One',
      'Figure 0\u20112 After',
    ])
    const bare = await openNumbered(p('Intro') + p('End'))
    insertChaptered(bare.editor, bare.parsed, topPos(bare.editor, 1), 'Alone')
    expect(shown(bare.editor)).toEqual([
      'Figure Error! No text of specified style in document.\u20111 Alone',
    ])
  })

  it('results that are not what the headings give are left as the file has them', async () => {
    const { editor, parsed } = await openNumbered(
      chapter('Alpha') + fromWord('7', '1', 'One') + chapter('Beta') + fromWord('7', '2', 'Two'),
    )
    expect(unreliable(editor, parsed).sort()).toEqual([STYLEREF_LABEL, 'Figure'])
    insertChapter(editor, 0, 'First')
    const { xml } = await save(editor, parsed)
    expect(numbersOf(xml)).toEqual(['7\u20111', '7\u20112'])
  })

  it('the numbering set for a label goes to the captions it already has, and comes off again', async () => {
    const { editor, parsed } = await openNumbered(
      chapter('Alpha') +
        generateCaptionXml('Figure', 1, 'One') +
        generateCaptionXml('Table', 1, 'T') +
        chapter('Beta') +
        generateCaptionXml('Figure', 2, 'Two') +
        p('End'),
    )
    const xml = generateCaptionXml('Figure', 3, 'New', {
      format: 'ROMAN',
      chapter: BY_CHAPTER.chapter,
    })
    const attrs = {
      docxIndex: null,
      blockType: 'passthrough',
      label: 'Caption',
      genXml: xml,
      fieldDisplay: { kind: 'text', left: textOfXml(xml) },
    }
    const numbering = { label: 'Figure', format: 'ROMAN', chapter: BY_CHAPTER.chapter } as const
    editor.commands.setTextSelection(editor.state.doc.content.size - 1)
    expect(captionPreview(editor, parsed.blocks, attrs, numbering)).toBe('Figure 2\u2011II New')
    editor.view.dispatch(
      editor.state.tr.insert(
        editor.state.doc.content.size,
        editor.schema.nodes.docProtected.create(attrs),
      ),
    )
    renumberAfterInsert(editor, parsed.blocks, numbering)
    expect(shown(editor)).toEqual([
      'Figure 1\u2011I One',
      'Table 1 T',
      'Figure 2\u2011I Two',
      'Figure 2\u2011II New',
    ])
    // one undo step takes the caption and the new numbering back
    editor.commands.undo()
    expect(shown(editor)).toEqual(['Figure 1 One', 'Table 1 T', 'Figure 2 Two'])
    editor.commands.redo()

    const tr = editor.state.tr
    const plain = generateCaptionXml('Figure', 9, 'Last')
    tr.insert(
      tr.doc.content.size,
      editor.schema.nodes.docProtected.create({
        ...attrs,
        genXml: plain,
        fieldDisplay: { kind: 'text', left: textOfXml(plain) },
      }),
    )
    editor.view.dispatch(tr)
    renumberAfterInsert(editor, parsed.blocks, { label: 'Figure', format: 'ARABIC', chapter: null })
    expect(shown(editor)).toEqual([
      'Figure 1 One',
      'Table 1 T',
      'Figure 2 Two',
      'Figure 3 New',
      'Figure 4 Last',
    ])
    const saved = await save(editor, parsed)
    expect(saved.xml).not.toContain('STYLEREF')
    expect(saved.xml).not.toContain('w:dirty')
    expect(seqOf(saved.xml).map(([, n]) => n)).toEqual(['1', '1', '2', '3', '4'])
  })

  it('the caption list a table of figures and a cross-reference read carries the chapter number', async () => {
    const { editor, parsed } = await openNumbered(WORD_FILE)
    const captions = collectCaptions(editor, parsed.blocks)
    expect(captions.map((c) => c.text)).toEqual([
      'Figure 1\u20111 One',
      'Figure 1\u20112 Two',
      'Figure 2\u20111 Three',
    ])
    expect(captionBody(captions[2].text, 'Figure')).toBe('Three')
    expect(captionBody('Figure 2.IV: Sales', 'Figure')).toBe('Sales')
    expect(captionBody('Figure 2 Sales', 'Figure')).toBe('Sales')
  })

  it('a STYLEREF typed as a field follows the heading before it', async () => {
    const { editor, parsed } = await openNumbered(
      chapter('Alpha') + p('x') + chapter('Beta') + p('y'),
    )
    const mark = (instr: string) => editor.schema.marks.instrField.create({ instr, dirty: true })
    const tr = editor.state.tr
    const y = topPos(editor, 3) + 1
    tr.insert(y, editor.schema.text('?', [mark('STYLEREF 1 \\s')]))
    tr.insert(y, editor.schema.text('?', [mark('STYLEREF "heading 1"')]))
    tr.insert(topPos(editor, 1) + 1, editor.schema.text('?', [mark('STYLEREF 1 \\n')]))
    editor.view.dispatch(tr)
    settleFieldResults(editor, parsed.blocks)
    const texts = () =>
      collectFields(editor.state.doc).map((f) => editor.state.doc.textBetween(f.from, f.to, ''))
    expect(texts()).toEqual(['Chapter 1', 'Beta', '2'])
    expect(collectFields(editor.state.doc).every((f) => f.mark.attrs.dirty === false)).toBe(true)
    insertChapter(editor, 0, 'First')
    settleFieldResults(editor, parsed.blocks)
    expect(texts()).toEqual(['Chapter 2', 'Beta', '3'])
  })

  it('a capitals switch on a STYLEREF shapes the heading text; a number format on it stays for Word', async () => {
    const { editor, parsed } = await openNumbered(chapter('Alpha beta') + p('x'))
    const mark = (instr: string) => editor.schema.marks.instrField.create({ instr, dirty: true })
    const tr = editor.state.tr
    const x = topPos(editor, 1) + 1
    tr.insert(x, editor.schema.text('?', [mark('STYLEREF 1 \\* roman')]))
    tr.insert(x, editor.schema.text('?', [mark('STYLEREF 1 \\* Caps \\* MERGEFORMAT')]))
    tr.insert(x, editor.schema.text('?', [mark('STYLEREF 1 \\* Upper')]))
    editor.view.dispatch(tr)
    settleFieldResults(editor, parsed.blocks)
    const fields = collectFields(editor.state.doc)
    expect(fields.map((f) => editor.state.doc.textBetween(f.from, f.to, ''))).toEqual([
      'ALPHA BETA',
      'Alpha Beta',
      '?',
    ])
    expect(fields.map((f) => f.mark.attrs.dirty)).toEqual([false, false, true])
  })

  it('a STYLEREF shows the heading number in context and where the heading is, on Update Field', async () => {
    const { editor, parsed } = await openNumbered(
      chapter('Alpha') + p('x') + chapter('Beta') + p('y'),
    )
    const mark = (instr: string) => editor.schema.marks.instrField.create({ instr, dirty: true })
    const tr = editor.state.tr
    const y = topPos(editor, 3) + 1
    tr.insert(y, editor.schema.text('?', [mark('STYLEREF 1 \\p')]))
    tr.insert(y, editor.schema.text('?', [mark('STYLEREF 1 \\n \\p')]))
    tr.insert(y, editor.schema.text('?', [mark('STYLEREF 1 \\w')]))
    tr.insert(y, editor.schema.text('?', [mark('STYLEREF 1 \\r')]))
    // under the first heading
    tr.insert(topPos(editor, 1) + 1, editor.schema.text('?', [mark('STYLEREF 1 \\r \\p')]))
    editor.view.dispatch(tr)
    const texts = () =>
      collectFields(editor.state.doc).map((f) => editor.state.doc.textBetween(f.from, f.to, ''))
    // the save has no words for above and below: those fields wait for an update
    settleFieldResults(editor, parsed.blocks)
    expect(texts()).toEqual(['?', 'Chapter 2', 'Chapter 2', '?', '?'])
    const update = editor.state.tr
    renumberSeqFields(update, parsed.blocks, {
      headings: numberingContextOf(editor),
      dropDirty: true,
      words: { above: 'above', below: 'below' },
    })
    editor.view.dispatch(update)
    expect(texts()).toEqual([
      'Chapter 1 above',
      'Chapter 2',
      'Chapter 2',
      'Chapter 2 above',
      'above',
    ])
    expect(collectFields(editor.state.doc).every((f) => f.mark.attrs.dirty === false)).toBe(true)
  })

  it('a STYLEREF takes \\l as it stands in the text, and \\t drops the words around a number in context', async () => {
    const { editor, parsed } = await openNumbered(
      chapter('Alpha') + p('x') + chapter('Beta') + p('y'),
    )
    const mark = (instr: string) => editor.schema.marks.instrField.create({ instr, dirty: true })
    const tr = editor.state.tr
    const y = topPos(editor, 3) + 1
    tr.insert(y, editor.schema.text('?', [mark('STYLEREF 1 \\w \\t')]))
    tr.insert(y, editor.schema.text('?', [mark('STYLEREF 1 \\r \\t')]))
    tr.insert(y, editor.schema.text('?', [mark('STYLEREF 1 \\l \\n')]))
    tr.insert(y, editor.schema.text('?', [mark('STYLEREF 1 \\l')]))
    editor.view.dispatch(tr)
    settleFieldResults(editor, parsed.blocks)
    const fields = collectFields(editor.state.doc)
    expect(fields.map((f) => editor.state.doc.textBetween(f.from, f.to, ''))).toEqual([
      'Beta',
      'Chapter 2',
      '2',
      '2',
    ])
    expect(fields.every((f) => f.mark.attrs.dirty === false)).toBe(true)
  })

  it('Cross-reference lists a heading numbered through its style under Heading, and refers to its number', async () => {
    const { editor, parsed } = await openNumbered(
      chapter('Alpha') + chapter('Plain', false) + chapter('Beta') + p('See '),
    )
    const targets = listRefTargets(editor, parsed.blocks, { kind: 'heading' })
    expect(targets.map((t) => [t.text, t.refText, t.level])).toEqual([
      ['Chapter 1. Alpha', 'Alpha', 1],
      ['Plain', 'Plain', 1],
      ['Chapter 2. Beta', 'Beta', 1],
    ])
    const words = { above: 'above', below: 'below', onPage: (n: number) => `on page ${n}` }
    const refer = (target: number, content: 'paraNumber' | 'paraNumberNoContext' | 'text') => {
      editor.commands.setTextSelection(editor.state.doc.content.size - 1)
      const list = listRefTargets(editor, parsed.blocks, { kind: 'heading' })
      const request = {
        content,
        hyperlink: true,
        includeAboveBelow: false,
        heading: true,
        pages: null,
        words,
      }
      expect(
        insertCrossReference(editor, parsed.blocks, { target: list[target], ...request }),
      ).toBe(true)
    }
    refer(2, 'paraNumber')
    refer(1, 'paraNumberNoContext')
    refer(2, 'text')
    expect(refTexts(editor)).toEqual(['Chapter 2', '0', 'Beta'])
    // a chapter put in front moves the number the reference shows
    insertChapter(editor, 0, 'First')
    expect(refTexts(editor)).toEqual(['Chapter 3', '0', 'Beta'])
  })
})
