import { afterEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { executeOps } from '../src/renderer/ai/ops'
import { setModuleLang } from '../src/renderer/i18n/locale'

interface JsonNode {
  type: string
  attrs?: Record<string, unknown>
  content?: JsonNode[]
  text?: string
  marks?: Array<{ type: string; attrs?: Record<string, unknown> }>
}

const text = (t: string, marks?: JsonNode['marks']): JsonNode => ({
  type: 'text',
  text: t,
  ...(marks && marks.length > 0 ? { marks } : {}),
})

const heading = (
  content: JsonNode[],
  level = 1,
  attrs: Record<string, unknown> = {},
): JsonNode => ({
  type: 'docHeading',
  attrs: { docxIndex: null, level, ...attrs },
  content,
})

const para = (content: JsonNode[], attrs: Record<string, unknown> = {}): JsonNode => ({
  type: 'docParagraph',
  attrs: { docxIndex: null, ...attrs },
  content,
})

const listItem = (content: JsonNode[], attrs: Record<string, unknown> = {}): JsonNode => ({
  type: 'docListItem',
  attrs: { docxIndex: null, kind: 'bullet', ...attrs },
  content,
})

const protectedTable = (attrs: Record<string, unknown> = {}): JsonNode => ({
  type: 'docProtected',
  attrs: {
    docxIndex: 90,
    blockType: 'table',
    label: 'Table 2×2',
    previewText: 'City GDP',
    ...attrs,
  },
})

const editors = new Set<Editor>()

afterEach(() => {
  for (const editor of editors) editor.destroy()
  editors.clear()
})

function createEditor(content: JsonNode[]): Editor {
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: { type: 'doc', content },
  })
  editors.add(editor)
  return editor
}

/** standard fixture: 0 h1 | 1 p | 2 h2 | 3 p | 4 li | 5 protected table */
function fixtureDoc(): JsonNode[] {
  return [
    heading(
      [
        text('Chapter 1 Overview', [
          { type: 'bold' },
          { type: 'docTextStyle', attrs: { sizeHalfPoints: 32 } },
        ]),
      ],
      1,
      {
        docxIndex: 0,
      },
    ),
    para([text('GenSpark intro,'), text('GenSpark is great', [{ type: 'bold' }])], {
      docxIndex: 1,
    }),
    heading([text('Risk Notes')], 2, { docxIndex: 2 }),
    para([text('Body paragraph')], { docxIndex: 3, align: 'center' }),
    listItem([text('List item')], { docxIndex: 4, numId: '1' }),
    protectedTable({ docxIndex: 5 }),
  ]
}

function textStyleOf(
  editor: Editor,
  blockIndex: number,
  childIndex = 0,
): Record<string, unknown> | null {
  const block = editor.state.doc.child(blockIndex)
  const child = block.child(childIndex)
  const mark = child.marks.find((m) => m.type.name === 'docTextStyle')
  return mark ? { ...mark.attrs } : null
}

setModuleLang('en')

describe('setFont', () => {
  it('sets color on all headings via nodeType, keeping bold and size (fields mask)', () => {
    const editor = createEditor(fixtureDoc())
    const outcome = executeOps(editor, [
      { op: 'setFont', target: { nodeType: 'docHeading' }, color: 'FF0000' },
    ])
    expect(outcome.ok).toBe(true)
    expect(outcome.results[0]).toMatchObject({ matched: 2, changed: 2, skippedProtected: 0 })

    const h1 = editor.state.doc.child(0).child(0)
    expect(h1.marks.some((m) => m.type.name === 'bold')).toBe(true)
    expect(textStyleOf(editor, 0)).toMatchObject({ color: 'FF0000', sizeHalfPoints: 32 })
    expect(textStyleOf(editor, 2)).toMatchObject({ color: 'FF0000' })
    // non-heading blocks untouched
    expect(textStyleOf(editor, 1)).toBeNull()
  })

  it('null clears one attr, keeping the others; mark removed when all attrs empty', () => {
    const editor = createEditor([
      para([
        text('red text', [
          { type: 'docTextStyle', attrs: { color: 'FF0000', sizeHalfPoints: 24 } },
        ]),
      ]),
      para([text('color only', [{ type: 'docTextStyle', attrs: { color: '00FF00' } }])]),
    ])
    const outcome = executeOps(editor, [
      { op: 'setFont', target: { nodeType: 'docParagraph' }, color: null },
    ])
    expect(outcome.ok).toBe(true)
    expect(textStyleOf(editor, 0)).toMatchObject({ color: null, sizeHalfPoints: 24 })
    expect(textStyleOf(editor, 1)).toBeNull()
  })

  it('font routes to its script slot: CJK keeps the Latin font, Latin keeps the CJK font', () => {
    const mixed = () =>
      createEditor([
        para([
          text('\u5408\u540c Contract', [
            { type: 'docTextStyle', attrs: { font: 'SimSun', fontAscii: 'Times New Roman' } },
          ]),
        ]),
      ])
    const style = (fontFamily: string | null) => [
      { op: 'setFont', target: { nodeType: 'docParagraph' as const }, fontFamily },
    ]

    const cjk = mixed()
    expect(executeOps(cjk, style('KaiTi')).ok).toBe(true)
    expect(textStyleOf(cjk, 0)).toMatchObject({ font: 'KaiTi', fontAscii: 'Times New Roman' })

    const latin = mixed()
    expect(executeOps(latin, style('Arial')).ok).toBe(true)
    expect(textStyleOf(latin, 0)).toMatchObject({ font: 'SimSun', fontAscii: 'Arial' })

    const cleared = mixed()
    expect(executeOps(cleared, style(null)).ok).toBe(true)
    expect(textStyleOf(cleared, 0)).toBeNull()
  })

  it('boolean fields add and remove basic marks', () => {
    const editor = createEditor([para([text('plain'), text('bold', [{ type: 'bold' }])])])
    executeOps(editor, [
      { op: 'setFont', target: { nodeType: 'docParagraph' }, bold: false, italic: true },
    ])
    const block = editor.state.doc.child(0)
    block.forEach((child) => {
      expect(child.marks.some((m) => m.type.name === 'bold')).toBe(false)
      expect(child.marks.some((m) => m.type.name === 'italic')).toBe(true)
    })
  })
})

describe('setFont google-parity fields', () => {
  it('baselineOffset SUPERSCRIPT sets vertAlign, NONE clears it, other attrs survive', () => {
    const editor = createEditor([
      para([text('x2', [{ type: 'docTextStyle', attrs: { color: 'FF0000' } }])]),
    ])
    executeOps(editor, [
      { op: 'setFont', target: { nodeType: 'docParagraph' }, baseline: 'superscript' },
    ])
    expect(textStyleOf(editor, 0)).toMatchObject({ vertAlign: 'superscript', color: 'FF0000' })

    executeOps(editor, [{ op: 'setFont', target: { nodeType: 'docParagraph' }, baseline: 'none' }])
    expect(textStyleOf(editor, 0)).toMatchObject({ vertAlign: null, color: 'FF0000' })
  })

  it('link adds and removes the link mark', () => {
    const editor = createEditor([para([text('official site')])])
    executeOps(editor, [
      { op: 'setFont', target: { nodeType: 'docParagraph' }, link: { url: 'https://example.com' } },
    ])
    const child = editor.state.doc.child(0).child(0)
    expect(child.marks.find((m) => m.type.name === 'link')?.attrs.href).toBe('https://example.com')

    executeOps(editor, [{ op: 'setFont', target: { nodeType: 'docParagraph' }, link: null }])
    expect(
      editor.state.doc
        .child(0)
        .child(0)
        .marks.some((m) => m.type.name === 'link'),
    ).toBe(false)
  })

  it('rejects a bad baselineOffset value', () => {
    const editor = createEditor([para([text('x')])])
    const outcome = executeOps(editor, [
      { op: 'setFont', target: { nodeType: 'docParagraph' }, baseline: 'UP' },
    ])
    expect(outcome.ok).toBe(false)
    expect(outcome.error).toContain('baseline')
  })
})

describe('updateParagraphStyle google-parity fields', () => {
  it('sets first-line indent and borders without touching other attrs', () => {
    const editor = createEditor(fixtureDoc())
    const outcome = executeOps(editor, [
      {
        op: 'setParagraphFormat',
        target: { nodeType: 'docParagraph' },
        indentFirstLine: 440,
        borders: 'b',
      },
    ])
    expect(outcome.ok).toBe(true)
    const node = editor.state.doc.child(3)
    expect(node.attrs.indentFirstLine).toBe(440)
    expect(node.attrs.borders).toBe('b')
    expect(node.attrs.align).toBe('center')
  })
})

describe('updateParagraphStyle', () => {
  it('sets lineSpacing without touching align (fields mask)', () => {
    const editor = createEditor(fixtureDoc())
    const outcome = executeOps(editor, [
      { op: 'setParagraphFormat', target: { nodeType: 'docParagraph' }, lineSpacing: 1.5 },
    ])
    expect(outcome.ok).toBe(true)
    expect(outcome.results[0].changed).toBe(2)
    const centered = editor.state.doc.child(3)
    expect(centered.attrs.lineSpacing).toBe(1.5)
    expect(centered.attrs.align).toBe('center')
    expect(centered.attrs.aiChanged).toBe(true)
    // headings not matched
    expect(editor.state.doc.child(0).attrs.lineSpacing).toBeNull()
  })
})

describe('setHeadingLevel', () => {
  it('demotes a matched heading level and clears styleId', () => {
    const editor = createEditor(fixtureDoc())
    const outcome = executeOps(editor, [
      {
        op: 'setHeadingLevel',
        target: { nodeType: 'docHeading', headingLevel: 2, containsText: 'Risk Notes' },
        level: 3,
      },
    ])
    expect(outcome.ok).toBe(true)
    const node = editor.state.doc.child(2)
    expect(node.type.name).toBe('docHeading')
    expect(node.attrs.level).toBe(3)
    expect(node.attrs.styleId).toBeNull()
    // the other heading keeps its level
    expect(editor.state.doc.child(0).attrs.level).toBe(1)
  })

  it('level 0 converts a heading to a plain paragraph', () => {
    const editor = createEditor(fixtureDoc())
    executeOps(editor, [
      { op: 'setHeadingLevel', target: { containsText: 'Risk Notes' }, level: 0 },
    ])
    const node = editor.state.doc.child(2)
    expect(node.type.name).toBe('docParagraph')
    expect(node.textContent).toBe('Risk Notes')
  })

  it('promotes a paragraph to a heading', () => {
    const editor = createEditor(fixtureDoc())
    executeOps(editor, [{ op: 'setHeadingLevel', target: { blockIndexes: [3] }, level: 2 }])
    const node = editor.state.doc.child(3)
    expect(node.type.name).toBe('docHeading')
    expect(node.attrs.level).toBe(2)
  })
})

describe('findReplace', () => {
  it('replaces every occurrence and keeps marks', () => {
    const editor = createEditor(fixtureDoc())
    const outcome = executeOps(editor, [
      { op: 'findReplace', find: 'GenSpark', replace: 'Genspark' },
    ])
    expect(outcome.ok).toBe(true)
    expect(outcome.results[0].detail).toBe('Replaced 2 occurrence(s)')
    const block = editor.state.doc.child(1)
    expect(block.textContent).toBe('Genspark intro,Genspark is great')
    const boldChild = block.child(block.childCount - 1)
    expect(boldChild.marks.some((m) => m.type.name === 'bold')).toBe(true)
  })

  it('matchCase: sensitive by default, insensitive when false', () => {
    const editor = createEditor(fixtureDoc())
    const strict = executeOps(editor, [{ op: 'findReplace', find: 'genspark', replace: 'X' }])
    expect(strict.results[0].changed).toBe(0)
    expect(editor.state.doc.child(1).textContent).toContain('GenSpark')

    const loose = executeOps(editor, [
      { op: 'findReplace', find: 'genspark', replace: 'X', matchCase: false },
    ])
    expect(loose.results[0].detail).toBe('Replaced 2 occurrence(s)')
    expect(editor.state.doc.child(1).textContent).toBe('X intro,X is great')
  })
})

describe('updateMatchedTextStyle', () => {
  const boldRanges = (editor: Editor, blockIndex: number): string[] => {
    const out: string[] = []
    editor.state.doc.child(blockIndex).forEach((child) => {
      if (child.isText && child.marks.some((m) => m.type.name === 'bold')) out.push(child.text!)
    })
    return out
  }

  it('styles only the matched text, not the whole block', () => {
    const editor = createEditor([
      para([text('a TODO: first and TODO: second item')]),
      para([text('no marker here')]),
    ])
    const outcome = executeOps(editor, [{ op: 'setMatchedFont', text: 'TODO:', bold: true }])
    expect(outcome.ok).toBe(true)
    expect(outcome.results[0].changed).toBe(1)
    expect(outcome.results[0].detail).toBe('2')
    expect(boldRanges(editor, 0)).toEqual(['TODO:', 'TODO:'])
    expect(boldRanges(editor, 1)).toEqual([])
    expect(editor.state.doc.child(0).attrs.aiChanged).toBe(true)
    expect(editor.state.doc.child(1).attrs.aiChanged).toBe(false)
  })

  it('merges docTextStyle attrs per match without wiping existing ones', () => {
    const editor = createEditor([
      para([
        text('keep GenSpark styled', [{ type: 'docTextStyle', attrs: { sizeHalfPoints: 32 } }]),
      ]),
    ])
    executeOps(editor, [{ op: 'setMatchedFont', text: 'GenSpark', color: 'FF0000' }])
    let styled: Record<string, unknown> | null = null
    editor.state.doc.child(0).forEach((child) => {
      if (child.isText && child.text === 'GenSpark') {
        styled = child.marks.find((m) => m.type.name === 'docTextStyle')?.attrs ?? null
      }
    })
    expect(styled).toMatchObject({ color: 'FF0000', sizeHalfPoints: 32 })
  })

  it('rejects invalid baselineOffset and malformed link like updateTextStyle does', () => {
    const editor = createEditor(fixtureDoc())
    const badBaseline = executeOps(editor, [
      { op: 'setMatchedFont', text: 'GenSpark', baseline: 'MIDDLE' },
    ] as never)
    expect(badBaseline.ok).toBe(false)
    expect(badBaseline.error).toContain('baseline')
    const badLink = executeOps(editor, [
      { op: 'setMatchedFont', text: 'GenSpark', link: 'https://example.com' },
    ] as never)
    expect(badLink.ok).toBe(false)
    expect(badLink.error).toContain('link')
  })

  it('rejects an empty needle and unknown fields', () => {
    const editor = createEditor(fixtureDoc())
    const empty = executeOps(editor, [{ op: 'setMatchedFont', text: '', bold: true }] as never)
    expect(empty.ok).toBe(false)
    const badField = executeOps(editor, [
      { op: 'setMatchedFont', text: 'x', bolder: true },
    ] as never)
    expect(badField.ok).toBe(false)
    expect(badField.error).toContain('unknown field')
  })
})

describe('deleteBlocks', () => {
  it('deletes multiple blocks without index drift', () => {
    const editor = createEditor(fixtureDoc())
    const outcome = executeOps(editor, [{ op: 'deleteBlocks', target: { blockIndexes: [1, 3] } }])
    expect(outcome.ok).toBe(true)
    expect(editor.state.doc.childCount).toBe(4)
    expect(editor.state.doc.child(0).textContent).toBe('Chapter 1 Overview')
    expect(editor.state.doc.child(1).textContent).toBe('Risk Notes')
    expect(editor.state.doc.child(2).textContent).toBe('List item')
    expect(editor.state.doc.child(3).type.name).toBe('docProtected')
  })

  it('deleting every block leaves one empty paragraph', () => {
    const editor = createEditor([para([text('a')]), para([text('b')])])
    executeOps(editor, [{ op: 'deleteBlocks', target: { blockIndexes: [0, 1] } }])
    expect(editor.state.doc.childCount).toBe(1)
    expect(editor.state.doc.child(0).textContent).toBe('')
  })
})

describe('moveBlocks', () => {
  it('moves blocks preserving relative order and docxIndex', () => {
    const editor = createEditor(fixtureDoc())
    const outcome = executeOps(editor, [
      { op: 'moveBlocks', blockIndexes: [3, 4], afterBlockIndex: 0 },
    ])
    expect(outcome.ok).toBe(true)
    const texts = [] as string[]
    editor.state.doc.forEach((n) => texts.push(n.textContent))
    expect(editor.state.doc.child(1).textContent).toBe('Body paragraph')
    expect(editor.state.doc.child(2).textContent).toBe('List item')
    expect(editor.state.doc.child(1).attrs.docxIndex).toBe(3)
    expect(editor.state.doc.child(2).attrs.docxIndex).toBe(4)
    // untouched anchors intact
    expect(editor.state.doc.child(0).attrs.docxIndex).toBe(0)
    expect(editor.state.doc.child(5).attrs.docxIndex).toBe(5)
  })

  it('afterBlockIndex -1 moves to the document start', () => {
    const editor = createEditor(fixtureDoc())
    executeOps(editor, [{ op: 'moveBlocks', blockIndexes: [2], afterBlockIndex: -1 }])
    expect(editor.state.doc.child(0).textContent).toBe('Risk Notes')
  })

  it('rejects the whole envelope when afterBlockIndex is a moved block', () => {
    const editor = createEditor(fixtureDoc())
    const before = JSON.stringify(editor.getJSON())
    const outcome = executeOps(editor, [
      { op: 'moveBlocks', blockIndexes: [1], afterBlockIndex: 1 },
    ])
    expect(outcome.ok).toBe(false)
    expect(JSON.stringify(editor.getJSON())).toBe(before)
  })
})

describe('protected blocks', () => {
  it('style commands skip docProtected and count it', () => {
    const editor = createEditor(fixtureDoc())
    const before = JSON.stringify(editor.getJSON())
    const outcome = executeOps(editor, [
      { op: 'setFont', target: { blockIndexes: [5] }, color: 'FF0000' },
    ])
    expect(outcome.ok).toBe(true)
    expect(outcome.results[0]).toMatchObject({ matched: 1, changed: 0, skippedProtected: 1 })
    expect(JSON.stringify(editor.getJSON())).toBe(before)
    expect(outcome.summary).toContain('protected block')
  })

  it('deleteBlocks does delete docProtected', () => {
    const editor = createEditor(fixtureDoc())
    executeOps(editor, [{ op: 'deleteBlocks', target: { blockIndexes: [5] } }])
    expect(editor.state.doc.childCount).toBe(5)
    editor.state.doc.forEach((n) => expect(n.type.name).not.toBe('docProtected'))
  })
})

describe('envelope validation', () => {
  it('an unknown op rejects the whole batch with no doc change', () => {
    const editor = createEditor(fixtureDoc())
    const before = JSON.stringify(editor.getJSON())
    const outcome = executeOps(editor, [
      { op: 'findReplace', find: 'GenSpark', replace: 'X' },
      { op: 'insertTable', rows: 2 },
    ])
    expect(outcome.ok).toBe(false)
    expect(outcome.error).toContain('unknown op')
    expect(outcome.error).toContain('Supported ops')
    expect(JSON.stringify(editor.getJSON())).toBe(before)
  })

  it('a setFont op without any font field is rejected with its usage line', () => {
    const editor = createEditor(fixtureDoc())
    const outcome = executeOps(editor, [{ op: 'setFont', target: { nodeType: 'docHeading' } }])
    expect(outcome.ok).toBe(false)
    expect(outcome.error).toContain('at least one font field')
    expect(outcome.error).toContain('usage: { op: "setFont"')
  })

  it('empty target rejects', () => {
    const editor = createEditor(fixtureDoc())
    const outcome = executeOps(editor, [{ op: 'deleteBlocks', target: {} }])
    expect(outcome.ok).toBe(false)
    expect(outcome.error).toContain('requires at least one condition')
  })
})

describe('transaction atomicity and aiChanged', () => {
  it('a multi-command round undoes in a single step', () => {
    const editor = createEditor(fixtureDoc())
    const before = JSON.stringify(editor.getJSON())
    const outcome = executeOps(editor, [
      { op: 'setFont', target: { nodeType: 'docHeading' }, color: 'FF0000' },
      { op: 'findReplace', find: 'GenSpark', replace: 'Genspark' },
      { op: 'deleteBlocks', target: { blockIndexes: [3] } },
    ])
    expect(outcome.ok).toBe(true)
    expect(JSON.stringify(editor.getJSON())).not.toBe(before)
    editor.commands.undo()
    expect(JSON.stringify(editor.getJSON())).toBe(before)
  })

  it('marks changed nodes aiChanged and leaves unmatched attrs untouched', () => {
    const editor = createEditor(fixtureDoc())
    const untouchedBefore = { ...editor.state.doc.child(3).attrs }
    executeOps(editor, [{ op: 'setFont', target: { nodeType: 'docHeading' }, color: 'FF0000' }])
    expect(editor.state.doc.child(0).attrs.aiChanged).toBe(true)
    expect(editor.state.doc.child(2).attrs.aiChanged).toBe(true)
    expect({ ...editor.state.doc.child(3).attrs }).toEqual(untouchedBefore)
    expect(editor.state.doc.child(1).attrs.aiChanged).toBe(false)
  })

  it('selection scope only touches blocks covered by the selection', () => {
    const editor = createEditor(fixtureDoc())
    // select inside block 1 (positions: block 0 occupies [0, nodeSize))
    const block0Size = editor.state.doc.child(0).nodeSize
    editor.commands.setTextSelection({ from: block0Size + 2, to: block0Size + 4 })
    const outcome = executeOps(editor, [
      {
        op: 'setParagraphFormat',
        target: { nodeType: 'docParagraph', scope: 'selection' },
        align: 'right',
      },
    ])
    expect(outcome.results[0].changed).toBe(1)
    expect(editor.state.doc.child(1).attrs.align).toBe('right')
    expect(editor.state.doc.child(3).attrs.align).toBe('center')
  })

  it('a frozen selection in the context wins over the live selection', () => {
    const editor = createEditor(fixtureDoc())
    // live selection sits in block 3; the frozen scope (captured at send time) says block 1
    const block3Pos =
      editor.state.doc.child(0).nodeSize +
      editor.state.doc.child(1).nodeSize +
      editor.state.doc.child(2).nodeSize
    editor.commands.setTextSelection({ from: block3Pos + 2, to: block3Pos + 4 })
    const outcome = executeOps(
      editor,
      [
        {
          op: 'setParagraphFormat',
          target: { nodeType: 'docParagraph', scope: 'selection' },
          align: 'right',
        },
      ],
      { selection: { startIndex: 1, endIndex: 1 } },
    )
    expect(outcome.results[0].changed).toBe(1)
    expect(editor.state.doc.child(1).attrs.align).toBe('right')
    expect(editor.state.doc.child(3).attrs.align).toBe('center')
  })

  it('replaceAllText with a target only replaces inside the targeted blocks', () => {
    const editor = createEditor([
      para([text('alpha one')]),
      para([text('alpha two')]),
      para([text('alpha three')]),
    ])
    const outcome = executeOps(editor, [
      { op: 'findReplace', find: 'alpha', replace: 'beta', target: { blockIndexes: [1] } },
    ])
    expect(outcome.ok).toBe(true)
    expect(outcome.results[0].changed).toBe(1)
    expect(editor.state.doc.child(0).textContent).toBe('alpha one')
    expect(editor.state.doc.child(1).textContent).toBe('beta two')
    expect(editor.state.doc.child(2).textContent).toBe('alpha three')
  })

  it('replaceAllText scoped to the frozen selection leaves other blocks alone', () => {
    const editor = createEditor([para([text('alpha one')]), para([text('alpha two')])])
    const outcome = executeOps(
      editor,
      [{ op: 'findReplace', find: 'alpha', replace: 'beta', target: { scope: 'selection' } }],
      { selection: { startIndex: 0, endIndex: 0 } },
    )
    expect(outcome.ok).toBe(true)
    expect(editor.state.doc.child(0).textContent).toBe('beta one')
    expect(editor.state.doc.child(1).textContent).toBe('alpha two')
  })

  it('frozen selection indexes beyond the document clamp to the last block', () => {
    const editor = createEditor(fixtureDoc())
    const outcome = executeOps(
      editor,
      [
        {
          op: 'setParagraphFormat',
          target: { nodeType: 'docListItem', scope: 'selection' },
          align: 'right',
        },
      ],
      { selection: { startIndex: 4, endIndex: 99 } },
    )
    expect(outcome.ok).toBe(true)
    expect(outcome.results[0].changed).toBe(1)
    expect(editor.state.doc.child(4).attrs.align).toBe('right')
  })
})

describe('createParagraphBullets / deleteParagraphBullets', () => {
  const NUM_IDS = { numIds: { bullet: '10', ordered: '20' } }

  it('converts paragraphs to bullet list items with the context numId', () => {
    const editor = createEditor(fixtureDoc())
    const outcome = executeOps(
      editor,
      [{ op: 'setList', target: { blockIndexes: [1, 3] } }],
      NUM_IDS,
    )
    expect(outcome.ok).toBe(true)
    expect(outcome.results[0].changed).toBe(2)
    for (const i of [1, 3]) {
      const node = editor.state.doc.child(i)
      expect(node.type.name).toBe('docListItem')
      expect(node.attrs.kind).toBe('bullet')
      expect(node.attrs.numId).toBe('10')
      expect(node.attrs.ilvl).toBe(0)
      expect(node.attrs.aiChanged).toBe(true)
    }
  })

  it('NUMBERED preset creates ordered items and switches existing list kind', () => {
    const editor = createEditor(fixtureDoc())
    executeOps(
      editor,
      [{ op: 'setList', target: { blockIndexes: [3, 4] }, kind: 'number' }],
      NUM_IDS,
    )
    expect(editor.state.doc.child(3).attrs.kind).toBe('ordered')
    expect(editor.state.doc.child(3).attrs.numId).toBe('20')
    // block 4 was already a bullet list item: kind switched
    expect(editor.state.doc.child(4).attrs.kind).toBe('ordered')
  })

  it('headings are matched but never converted', () => {
    const editor = createEditor(fixtureDoc())
    const outcome = executeOps(editor, [{ op: 'setList', target: { blockIndexes: [0] } }], NUM_IDS)
    expect(outcome.results[0]).toMatchObject({ matched: 1, changed: 0 })
    expect(editor.state.doc.child(0).type.name).toBe('docHeading')
  })

  it('deleteParagraphBullets converts list items back to paragraphs, keeping text', () => {
    const editor = createEditor(fixtureDoc())
    const outcome = executeOps(editor, [{ op: 'clearList', target: { nodeType: 'docListItem' } }])
    expect(outcome.results[0].changed).toBe(1)
    const node = editor.state.doc.child(4)
    expect(node.type.name).toBe('docParagraph')
    expect(node.textContent).toBe('List item')
    expect(node.attrs.aiChanged).toBe(true)
  })

  it('round trip undoes in one step', () => {
    const editor = createEditor(fixtureDoc())
    const before = JSON.stringify(editor.getJSON())
    executeOps(editor, [{ op: 'setList', target: { blockIndexes: [1, 3] } }], NUM_IDS)
    editor.commands.undo()
    expect(JSON.stringify(editor.getJSON())).toBe(before)
  })
})

describe('setImageProperties', () => {
  const imageDoc = () => [
    para([text('before')]),
    {
      type: 'docProtected',
      attrs: {
        docxIndex: 7,
        blockType: 'image',
        label: 'Image',
        imageWidthPx: 400,
        imageHeightPx: 200,
      },
    } as JsonNode,
    protectedTable({ docxIndex: 8 }),
  ]

  it('sets align on image blocks via nodeType image', () => {
    const editor = createEditor(imageDoc())
    const outcome = executeOps(editor, [
      { op: 'setImageProperties', target: { nodeType: 'image' }, align: 'center' },
    ])
    expect(outcome.ok).toBe(true)
    expect(outcome.results[0]).toMatchObject({ matched: 1, changed: 1 })
    expect(editor.state.doc.child(1).attrs.imageAlign).toBe('center')
    // table untouched
    expect(editor.state.doc.child(2).attrs.blockType).toBe('table')
  })

  it('scales the other dimension proportionally when only widthPx is given', () => {
    const editor = createEditor(imageDoc())
    executeOps(editor, [{ op: 'setImageProperties', target: { nodeType: 'image' }, widthPx: 200 }])
    const node = editor.state.doc.child(1)
    expect(node.attrs.imageWidthPx).toBe(200)
    expect(node.attrs.imageHeightPx).toBe(100)
  })

  it('nodeType image never matches non-image blocks', () => {
    const editor = createEditor(imageDoc())
    const outcome = executeOps(editor, [{ op: 'deleteBlocks', target: { nodeType: 'image' } }])
    expect(outcome.results[0].matched).toBe(1)
    expect(editor.state.doc.childCount).toBe(2)
    expect(editor.state.doc.child(0).textContent).toBe('before')
  })

  it('rejects unknown fields', () => {
    const editor = createEditor(imageDoc())
    const outcome = executeOps(editor, [
      { op: 'setImageProperties', target: { nodeType: 'image' }, align: 'center', rotation: 90 },
    ])
    expect(outcome.ok).toBe(false)
    expect(outcome.error).toContain('unknown field')
  })
})

describe('partial selection (character-precise scope)', () => {
  // block 1 = 'GenSpark intro,' + bold 'GenSpark is great'; content starts at pos 21
  const BLOCK1_CONTENT = 21
  const styleOfText = (editor: Editor, blockIndex: number, needle: string) => {
    const block = editor.state.doc.child(blockIndex)
    let found: Record<string, unknown> | null | undefined
    block.forEach((child) => {
      if (child.isText && child.text === needle) {
        const mark = child.marks.find((m) => m.type.name === 'docTextStyle')
        found = mark ? { ...mark.attrs } : null
      }
    })
    return found
  }

  it('updateTextStyle with scope selection styles only the selected characters', () => {
    const editor = createEditor(fixtureDoc())
    // 'is great' = offsets 24..32 of the block text
    editor.commands.setTextSelection({ from: BLOCK1_CONTENT + 24, to: BLOCK1_CONTENT + 32 })
    const outcome = executeOps(editor, [
      { op: 'setFont', target: { scope: 'selection' }, color: 'FF0000' },
    ])
    expect(outcome.ok).toBe(true)
    expect(outcome.results[0]).toMatchObject({ matched: 1, changed: 1 })
    const block = editor.state.doc.child(1)
    expect(block.childCount).toBe(3)
    expect(block.child(0).text).toBe('GenSpark intro,')
    expect(block.child(1).text).toBe('GenSpark ')
    expect(styleOfText(editor, 1, 'GenSpark intro,')).toBeNull()
    expect(styleOfText(editor, 1, 'GenSpark ')).toBeNull()
    expect(styleOfText(editor, 1, 'is great')).toMatchObject({ color: 'FF0000' })
    // the bold of the enclosing run survives on the styled slice
    expect(block.child(2).marks.some((m) => m.type.name === 'bold')).toBe(true)
    expect(block.attrs.aiChanged).toBe(true)
  })

  it('boolean marks on a partial selection split the run at the selection edges', () => {
    const editor = createEditor(fixtureDoc())
    // 'intro' = offsets 9..14
    editor.commands.setTextSelection({ from: BLOCK1_CONTENT + 9, to: BLOCK1_CONTENT + 14 })
    executeOps(editor, [{ op: 'setFont', target: { scope: 'selection' }, bold: true }])
    const block = editor.state.doc.child(1)
    expect(block.childCount).toBe(4)
    expect(block.child(1).text).toBe('intro')
    expect(block.child(1).marks.some((m) => m.type.name === 'bold')).toBe(true)
    expect(block.child(0).marks.some((m) => m.type.name === 'bold')).toBe(false)
    expect(block.child(2).text).toBe(',')
  })

  it('a caret (no range) keeps the whole-block behaviour', () => {
    const editor = createEditor(fixtureDoc())
    editor.commands.setTextSelection(BLOCK1_CONTENT + 9)
    executeOps(editor, [{ op: 'setFont', target: { scope: 'selection' }, color: 'FF0000' }])
    const block = editor.state.doc.child(1)
    expect(block.childCount).toBe(2)
    block.forEach((child) => {
      expect(child.marks.find((m) => m.type.name === 'docTextStyle')?.attrs).toMatchObject({
        color: 'FF0000',
      })
    })
  })

  it('paragraph-level commands ignore the character clip and format the whole block', () => {
    const editor = createEditor(fixtureDoc())
    editor.commands.setTextSelection({ from: BLOCK1_CONTENT + 24, to: BLOCK1_CONTENT + 32 })
    const outcome = executeOps(editor, [
      { op: 'setParagraphFormat', target: { scope: 'selection' }, align: 'right' },
    ])
    expect(outcome.results[0].changed).toBe(1)
    expect(editor.state.doc.child(1).attrs.align).toBe('right')
    expect(editor.state.doc.child(1).textContent).toBe('GenSpark intro,GenSpark is great')
  })

  it('a frozen selection with positions stays character-precise even after the live selection moved', () => {
    const editor = createEditor(fixtureDoc())
    const block3Pos =
      editor.state.doc.child(0).nodeSize +
      editor.state.doc.child(1).nodeSize +
      editor.state.doc.child(2).nodeSize
    editor.commands.setTextSelection(block3Pos + 2)
    executeOps(editor, [{ op: 'setFont', target: { scope: 'selection' }, color: 'FF0000' }], {
      selection: {
        startIndex: 1,
        endIndex: 1,
        from: BLOCK1_CONTENT + 24,
        to: BLOCK1_CONTENT + 32,
      },
    })
    expect(styleOfText(editor, 1, 'is great')).toMatchObject({ color: 'FF0000' })
    expect(styleOfText(editor, 1, 'GenSpark intro,')).toBeNull()
    expect(styleOfText(editor, 3, 'Body paragraph')).toBeNull()
  })

  it('a frozen block-only selection (no positions) still styles whole blocks', () => {
    const editor = createEditor(fixtureDoc())
    executeOps(editor, [{ op: 'setFont', target: { scope: 'selection' }, color: 'FF0000' }], {
      selection: { startIndex: 1, endIndex: 1 },
    })
    expect(styleOfText(editor, 1, 'GenSpark intro,')).toMatchObject({ color: 'FF0000' })
    expect(styleOfText(editor, 1, 'GenSpark is great')).toMatchObject({ color: 'FF0000' })
  })

  it('replaceAllText scoped to a partial selection skips matches outside the selected span', () => {
    const editor = createEditor(fixtureDoc())
    // select the second run 'GenSpark is great' (offsets 15..32)
    editor.commands.setTextSelection({ from: BLOCK1_CONTENT + 15, to: BLOCK1_CONTENT + 32 })
    const outcome = executeOps(editor, [
      { op: 'findReplace', find: 'GenSpark', replace: 'Acme', target: { scope: 'selection' } },
    ])
    expect(outcome.ok).toBe(true)
    expect(editor.state.doc.child(1).textContent).toBe('GenSpark intro,Acme is great')
  })

  it('updateMatchedTextStyle scoped to a partial selection styles only matches inside the span', () => {
    const editor = createEditor(fixtureDoc())
    editor.commands.setTextSelection({ from: BLOCK1_CONTENT + 15, to: BLOCK1_CONTENT + 32 })
    executeOps(editor, [
      { op: 'setMatchedFont', text: 'GenSpark', target: { scope: 'selection' }, italic: true },
    ])
    const block = editor.state.doc.child(1)
    const italic: string[] = []
    block.forEach((child) => {
      if (child.marks.some((m) => m.type.name === 'italic')) italic.push(child.text ?? '')
    })
    expect(italic).toEqual(['GenSpark'])
    expect(block.child(0).text).toBe('GenSpark intro,')
    expect(block.child(0).marks.some((m) => m.type.name === 'italic')).toBe(false)
  })
})

describe('op registry contract', () => {
  it('every registered op is documented in the system prompt and the apply_ops tool description', async () => {
    const { opNames, opSignatures } = await import('../src/renderer/ai/ops')
    const { AGENT_SYSTEM_PROMPT } = await import('../src/renderer/ai/protocol')
    const { AGENT_TOOLS } = await import('../src/renderer/ai/tools')
    const tool = AGENT_TOOLS.find((t) => t.name === 'apply_ops')
    expect(tool).toBeDefined()
    for (const name of opNames()) {
      expect(tool!.description).toContain(name)
      expect(AGENT_SYSTEM_PROMPT).toContain(`{ op: "${name}"`)
    }
    for (const signature of opSignatures()) expect(AGENT_SYSTEM_PROMPT).toContain(signature)
    expect(AGENT_SYSTEM_PROMPT).not.toContain('apply_commands')
  })

  it('dryRun validates and plans without touching the document', () => {
    const editor = createEditor(fixtureDoc())
    const before = JSON.stringify(editor.getJSON())
    const outcome = executeOps(
      editor,
      [
        { op: 'setFont', target: { nodeType: 'docHeading' }, color: '#ff0000' },
        { op: 'deleteBlocks', target: { blockIndexes: [3] } },
      ],
      { dryRun: true },
    )
    expect(outcome.ok).toBe(true)
    expect(outcome.plan).toEqual([
      '#0 setFont {"nodeType":"docHeading"}',
      '#1 deleteBlocks {"blockIndexes":[3]}',
    ])
    expect(outcome.results).toEqual([])
    expect(JSON.stringify(editor.getJSON())).toBe(before)
  })

  it('colors accept #RRGGBB in any case and are stored as upper-case hex without the hash', () => {
    const editor = createEditor(fixtureDoc())
    const outcome = executeOps(editor, [
      { op: 'setFont', target: { blockIndexes: [1] }, color: '#1a73e8' },
      { op: 'setParagraphFormat', target: { blockIndexes: [1] }, shadingFill: 'f2f2f2' },
    ])
    expect(outcome.ok).toBe(true)
    expect(textStyleOf(editor, 1)).toMatchObject({ color: '1A73E8' })
    expect(editor.state.doc.child(1).attrs.shadingFill).toBe('F2F2F2')
    const bad = executeOps(editor, [{ op: 'setFont', target: { blockIndexes: [1] }, color: 'red' }])
    expect(bad.ok).toBe(false)
    expect(bad.error).toContain('6-digit hex')
  })

  it('markAi: false applies the change without the AI highlight (the ribbon path)', () => {
    const editor = createEditor(fixtureDoc())
    const outcome = executeOps(
      editor,
      [{ op: 'setParagraphFormat', target: { blockIndexes: [1] }, align: 'right' }],
      { markAi: false },
    )
    expect(outcome.ok).toBe(true)
    expect(editor.state.doc.child(1).attrs.align).toBe('right')
    expect(editor.state.doc.child(1).attrs.aiChanged).toBe(false)
  })

  it('all validation errors of a batch are reported together, each with its usage line', () => {
    const editor = createEditor(fixtureDoc())
    const outcome = executeOps(editor, [
      { op: 'setHeadingLevel', target: { nodeType: 'docHeading' }, level: 9 },
      { op: 'moveBlocks', blockIndexes: [], afterBlockIndex: 0 },
    ])
    expect(outcome.ok).toBe(false)
    expect(outcome.error).toContain('op #0 setHeadingLevel: level must be')
    expect(outcome.error).toContain('op #1 moveBlocks: blockIndexes must be')
    expect(outcome.error?.match(/usage: \{ op: /g)).toHaveLength(2)
  })
})

describe('null clears numeric fields', () => {
  it('fontSize: null clears the size and widthPx/heightPx: null are accepted', () => {
    const editor = createEditor([
      para([
        text('sized', [{ type: 'docTextStyle', attrs: { sizeHalfPoints: 24, color: 'FF0000' } }]),
      ]),
    ])
    const outcome = executeOps(editor, [
      { op: 'setFont', target: { blockIndexes: [0] }, fontSize: null },
    ])
    expect(outcome.ok).toBe(true)
    expect(textStyleOf(editor, 0)).toMatchObject({ sizeHalfPoints: null, color: 'FF0000' })
    const zero = executeOps(editor, [{ op: 'setFont', target: { blockIndexes: [0] }, fontSize: 0 }])
    expect(zero.ok).toBe(false)
    const image = createEditor([
      {
        type: 'docProtected',
        attrs: {
          docxIndex: 7,
          blockType: 'image',
          label: 'Image',
          imageWidthPx: 400,
          imageHeightPx: 200,
        },
      } as JsonNode,
    ])
    const cleared = executeOps(image, [
      { op: 'setImageProperties', target: { nodeType: 'image' }, widthPx: null, align: 'center' },
    ])
    expect(cleared.ok).toBe(true)
  })
})

describe('UI ops: the ribbon issues the same ops as the model', () => {
  const cellTable = (): JsonNode => ({
    type: 'docTable',
    attrs: { docxIndex: null },
    content: [
      {
        type: 'docTableRow',
        content: [
          {
            type: 'docTableCell',
            content: [{ type: 'docParagraph', attrs: { docxIndex: null }, content: [text('A1')] }],
          },
          {
            type: 'docTableCell',
            content: [{ type: 'docParagraph', attrs: { docxIndex: null }, content: [text('B1')] }],
          },
        ],
      },
    ],
  })
  const cellPara = (editor: Editor, col: number) =>
    editor.state.doc.child(1).child(0).child(col).child(0)

  it('UI-only ops are unknown to the model but run for the UI without the AI highlight', async () => {
    const { runUiOps, opNames } = await import('../src/renderer/ai/ops')
    const { AGENT_SYSTEM_PROMPT } = await import('../src/renderer/ai/protocol')
    const editor = createEditor(fixtureDoc())
    const ai = executeOps(editor, [
      { op: 'setParagraphAttrs', target: { blockIndexes: [1] }, attrs: { align: 'right' } },
    ])
    expect(ai.ok).toBe(false)
    expect(ai.error).toContain('unknown op "setParagraphAttrs"')
    for (const name of ['setParagraphAttrs', 'stepIndent', 'stepHangingIndent']) {
      expect(opNames()).not.toContain(name)
      expect(AGENT_SYSTEM_PROMPT).not.toContain(`"${name}"`)
    }
    expect(
      runUiOps(editor, [
        { op: 'setParagraphAttrs', target: { blockIndexes: [1] }, attrs: { align: 'right' } },
      ]),
    ).toBe(true)
    expect(editor.state.doc.child(1).attrs.align).toBe('right')
    expect(editor.state.doc.child(1).attrs.aiChanged).toBe(false)
  })

  it('a range target stands in for the selection (blur-committed dialog inputs)', async () => {
    const { runUiOps } = await import('../src/renderer/ai/ops')
    const editor = createEditor(fixtureDoc())
    const block3Pos =
      editor.state.doc.child(0).nodeSize +
      editor.state.doc.child(1).nodeSize +
      editor.state.doc.child(2).nodeSize
    editor.commands.setTextSelection(2) // live selection in block 0
    runUiOps(
      editor,
      [
        {
          op: 'setParagraphAttrs',
          target: { range: { from: block3Pos + 2, to: block3Pos + 4 } },
          attrs: { spaceBefore: 240 },
        },
      ],
      { focus: false },
    )
    expect(editor.state.doc.child(3).attrs.spaceBefore).toBe(240)
    // an explicit spacing value turns Word's "Auto" spacing off
    expect(editor.state.doc.child(3).attrs.spaceBeforeAuto).toBe(false)
    expect(editor.state.doc.child(0).attrs.spaceBefore).toBeNull()
  })

  it('paragraph ops reach the paragraphs inside a table: the whole table for the model, the selected cell for the UI', async () => {
    const { runUiOps } = await import('../src/renderer/ai/ops')
    const ai = createEditor([heading([text('T')]), cellTable()])
    const outcome = executeOps(ai, [
      { op: 'setParagraphFormat', target: { blockIndexes: [1] }, align: 'center' },
    ])
    expect(outcome.ok).toBe(true)
    expect(outcome.results[0].changed).toBe(2)
    expect(cellPara(ai, 0).attrs.align).toBe('center')
    expect(cellPara(ai, 1).attrs.align).toBe('center')

    const ui = createEditor([heading([text('T')]), cellTable()])
    // caret inside cell B1: table(1) + row(1) + cellA(A1 para = 4, cell = 6) + cellB(1) + para(1)
    const b1 = ui.state.doc.child(0).nodeSize + 1 + 1 + 6 + 1 + 1 + 1
    expect(ui.state.doc.textBetween(b1 - 1, b1 + 1)).toBe('B1')
    ui.commands.setTextSelection(b1)
    runUiOps(ui, [
      { op: 'setParagraphAttrs', target: { scope: 'selection' }, attrs: { align: 'right' } },
    ])
    expect(cellPara(ui, 1).attrs.align).toBe('right')
    expect(cellPara(ui, 0).attrs.align).toBeNull()
  })

  it('UI alignment also lands on selected images as their w:jc; the model op skips them as protected', async () => {
    const { runUiOps } = await import('../src/renderer/ai/ops')
    const image = (): JsonNode => ({
      type: 'docProtected',
      attrs: {
        docxIndex: 7,
        blockType: 'image',
        label: 'Image',
        imageWidthPx: 400,
        imageHeightPx: 200,
      },
    })
    const ui = createEditor([para([text('a')]), image(), para([text('b')])])
    runUiOps(ui, [
      {
        op: 'setParagraphAttrs',
        target: { range: { from: 0, to: ui.state.doc.content.size } },
        attrs: { align: 'center' },
      },
    ])
    expect(ui.state.doc.child(1).attrs.imageAlign).toBe('center')
    expect(ui.state.doc.child(0).attrs.align).toBe('center')

    const ai = createEditor([para([text('a')]), image()])
    const outcome = executeOps(ai, [
      { op: 'setParagraphFormat', target: { blockIndexes: [0, 1] }, align: 'center' },
    ])
    expect(outcome.results[0]).toMatchObject({ changed: 1, skippedProtected: 1 })
    expect(ai.state.doc.child(1).attrs.imageAlign ?? null).toBeNull()
  })

  it('stepIndent treats each paragraph on its own: list items change level, paragraphs snap to half-inch stops', async () => {
    const { runUiOps } = await import('../src/renderer/ai/ops')
    const editor = createEditor([
      listItem([text('item')], { numId: '1', ilvl: 1 }),
      para([text('plain')], { indentLeft: 500 }),
    ])
    runUiOps(editor, [
      {
        op: 'stepIndent',
        target: { range: { from: 0, to: editor.state.doc.content.size } },
        delta: 1,
      },
    ])
    expect(editor.state.doc.child(0).attrs.ilvl).toBe(2)
    expect(editor.state.doc.child(1).attrs.indentLeft).toBe(720)
    runUiOps(editor, [{ op: 'stepHangingIndent', target: { blockIndexes: [1] }, delta: 1 }])
    expect(editor.state.doc.child(1).attrs.indentLeft).toBe(1440)
    expect(editor.state.doc.child(1).attrs.indentFirstLine).toBe(-720)
    const bad = executeOps(
      editor,
      [{ op: 'stepIndent', target: { blockIndexes: [1] }, delta: 2 }],
      { source: 'ui' },
    )
    expect(bad.ok).toBe(false)
    expect(bad.error).toContain('delta must be 1 or -1')
  })
})
