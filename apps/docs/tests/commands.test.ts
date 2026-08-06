import { afterEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { editorExtensions } from '../src/renderer/editor/extensions'
import {
  executeCommands,
  parseCommandEnvelope,
  type CommandEnvelope,
} from '../src/renderer/ai/commands'

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

describe('updateTextStyle', () => {
  it('sets color on all headings via nodeType, keeping bold and size (fields mask)', () => {
    const editor = createEditor(fixtureDoc())
    const outcome = executeCommands(editor, {
      commands: [
        {
          updateTextStyle: {
            target: { nodeType: 'docHeading' },
            style: { color: 'FF0000' },
            fields: ['color'],
          },
        },
      ],
    })
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
    const outcome = executeCommands(editor, {
      commands: [
        {
          updateTextStyle: {
            target: { nodeType: 'docParagraph' },
            style: { color: null },
            fields: ['color'],
          },
        },
      ],
    })
    expect(outcome.ok).toBe(true)
    expect(textStyleOf(editor, 0)).toMatchObject({ color: null, sizeHalfPoints: 24 })
    expect(textStyleOf(editor, 1)).toBeNull()
  })

  it('font routes to its script slot: CJK keeps the Latin font, Latin keeps the CJK font', () => {
    const mixed = () =>
      createEditor([
        para([
          text('合同 Contract', [
            { type: 'docTextStyle', attrs: { font: 'SimSun', fontAscii: 'Times New Roman' } },
          ]),
        ]),
      ])
    const style = (font: string | null) => ({
      commands: [
        {
          updateTextStyle: {
            target: { nodeType: 'docParagraph' as const },
            style: { font },
            fields: ['font' as const],
          },
        },
      ],
    })

    const cjk = mixed()
    expect(executeCommands(cjk, style('KaiTi')).ok).toBe(true)
    expect(textStyleOf(cjk, 0)).toMatchObject({ font: 'KaiTi', fontAscii: 'Times New Roman' })

    const latin = mixed()
    expect(executeCommands(latin, style('Arial')).ok).toBe(true)
    expect(textStyleOf(latin, 0)).toMatchObject({ font: 'SimSun', fontAscii: 'Arial' })

    const cleared = mixed()
    expect(executeCommands(cleared, style(null)).ok).toBe(true)
    expect(textStyleOf(cleared, 0)).toBeNull()
  })

  it('boolean fields add and remove basic marks', () => {
    const editor = createEditor([para([text('plain'), text('bold', [{ type: 'bold' }])])])
    executeCommands(editor, {
      commands: [
        {
          updateTextStyle: {
            target: { nodeType: 'docParagraph' },
            style: { bold: false, italic: true },
            fields: ['bold', 'italic'],
          },
        },
      ],
    })
    const block = editor.state.doc.child(0)
    block.forEach((child) => {
      expect(child.marks.some((m) => m.type.name === 'bold')).toBe(false)
      expect(child.marks.some((m) => m.type.name === 'italic')).toBe(true)
    })
  })
})

describe('updateTextStyle google-parity fields', () => {
  it('baselineOffset SUPERSCRIPT sets vertAlign, NONE clears it, other attrs survive', () => {
    const editor = createEditor([
      para([text('x2', [{ type: 'docTextStyle', attrs: { color: 'FF0000' } }])]),
    ])
    executeCommands(editor, {
      commands: [
        {
          updateTextStyle: {
            target: { nodeType: 'docParagraph' },
            style: { baselineOffset: 'SUPERSCRIPT' },
            fields: ['baselineOffset'],
          },
        },
      ],
    })
    expect(textStyleOf(editor, 0)).toMatchObject({ vertAlign: 'superscript', color: 'FF0000' })

    executeCommands(editor, {
      commands: [
        {
          updateTextStyle: {
            target: { nodeType: 'docParagraph' },
            style: { baselineOffset: 'NONE' },
            fields: ['baselineOffset'],
          },
        },
      ],
    })
    expect(textStyleOf(editor, 0)).toMatchObject({ vertAlign: null, color: 'FF0000' })
  })

  it('link adds and removes the link mark', () => {
    const editor = createEditor([para([text('official site')])])
    executeCommands(editor, {
      commands: [
        {
          updateTextStyle: {
            target: { nodeType: 'docParagraph' },
            style: { link: { url: 'https://example.com' } },
            fields: ['link'],
          },
        },
      ],
    })
    const child = editor.state.doc.child(0).child(0)
    expect(child.marks.find((m) => m.type.name === 'link')?.attrs.href).toBe('https://example.com')

    executeCommands(editor, {
      commands: [
        {
          updateTextStyle: {
            target: { nodeType: 'docParagraph' },
            style: { link: null },
            fields: ['link'],
          },
        },
      ],
    })
    expect(
      editor.state.doc
        .child(0)
        .child(0)
        .marks.some((m) => m.type.name === 'link'),
    ).toBe(false)
  })

  it('rejects a bad baselineOffset value', () => {
    const editor = createEditor([para([text('x')])])
    const outcome = executeCommands(editor, {
      commands: [
        {
          updateTextStyle: {
            target: { nodeType: 'docParagraph' },
            style: { baselineOffset: 'UP' },
            fields: ['baselineOffset'],
          },
        },
      ],
    } as unknown as CommandEnvelope)
    expect(outcome.ok).toBe(false)
    expect(outcome.error).toContain('baselineOffset')
  })
})

describe('updateParagraphStyle google-parity fields', () => {
  it('sets first-line indent and borders without touching other attrs', () => {
    const editor = createEditor(fixtureDoc())
    const outcome = executeCommands(editor, {
      commands: [
        {
          updateParagraphStyle: {
            target: { nodeType: 'docParagraph' },
            style: { indentFirstLine: 440, borders: 'b' },
            fields: ['indentFirstLine', 'borders'],
          },
        },
      ],
    })
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
    const outcome = executeCommands(editor, {
      commands: [
        {
          updateParagraphStyle: {
            target: { nodeType: 'docParagraph' },
            style: { lineSpacing: 1.5 },
            fields: ['lineSpacing'],
          },
        },
      ],
    })
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
    const outcome = executeCommands(editor, {
      commands: [
        {
          setHeadingLevel: {
            target: { nodeType: 'docHeading', headingLevel: 2, containsText: 'Risk Notes' },
            level: 3,
          },
        },
      ],
    })
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
    executeCommands(editor, {
      commands: [{ setHeadingLevel: { target: { containsText: 'Risk Notes' }, level: 0 } }],
    })
    const node = editor.state.doc.child(2)
    expect(node.type.name).toBe('docParagraph')
    expect(node.textContent).toBe('Risk Notes')
  })

  it('promotes a paragraph to a heading', () => {
    const editor = createEditor(fixtureDoc())
    executeCommands(editor, {
      commands: [{ setHeadingLevel: { target: { blockIndexes: [3] }, level: 2 } }],
    })
    const node = editor.state.doc.child(3)
    expect(node.type.name).toBe('docHeading')
    expect(node.attrs.level).toBe(2)
  })
})

describe('replaceAllText', () => {
  it('replaces every occurrence and keeps marks', () => {
    const editor = createEditor(fixtureDoc())
    const outcome = executeCommands(editor, {
      commands: [{ replaceAllText: { containsText: 'GenSpark', replaceText: 'Genspark' } }],
    })
    expect(outcome.ok).toBe(true)
    expect(outcome.results[0].detail).toBe('共替换 2 处')
    const block = editor.state.doc.child(1)
    expect(block.textContent).toBe('Genspark intro,Genspark is great')
    const boldChild = block.child(block.childCount - 1)
    expect(boldChild.marks.some((m) => m.type.name === 'bold')).toBe(true)
  })

  it('matchCase: sensitive by default, insensitive when false', () => {
    const editor = createEditor(fixtureDoc())
    const strict = executeCommands(editor, {
      commands: [{ replaceAllText: { containsText: 'genspark', replaceText: 'X' } }],
    })
    expect(strict.results[0].changed).toBe(0)
    expect(editor.state.doc.child(1).textContent).toContain('GenSpark')

    const loose = executeCommands(editor, {
      commands: [
        { replaceAllText: { containsText: 'genspark', replaceText: 'X', matchCase: false } },
      ],
    })
    expect(loose.results[0].detail).toBe('共替换 2 处')
    expect(editor.state.doc.child(1).textContent).toBe('X intro,X is great')
  })
})

describe('deleteBlocks', () => {
  it('deletes multiple blocks without index drift', () => {
    const editor = createEditor(fixtureDoc())
    const outcome = executeCommands(editor, {
      commands: [{ deleteBlocks: { target: { blockIndexes: [1, 3] } } }],
    })
    expect(outcome.ok).toBe(true)
    expect(editor.state.doc.childCount).toBe(4)
    expect(editor.state.doc.child(0).textContent).toBe('Chapter 1 Overview')
    expect(editor.state.doc.child(1).textContent).toBe('Risk Notes')
    expect(editor.state.doc.child(2).textContent).toBe('List item')
    expect(editor.state.doc.child(3).type.name).toBe('docProtected')
  })

  it('deleting every block leaves one empty paragraph', () => {
    const editor = createEditor([para([text('a')]), para([text('b')])])
    executeCommands(editor, {
      commands: [{ deleteBlocks: { target: { blockIndexes: [0, 1] } } }],
    })
    expect(editor.state.doc.childCount).toBe(1)
    expect(editor.state.doc.child(0).textContent).toBe('')
  })
})

describe('moveBlocks', () => {
  it('moves blocks preserving relative order and docxIndex', () => {
    const editor = createEditor(fixtureDoc())
    const outcome = executeCommands(editor, {
      commands: [{ moveBlocks: { blockIndexes: [3, 4], afterBlockIndex: 0 } }],
    })
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
    executeCommands(editor, {
      commands: [{ moveBlocks: { blockIndexes: [2], afterBlockIndex: -1 } }],
    })
    expect(editor.state.doc.child(0).textContent).toBe('Risk Notes')
  })

  it('rejects the whole envelope when afterBlockIndex is a moved block', () => {
    const editor = createEditor(fixtureDoc())
    const before = JSON.stringify(editor.getJSON())
    const outcome = executeCommands(editor, {
      commands: [{ moveBlocks: { blockIndexes: [1], afterBlockIndex: 1 } }],
    })
    expect(outcome.ok).toBe(false)
    expect(JSON.stringify(editor.getJSON())).toBe(before)
  })
})

describe('protected blocks', () => {
  it('style commands skip docProtected and count it', () => {
    const editor = createEditor(fixtureDoc())
    const before = JSON.stringify(editor.getJSON())
    const outcome = executeCommands(editor, {
      commands: [
        {
          updateTextStyle: {
            target: { blockIndexes: [5] },
            style: { color: 'FF0000' },
            fields: ['color'],
          },
        },
      ],
    })
    expect(outcome.ok).toBe(true)
    expect(outcome.results[0]).toMatchObject({ matched: 1, changed: 0, skippedProtected: 1 })
    expect(JSON.stringify(editor.getJSON())).toBe(before)
    expect(outcome.summary).toContain('受保护块')
  })

  it('deleteBlocks does delete docProtected', () => {
    const editor = createEditor(fixtureDoc())
    executeCommands(editor, { commands: [{ deleteBlocks: { target: { blockIndexes: [5] } } }] })
    expect(editor.state.doc.childCount).toBe(5)
    editor.state.doc.forEach((n) => expect(n.type.name).not.toBe('docProtected'))
  })
})

describe('envelope validation', () => {
  it('unknown command rejects the whole envelope with no doc change', () => {
    const editor = createEditor(fixtureDoc())
    const before = JSON.stringify(editor.getJSON())
    const outcome = executeCommands(editor, {
      commands: [
        { replaceAllText: { containsText: 'GenSpark', replaceText: 'X' } },
        { insertTable: { rows: 2 } },
      ],
    } as unknown as CommandEnvelope)
    expect(outcome.ok).toBe(false)
    expect(outcome.error).toContain('unknown command')
    expect(JSON.stringify(editor.getJSON())).toBe(before)
  })

  it('missing fields rejects updateTextStyle', () => {
    const editor = createEditor(fixtureDoc())
    const outcome = executeCommands(editor, {
      commands: [
        { updateTextStyle: { target: { nodeType: 'docHeading' }, style: { color: 'FF0000' } } },
      ],
    } as unknown as CommandEnvelope)
    expect(outcome.ok).toBe(false)
    expect(outcome.error).toContain('fields')
  })

  it('empty target rejects', () => {
    const editor = createEditor(fixtureDoc())
    const outcome = executeCommands(editor, {
      commands: [{ deleteBlocks: { target: {} } }],
    })
    expect(outcome.ok).toBe(false)
    expect(outcome.error).toContain('requires at least one condition')
  })
})

describe('transaction atomicity and aiChanged', () => {
  it('a multi-command round undoes in a single step', () => {
    const editor = createEditor(fixtureDoc())
    const before = JSON.stringify(editor.getJSON())
    const outcome = executeCommands(editor, {
      commands: [
        {
          updateTextStyle: {
            target: { nodeType: 'docHeading' },
            style: { color: 'FF0000' },
            fields: ['color'],
          },
        },
        { replaceAllText: { containsText: 'GenSpark', replaceText: 'Genspark' } },
        { deleteBlocks: { target: { blockIndexes: [3] } } },
      ],
    })
    expect(outcome.ok).toBe(true)
    expect(JSON.stringify(editor.getJSON())).not.toBe(before)
    editor.commands.undo()
    expect(JSON.stringify(editor.getJSON())).toBe(before)
  })

  it('marks changed nodes aiChanged and leaves unmatched attrs untouched', () => {
    const editor = createEditor(fixtureDoc())
    const untouchedBefore = { ...editor.state.doc.child(3).attrs }
    executeCommands(editor, {
      commands: [
        {
          updateTextStyle: {
            target: { nodeType: 'docHeading' },
            style: { color: 'FF0000' },
            fields: ['color'],
          },
        },
      ],
    })
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
    const outcome = executeCommands(editor, {
      commands: [
        {
          updateParagraphStyle: {
            target: { nodeType: 'docParagraph', scope: 'selection' },
            style: { align: 'right' },
            fields: ['align'],
          },
        },
      ],
    })
    expect(outcome.results[0].changed).toBe(1)
    expect(editor.state.doc.child(1).attrs.align).toBe('right')
    expect(editor.state.doc.child(3).attrs.align).toBe('center')
  })
})

describe('createParagraphBullets / deleteParagraphBullets', () => {
  const NUM_IDS = { numIds: { bullet: '10', ordered: '20' } }

  it('converts paragraphs to bullet list items with the context numId', () => {
    const editor = createEditor(fixtureDoc())
    const outcome = executeCommands(
      editor,
      { commands: [{ createParagraphBullets: { target: { blockIndexes: [1, 3] } } }] },
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
    executeCommands(
      editor,
      {
        commands: [
          {
            createParagraphBullets: {
              target: { blockIndexes: [3, 4] },
              bulletPreset: 'NUMBERED_DECIMAL_ALPHA_ROMAN',
            },
          },
        ],
      },
      NUM_IDS,
    )
    expect(editor.state.doc.child(3).attrs.kind).toBe('ordered')
    expect(editor.state.doc.child(3).attrs.numId).toBe('20')
    // block 4 was already a bullet list item: kind switched
    expect(editor.state.doc.child(4).attrs.kind).toBe('ordered')
  })

  it('headings are matched but never converted', () => {
    const editor = createEditor(fixtureDoc())
    const outcome = executeCommands(
      editor,
      { commands: [{ createParagraphBullets: { target: { blockIndexes: [0] } } }] },
      NUM_IDS,
    )
    expect(outcome.results[0]).toMatchObject({ matched: 1, changed: 0 })
    expect(editor.state.doc.child(0).type.name).toBe('docHeading')
  })

  it('deleteParagraphBullets converts list items back to paragraphs, keeping text', () => {
    const editor = createEditor(fixtureDoc())
    const outcome = executeCommands(editor, {
      commands: [{ deleteParagraphBullets: { target: { nodeType: 'docListItem' } } }],
    })
    expect(outcome.results[0].changed).toBe(1)
    const node = editor.state.doc.child(4)
    expect(node.type.name).toBe('docParagraph')
    expect(node.textContent).toBe('List item')
    expect(node.attrs.aiChanged).toBe(true)
  })

  it('round trip undoes in one step', () => {
    const editor = createEditor(fixtureDoc())
    const before = JSON.stringify(editor.getJSON())
    executeCommands(
      editor,
      { commands: [{ createParagraphBullets: { target: { blockIndexes: [1, 3] } } }] },
      NUM_IDS,
    )
    editor.commands.undo()
    expect(JSON.stringify(editor.getJSON())).toBe(before)
  })
})

describe('updateImageProperties', () => {
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
    const outcome = executeCommands(editor, {
      commands: [
        {
          updateImageProperties: {
            target: { nodeType: 'image' },
            properties: { align: 'center' },
            fields: ['align'],
          },
        },
      ],
    })
    expect(outcome.ok).toBe(true)
    expect(outcome.results[0]).toMatchObject({ matched: 1, changed: 1 })
    expect(editor.state.doc.child(1).attrs.imageAlign).toBe('center')
    // table untouched
    expect(editor.state.doc.child(2).attrs.blockType).toBe('table')
  })

  it('scales the other dimension proportionally when only widthPx is given', () => {
    const editor = createEditor(imageDoc())
    executeCommands(editor, {
      commands: [
        {
          updateImageProperties: {
            target: { nodeType: 'image' },
            properties: { widthPx: 200 },
            fields: ['widthPx'],
          },
        },
      ],
    })
    const node = editor.state.doc.child(1)
    expect(node.attrs.imageWidthPx).toBe(200)
    expect(node.attrs.imageHeightPx).toBe(100)
  })

  it('nodeType image never matches non-image blocks', () => {
    const editor = createEditor(imageDoc())
    const outcome = executeCommands(editor, {
      commands: [{ deleteBlocks: { target: { nodeType: 'image' } } }],
    })
    expect(outcome.results[0].matched).toBe(1)
    expect(editor.state.doc.childCount).toBe(2)
    expect(editor.state.doc.child(0).textContent).toBe('before')
  })

  it('rejects unknown fields', () => {
    const editor = createEditor(imageDoc())
    const outcome = executeCommands(editor, {
      commands: [
        {
          updateImageProperties: {
            target: { nodeType: 'image' },
            properties: { align: 'center' },
            fields: ['align', 'rotation'],
          },
        },
      ],
    } as unknown as CommandEnvelope)
    expect(outcome.ok).toBe(false)
    expect(outcome.error).toContain('unknown fields')
  })
})

describe('parseCommandEnvelope', () => {
  it('extracts a fenced envelope', () => {
    const envelope = parseCommandEnvelope(
      '```json\n{"commands":[{"replaceAllText":{"containsText":"a","replaceText":"b"}}]}\n```',
    )
    expect(envelope?.commands).toHaveLength(1)
  })

  it('accepts a bare JSON object', () => {
    const envelope = parseCommandEnvelope('{"commands":[]}')
    expect(envelope).not.toBeNull()
  })

  it('returns null for HTML content', () => {
    expect(parseCommandEnvelope('<p>hello</p>')).toBeNull()
    expect(parseCommandEnvelope('Done:{"commands":[]}')).toBeNull()
  })
})
