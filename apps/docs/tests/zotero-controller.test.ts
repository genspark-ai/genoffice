import { Editor, Mark, Node } from '@tiptap/core'
import { afterEach, describe, expect, it } from 'vitest'
import { ZoteroDocumentController } from '../src/renderer/zotero/controller'

const Doc = Node.create({ name: 'doc', topNode: true, content: 'paragraph+' })
const Paragraph = Node.create({
  name: 'paragraph',
  group: 'block',
  content: 'inline*',
  addAttributes: () => ({
    docxIndex: { default: null },
    indentLeft: { default: null },
    indentRight: { default: null },
    indentFirstLine: { default: null },
    spaceBefore: { default: null },
    spaceAfter: { default: null },
    align: { default: null },
  }),
  parseHTML: () => [{ tag: 'p' }],
  renderHTML: () => ['p', 0],
})
const Text = Node.create({ name: 'text', group: 'inline' })
const Bold = Mark.create({ name: 'bold', renderHTML: () => ['strong', 0] })
const Italic = Mark.create({ name: 'italic', renderHTML: () => ['em', 0] })
const Underline = Mark.create({ name: 'underline', renderHTML: () => ['u', 0] })
const Strike = Mark.create({ name: 'strike', renderHTML: () => ['s', 0] })
const DocTextStyle = Mark.create({
  name: 'docTextStyle',
  addAttributes: () => ({
    vertAlign: { default: null },
    caps: { default: null },
    sizeHalfPoints: { default: null },
  }),
  renderHTML: () => ['span', 0],
})
const InstrField = Mark.create({
  name: 'instrField',
  inclusive: false,
  addAttributes: () => ({
    instr: { default: '' },
    beginXml: { default: null },
    fieldId: { default: null, rendered: false },
    fieldPart: { default: null, rendered: false },
  }),
  renderHTML: ({ mark }) => ['span', { 'data-instr-field': mark.attrs.instr }, 0],
})

let editor: Editor | null = null

interface TestJsonNode {
  text?: string
  attrs?: Record<string, unknown>
  marks?: Array<{ type: string; attrs?: Record<string, unknown> }>
  content?: TestJsonNode[]
}

afterEach(() => {
  editor?.destroy()
  editor = null
})

describe('ZoteroDocumentController', () => {
  const extensions = [
    Doc,
    Paragraph,
    Text,
    Bold,
    Italic,
    Underline,
    Strike,
    DocTextStyle,
    InstrField,
  ]

  it('inserts and updates a Zotero field using the LibreOffice wire contract', async () => {
    let documentData = ''
    editor = new Editor({
      element: document.createElement('div'),
      extensions,
      content: { type: 'doc', content: [{ type: 'paragraph' }] },
    })
    const controller = new ZoteroDocumentController(editor, {
      get: () => documentData,
      set: (value) => {
        documentData = value
      },
    })
    const call = (command: string, args: unknown[]) =>
      controller.handle({ requestId: 'test', command, args })

    expect(await call('Application_getActiveDocument', [3])).toEqual([3, 1])
    const inserted = (await call('Document_insertField', [1, 'ReferenceMark', 0])) as [
      number,
      string,
      number,
    ]
    expect(inserted.slice(1)).toEqual(['TEMP', 0])

    const code = 'ITEM CSL_CITATION {"citationID":"citation-1"}'
    await call('Field_setCode', [1, inserted[0], code])
    await call('Field_setText', [1, inserted[0], '(Smith, 2024)', false])
    expect(await call('Field_getText', [1, inserted[0]])).toBe('(Smith, 2024)')
    expect(await call('Document_getFields', [1, 'ReferenceMark'])).toEqual([
      [inserted[0]],
      [code],
      [0],
      [-1],
    ])

    await call('Document_setDocumentData', [1, '<data data-version="3"/>'])
    expect(await call('Document_getDocumentData', [1])).toBe('<data data-version="3"/>')
    await call('Field_removeCode', [1, inserted[0]])
    expect(editor.getText()).toBe('(Smith, 2024)')
    expect(await call('Document_getFields', [1, 'ReferenceMark'])).toEqual([[], [], [], []])
  })

  it('keeps a multi-paragraph Zotero bibliography as one field', async () => {
    editor = new Editor({
      element: document.createElement('div'),
      extensions,
      content: { type: 'doc', content: [{ type: 'paragraph' }] },
    })
    const controller = new ZoteroDocumentController(editor, { get: () => '', set: () => {} })
    const call = (command: string, args: unknown[]) =>
      controller.handle({ requestId: 'test', command, args })
    const [fieldId] = (await call('Document_insertField', [1, 'ReferenceMark', 0])) as [
      number,
      string,
      number,
    ]
    const code = 'BIBL {"uncited":[],"omitted":[],"custom":[]} CSL_BIBLIOGRAPHY'
    await call('Field_setCode', [1, fieldId, code])
    await call('Field_setText', [
      1,
      fieldId,
      '{\\rtf1 Alpha, A. (2024).\\par Beta, B. (2023).\\par Gamma, G. (2022).}',
      true,
    ])

    const blocks = (editor.getJSON().content ?? []) as TestJsonNode[]
    expect(editor.state.doc.content.content.map((block) => block.textContent)).toEqual([
      'Alpha, A. (2024).',
      'Beta, B. (2023).',
      'Gamma, G. (2022).',
    ])
    expect(
      blocks.map(
        (block) =>
          block.content?.[0]?.marks?.find((mark) => mark.type === 'instrField')?.attrs?.fieldPart,
      ),
    ).toEqual(['begin', 'inside', 'end'])
    expect(await call('Field_getText', [1, fieldId])).toBe(
      'Alpha, A. (2024).\nBeta, B. (2023).\nGamma, G. (2022).',
    )
    expect(await call('Document_getFields', [1, 'ReferenceMark'])).toEqual([
      [fieldId],
      [code],
      [0],
      [-1],
    ])

    await call('Field_removeCode', [1, fieldId])
    expect(editor.state.doc.content.content.map((block) => block.textContent)).toEqual([
      'Alpha, A. (2024).',
      'Beta, B. (2023).',
      'Gamma, G. (2022).',
    ])
  })

  it('maps Zotero RTF styles and hanging indents into editable document formatting', async () => {
    editor = new Editor({
      element: document.createElement('div'),
      extensions,
      content: { type: 'doc', content: [{ type: 'paragraph' }] },
    })
    const controller = new ZoteroDocumentController(editor, { get: () => '', set: () => {} })
    const call = (command: string, args: unknown[]) =>
      controller.handle({ requestId: 'test', command, args })
    const [fieldId] = (await call('Document_insertField', [1, 'ReferenceMark', 0])) as [
      number,
      string,
      number,
    ]
    await call('Field_setCode', [1, fieldId, 'BIBL {} CSL_BIBLIOGRAPHY'])
    await call('Field_setText', [
      1,
      fieldId,
      '{\\rtf1\\pard\\li720\\fi-360 Alpha {\\i Journal} {\\super 2}\\par ' +
        '\\pard\\li720\\fi-360 {\\b Beta} {\\sub n}\\par}',
      true,
    ])

    const blocks = (editor.getJSON().content ?? []) as TestJsonNode[]
    expect(blocks.map((block) => block.attrs)).toEqual([
      expect.objectContaining({ indentLeft: 720, indentFirstLine: -360 }),
      expect.objectContaining({ indentLeft: 720, indentFirstLine: -360 }),
    ])
    const textNodes = blocks.flatMap((block) => block.content ?? [])
    const marksOf = (text: string) => textNodes.find((node) => node.text === text)?.marks ?? []
    expect(marksOf('Journal').map((mark) => mark.type)).toContain('italic')
    expect(marksOf('2')).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          type: 'docTextStyle',
          attrs: expect.objectContaining({ vertAlign: 'superscript' }),
        }),
      ]),
    )
    expect(marksOf('Beta').map((mark) => mark.type)).toContain('bold')
    expect(marksOf('n')).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          type: 'docTextStyle',
          attrs: expect.objectContaining({ vertAlign: 'subscript' }),
        }),
      ]),
    )

    await call('Field_removeCode', [1, fieldId])
    const withoutCodes = ((editor.getJSON().content ?? []) as TestJsonNode[]).flatMap(
      (block) => block.content ?? [],
    )
    expect(
      withoutCodes
        .find((node) => node.text === 'Journal')
        ?.marks?.some((mark) => mark.type === 'italic'),
    ).toBe(true)
  })

  it('answers Zotero alerts with the button index Zotero expects', async () => {
    editor = new Editor({
      element: document.createElement('div'),
      extensions,
      content: { type: 'doc', content: [{ type: 'paragraph' }] },
    })
    const controller = new ZoteroDocumentController(editor, { get: () => '', set: () => {} })
    const call = (command: string, args: unknown[]) =>
      controller.handle({ requestId: 'test', command, args })
    const confirm = window.confirm
    try {
      window.confirm = () => true
      expect(await call('Document_displayAlert', [1, 'Replace?', 2, 3])).toBe(2)
      expect(await call('Document_displayAlert', [1, 'Continue?', 2, 1])).toBe(1)
      window.confirm = () => false
      expect(await call('Document_displayAlert', [1, 'Replace?', 2, 3])).toBe(0)
    } finally {
      window.confirm = confirm
    }
  })

  it('inserts Zotero text literally instead of parsing it as HTML', async () => {
    editor = new Editor({
      element: document.createElement('div'),
      extensions,
      content: { type: 'doc', content: [{ type: 'paragraph' }] },
    })
    const controller = new ZoteroDocumentController(editor, { get: () => '', set: () => {} })
    await controller.handle({
      requestId: 'test',
      command: 'Document_insertText',
      args: [1, '{\\rtf1 x <b>bold</b> & y}'],
    })
    expect(editor.getText()).toBe('x <b>bold</b> & y')
  })

  it('keeps the text around a field whose new result spans paragraphs', async () => {
    const mark = {
      type: 'instrField',
      attrs: { instr: 'ADDIN ZOTERO_BIBL {} CSL_BIBLIOGRAPHY', fieldId: 7, fieldPart: 'single' },
    }
    editor = new Editor({
      element: document.createElement('div'),
      extensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'paragraph',
            content: [
              { type: 'text', text: 'References: ' },
              { type: 'text', text: '(pending)', marks: [mark] },
              { type: 'text', text: ' trailing.' },
            ],
          },
        ],
      },
    })
    const controller = new ZoteroDocumentController(editor, { get: () => '', set: () => {} })
    await controller.handle({
      requestId: 'test',
      command: 'Field_setText',
      args: [1, 7, '{\\rtf1 Alpha\\par Beta}', true],
    })
    expect(editor.state.doc.content.content.map((block) => block.textContent)).toEqual([
      'References: Alpha',
      'Beta trailing.',
    ])
    const blocks = (editor.getJSON().content ?? []) as TestJsonNode[]
    const parts = blocks.flatMap((block) =>
      (block.content ?? [])
        .map((node) => node.marks?.find((m) => m.type === 'instrField')?.attrs?.fieldPart)
        .filter(Boolean),
    )
    expect(parts).toEqual(['begin', 'end'])
    expect(
      await controller.handle({ requestId: 'test', command: 'Field_getText', args: [1, 7] }),
    ).toBe('Alpha\nBeta')
  })

  it('writes ids back onto pasted fields so later Field_ calls resolve', async () => {
    const instr = 'ADDIN ZOTERO_BIBL {} CSL_BIBLIOGRAPHY'
    const lost = { type: 'instrField', attrs: { instr } }
    const other = { type: 'instrField', attrs: { instr: 'ADDIN ZOTERO_ITEM CSL_CITATION {}' } }
    editor = new Editor({
      element: document.createElement('div'),
      extensions,
      content: {
        type: 'doc',
        content: [
          { type: 'paragraph', content: [{ type: 'text', text: 'Alpha', marks: [lost] }] },
          { type: 'paragraph', content: [{ type: 'text', text: 'Beta', marks: [lost] }] },
          { type: 'paragraph', content: [{ type: 'text', text: '(Doe)', marks: [other] }] },
        ],
      },
    })
    const controller = new ZoteroDocumentController(editor, { get: () => '', set: () => {} })
    const call = (command: string, args: unknown[]) =>
      controller.handle({ requestId: 'test', command, args })
    const [ids] = (await call('Document_getFields', [1, 'ReferenceMark'])) as [number[]]
    expect(ids).toHaveLength(2)
    expect(await call('Document_getFields', [1, 'ReferenceMark'])).toEqual(
      expect.arrayContaining([ids]),
    )
    expect(await call('Field_getText', [1, ids[0]])).toBe('Alpha\nBeta')
    expect(await call('Field_getText', [1, ids[1]])).toBe('(Doe)')
    const blocks = (editor.getJSON().content ?? []) as TestJsonNode[]
    expect(blocks.map((block) => block.content?.[0]?.marks?.[0]?.attrs?.fieldPart)).toEqual([
      'begin',
      'end',
      'single',
    ])
  })

  it('deletes a field spanning paragraphs without the text around it', async () => {
    const attrs = { instr: 'ADDIN ZOTERO_BIBL {} CSL_BIBLIOGRAPHY', fieldId: 9 }
    editor = new Editor({
      element: document.createElement('div'),
      extensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'paragraph',
            content: [
              { type: 'text', text: 'References: ' },
              {
                type: 'text',
                text: 'Alpha',
                marks: [{ type: 'instrField', attrs: { ...attrs, fieldPart: 'begin' } }],
              },
            ],
          },
          {
            type: 'paragraph',
            content: [
              {
                type: 'text',
                text: 'Beta',
                marks: [{ type: 'instrField', attrs: { ...attrs, fieldPart: 'end' } }],
              },
              { type: 'text', text: ' trailing.' },
            ],
          },
        ],
      },
    })
    const controller = new ZoteroDocumentController(editor, { get: () => '', set: () => {} })
    const call = (command: string, args: unknown[]) =>
      controller.handle({ requestId: 'test', command, args })
    editor.commands.setTextSelection(3)
    expect(await call('Document_cursorInField', [1])).toBeNull()
    editor.commands.setTextSelection(15)
    expect(await call('Document_cursorInField', [1])).toEqual([9, expect.any(String), 0])
    await call('Field_delete', [1, 9])
    expect(editor.state.doc.content.content.map((block) => block.textContent)).toEqual([
      'References:  trailing.',
    ])
  })

  it('keeps two pasted copies of a citation apart when text sits between them', async () => {
    const lost = { type: 'instrField', attrs: { instr: 'ADDIN ZOTERO_ITEM CSL_CITATION {}' } }
    editor = new Editor({
      element: document.createElement('div'),
      extensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'paragraph',
            content: [
              { type: 'text', text: '(Doe)', marks: [lost] },
              { type: 'text', text: ' and again ' },
              { type: 'text', text: '(Doe)', marks: [lost] },
            ],
          },
        ],
      },
    })
    const controller = new ZoteroDocumentController(editor, { get: () => '', set: () => {} })
    const [ids] = (await controller.handle({
      requestId: 'test',
      command: 'Document_getFields',
      args: [1, 'ReferenceMark'],
    })) as [number[]]
    expect(ids).toHaveLength(2)
    expect(editor.getText()).toBe('(Doe) and again (Doe)')
  })
})
