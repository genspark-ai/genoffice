import { Editor, Mark, Node } from '@tiptap/core'
import { afterEach, describe, expect, it, vi } from 'vitest'
import type { ZoteroLocalCiteData } from '../src/shared/ipc'
import { ZoteroDocumentController } from '../src/renderer/zotero/controller'
import {
  CSL_CITATION_SCHEMA,
  citationInstruction,
  citeFromZotero,
  clusterText,
  documentStyleId,
  hasZoteroFields,
  insertZoteroCitation,
  plainCitation,
  zoteroLocalErrorKey,
} from '../src/renderer/zotero/quick-cite'
import {
  loadZoteroAutoRefresh,
  saveZoteroAutoRefresh,
  shouldAutoRefresh,
} from '../src/renderer/zotero/auto-refresh'
import { collectionTree } from '../src/renderer/components/ZoteroLibraryPane'

const Doc = Node.create({ name: 'doc', topNode: true, content: 'paragraph+' })
const Paragraph = Node.create({
  name: 'paragraph',
  group: 'block',
  content: 'inline*',
  parseHTML: () => [{ tag: 'p' }],
  renderHTML: () => ['p', 0],
})
const Text = Node.create({ name: 'text', group: 'inline' })
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
afterEach(() => {
  editor?.destroy()
  editor = null
  localStorage.clear()
})

function makeEditor(text = 'As shown before.', editable = true): Editor {
  editor = new Editor({
    element: document.createElement('div'),
    extensions: [Doc, Paragraph, Text, InstrField],
    content: { type: 'doc', content: [{ type: 'paragraph', content: [{ type: 'text', text }] }] },
    editable,
  })
  return editor
}

const smith: ZoteroLocalCiteData = {
  key: 'ABCD2345',
  libraryId: 18084754,
  itemData: { id: 'http://zotero.org/users/18084754/items/ABCD2345', type: 'article-journal' },
  citation: '<span>(Smith, 2020)</span>',
}
const lee: ZoteroLocalCiteData = {
  key: 'WXYZ6789',
  libraryId: 18084754,
  itemData: { id: 'http://zotero.org/users/18084754/items/WXYZ6789', type: 'book' },
  citation: '(Lee &amp; Park, 2021)',
}

function instructionPayload(instr: string) {
  const prefix = 'ADDIN ZOTERO_ITEM CSL_CITATION '
  expect(instr.startsWith(prefix)).toBe(true)
  return JSON.parse(instr.slice(prefix.length)) as {
    citationID: string
    properties: { formattedCitation: string; plainCitation: string; noteIndex: number }
    citationItems: Array<{ id: unknown; uris?: string[]; itemData: unknown }>
    schema: string
  }
}

describe('Quick Cite field text', () => {
  it('reads the document style from its Zotero preferences, apa until one is set', () => {
    expect(documentStyleId('')).toBe('apa')
    expect(
      documentStyleId(
        '<data data-version="3"><style id="http://www.zotero.org/styles/ieee" hasBibliography="1"/></data>',
      ),
    ).toBe('ieee')
    expect(documentStyleId('<data><style id="file:///custom.csl"/></data>')).toBe('apa')
  })

  it('turns the HTML Zotero formats into plain text', () => {
    expect(plainCitation('<i>Nature</i> (2020)')).toBe('Nature (2020)')
    expect(plainCitation('(Lee &amp; Park, 2021)')).toBe('(Lee & Park, 2021)')
    expect(plainCitation('(Julianti &#38; Tanjung, 2023)')).toBe('(Julianti & Tanjung, 2023)')
    expect(plainCitation('Fran&#xE7;ois')).toBe('François')
  })

  it('joins citations that share their brackets into one cluster', () => {
    expect(clusterText(['(Smith, 2020)', '(Lee, 2021)'])).toBe('(Smith, 2020; Lee, 2021)')
    expect(clusterText(['[Smith 2020]', '[Lee 2021]'])).toBe('[Smith 2020; Lee 2021]')
    // a numeric style numbers each item alone: Zotero's Refresh numbers the cluster
    expect(clusterText(['[1]', '[1]'])).toBe('{Citation}')
    expect(clusterText(['(1)', '(1)'])).toBe('{Citation}')
    expect(clusterText(['[1]'])).toBe('{Citation}')
    expect(clusterText(['Smith (2020)', '(Lee, 2021)'])).toBe('Smith (2020); (Lee, 2021)')
    expect(clusterText(['(Smith, 2020)'])).toBe('(Smith, 2020)')
  })

  it('writes the field instruction Zotero writes, with each item URI and CSL-JSON', () => {
    const payload = instructionPayload(citationInstruction([smith, lee], '(A; B)', 'abc12345'))
    expect(payload.citationID).toBe('abc12345')
    expect(payload.properties).toEqual({
      formattedCitation: '(A; B)',
      plainCitation: '(A; B)',
      noteIndex: 0,
    })
    expect(payload.schema).toBe(CSL_CITATION_SCHEMA)
    expect(payload.citationItems).toEqual([
      {
        id: smith.itemData.id,
        uris: ['http://zotero.org/users/18084754/items/ABCD2345'],
        itemData: smith.itemData,
      },
      {
        id: lee.itemData.id,
        uris: ['http://zotero.org/users/18084754/items/WXYZ6789'],
        itemData: lee.itemData,
      },
    ])
  })

  it('leaves out the URI of an item in an unsynced library', () => {
    const local = { ...smith, libraryId: 0, itemData: { type: 'book' } }
    const payload = instructionPayload(citationInstruction([local], '(Smith, 2020)'))
    expect(payload.citationItems).toEqual([{ id: 'ABCD2345', itemData: { type: 'book' } }])
  })
})

describe('inserting a Quick Cite citation', () => {
  it('puts one field at the caret that the Zotero integration reads as its own', async () => {
    const ed = makeEditor()
    ed.commands.setTextSelection(1 + 'As shown before'.length)
    expect(hasZoteroFields(ed)).toBe(false)
    expect(insertZoteroCitation(ed, [smith, lee])).toBe(true)
    expect(ed.getText()).toBe('As shown before(Smith, 2020; Lee & Park, 2021).')
    expect(hasZoteroFields(ed)).toBe(true)

    const controller = new ZoteroDocumentController(ed, { get: () => '', set: () => {} })
    const [ids, codes] = (await controller.handle({
      requestId: 'test',
      command: 'Document_getFields',
      args: [1, 'Field'],
    })) as [unknown[], string[], number[], number[]]
    expect(ids).toHaveLength(1)
    expect(codes[0].startsWith('ITEM CSL_CITATION {')).toBe(true)
    expect(JSON.parse(codes[0].slice('ITEM CSL_CITATION '.length)).citationItems).toHaveLength(2)
    expect(
      await controller.handle({ requestId: 'test', command: 'Field_getText', args: [1, ids[0]] }),
    ).toBe('(Smith, 2020; Lee & Park, 2021)')
  })

  it('refuses a read-only document', () => {
    const ed = makeEditor('Locked.', false)
    expect(insertZoteroCitation(ed, [smith])).toBe(false)
    expect(hasZoteroFields(ed)).toBe(false)
  })
})

describe('citeFromZotero', () => {
  const prefs = '<data><style id="http://www.zotero.org/styles/ieee"/></data>'

  it('asks Zotero in the document style and refreshes a document that has preferences', async () => {
    const ed = makeEditor()
    const citeData = vi.fn().mockResolvedValue({ ok: true, value: [smith] })
    const refresh = vi.fn().mockResolvedValue(undefined)
    const result = await citeFromZotero(ed, ['ABCD2345'], {
      documentData: prefs,
      noteFields: false,
      citeData,
      refresh,
    })
    expect(result).toEqual({ ok: true, refreshed: true })
    expect(citeData).toHaveBeenCalledWith(['ABCD2345'], 'ieee')
    expect(refresh).toHaveBeenCalledTimes(1)
    expect(hasZoteroFields(ed)).toBe(true)
  })

  it('leaves a document without preferences unrefreshed, so Zotero opens no dialog', async () => {
    const ed = makeEditor()
    const refresh = vi.fn()
    const result = await citeFromZotero(ed, ['ABCD2345'], {
      documentData: '',
      noteFields: false,
      citeData: vi.fn().mockResolvedValue({ ok: true, value: [smith] }),
      refresh,
    })
    expect(result).toEqual({ ok: true, refreshed: false })
    expect(refresh).not.toHaveBeenCalled()
  })

  it('does not refresh when citations in notes would drop out of the bibliography', async () => {
    const refresh = vi.fn()
    await citeFromZotero(makeEditor(), ['ABCD2345'], {
      documentData: prefs,
      noteFields: true,
      citeData: vi.fn().mockResolvedValue({ ok: true, value: [smith] }),
      refresh,
    })
    expect(refresh).not.toHaveBeenCalled()
  })

  it('keeps the citation when the refresh fails', async () => {
    const ed = makeEditor()
    vi.spyOn(console, 'error').mockImplementation(() => {})
    const result = await citeFromZotero(ed, ['ABCD2345'], {
      documentData: prefs,
      noteFields: false,
      citeData: vi.fn().mockResolvedValue({ ok: true, value: [smith] }),
      refresh: vi.fn().mockRejectedValue(new Error('ECONNREFUSED')),
    })
    expect(result).toEqual({ ok: true, refreshed: false })
    expect(hasZoteroFields(ed)).toBe(true)
  })

  it('names the fix when the local API is off, and inserts nothing', async () => {
    const ed = makeEditor()
    const result = await citeFromZotero(ed, ['ABCD2345'], {
      documentData: '',
      noteFields: false,
      citeData: vi.fn().mockResolvedValue({ ok: false, errorCode: 'off', error: 'off' }),
      refresh: vi.fn(),
    })
    expect(result).toEqual({ ok: false, error: 'zoteroLocalApiOff' })
    expect(hasZoteroFields(ed)).toBe(false)
    expect(zoteroLocalErrorKey('unreachable')).toBe('zoteroConnectionError')
    expect(zoteroLocalErrorKey('bad-response')).toBe('zoteroOperationError')
  })

  it('does not ask Zotero for a read-only document', async () => {
    const citeData = vi.fn()
    const result = await citeFromZotero(makeEditor('x', false), ['ABCD2345'], {
      documentData: '',
      noteFields: false,
      citeData,
      refresh: vi.fn(),
    })
    expect(result).toEqual({ ok: false, error: 'zoteroCannotInsert' })
    expect(citeData).not.toHaveBeenCalled()
  })
})

describe('Refresh on Open', () => {
  const ready = {
    enabled: true,
    zoteroInstalled: true,
    editable: true,
    hasDocumentData: true,
    hasFields: true,
    noteFields: false,
  }

  it('refreshes an editable document with Zotero citations and preferences', () => {
    expect(shouldAutoRefresh(ready)).toBe(true)
  })

  it.each([
    ['turned off', { enabled: false }],
    ['Zotero not installed', { zoteroInstalled: false }],
    ['read-only', { editable: false }],
    ['no Zotero preferences', { hasDocumentData: false }],
    ['no Zotero fields', { hasFields: false }],
    ['citations in notes', { noteFields: true }],
  ])('skips a document: %s', (_, over) => {
    expect(shouldAutoRefresh({ ...ready, ...over })).toBe(false)
  })

  it('is on until turned off, and remembers the choice', () => {
    expect(loadZoteroAutoRefresh()).toBe(true)
    saveZoteroAutoRefresh(false)
    expect(loadZoteroAutoRefresh()).toBe(false)
    saveZoteroAutoRefresh(true)
    expect(loadZoteroAutoRefresh()).toBe(true)
  })
})

describe('library pane collections', () => {
  it('lists parents before their children, nested by depth', () => {
    const tree = collectionTree([
      { key: 'AAAAAAAA', name: 'Methods', parent: null },
      { key: 'BBBBBBBB', name: 'Surveys', parent: 'AAAAAAAA' },
      { key: 'CCCCCCCC', name: 'Theory', parent: null },
      { key: 'DDDDDDDD', name: 'Orphan', parent: 'ZZZZZZZZ' },
    ])
    expect(tree.map(({ collection, depth }) => `${depth}:${collection.name}`)).toEqual([
      '0:Methods',
      '1:Surveys',
      '0:Theory',
      '0:Orphan',
    ])
  })
})
