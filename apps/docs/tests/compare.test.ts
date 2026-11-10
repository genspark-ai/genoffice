import { afterEach, describe, expect, it, vi } from 'vitest'
import { parseDocx, type Block } from '@genoffice/docx-engine'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import {
  blockTexts,
  compareParagraphs,
  editorBlockTexts,
  summarize,
} from '../src/renderer/editor/compare'
import { compareWithFile, type ReviewContext } from '../src/renderer/review-actions'

afterEach(() => vi.unstubAllGlobals())

describe('editorBlockTexts', () => {
  it('includes an unsaved paragraph on the current-document side', () => {
    const saved: Block[] = [
      {
        id: 'saved',
        type: 'paragraph',
        docxIndex: 0,
        originalXml: null,
        runs: [{ text: 'Saved' }],
      },
    ]
    const live = {
      type: 'doc',
      content: [
        { type: 'docParagraph', content: [{ type: 'text', text: 'Saved' }] },
        { type: 'docParagraph', content: [{ type: 'text', text: 'Unsaved' }] },
      ],
    }

    expect(editorBlockTexts(live, saved)).toEqual(['Saved', 'Unsaved'])
    expect(compareParagraphs(editorBlockTexts(live, saved), blockTexts(saved))).toContainEqual({
      kind: 'removed',
      left: 'Unsaved',
    })
  })

  it('keeps protected block previews and inline break text comparable', () => {
    const saved: Block[] = [
      { id: 'image', type: 'image', docxIndex: 0, originalXml: null, previewText: 'Chart' },
      { id: 'line', type: 'paragraph', docxIndex: 1, originalXml: null, runs: [{ text: 'A\nB' }] },
    ]
    const live = {
      type: 'doc',
      content: [
        { type: 'docProtected', attrs: { previewText: 'Chart' } },
        {
          type: 'docParagraph',
          content: [
            { type: 'text', text: 'A' },
            { type: 'hardBreak' },
            { type: 'text', text: 'B' },
          ],
        },
      ],
    }

    expect(editorBlockTexts(live, saved)).toEqual(blockTexts(saved))
  })
})

it('compares the live unsaved document after picking the original DOCX', async () => {
  const bytes = await buildDocx({ bodyXml: '<w:p><w:r><w:t>Saved</w:t></w:r></w:p>' })
  const parsed = await parseDocx(bytes)
  vi.stubGlobal('desktop', {
    openDocx: async () => ({ name: 'original.docx', dataUrl: 'test://original' }),
  })
  vi.stubGlobal('fetch', async () => ({ ok: true, arrayBuffer: async () => bytes.buffer }))
  const setCompareResult = vi.fn()
  const ctx = {
    doc: { parsed },
    editor: {
      getJSON: () => ({
        type: 'doc',
        content: [
          { type: 'docParagraph', content: [{ type: 'text', text: 'Saved' }] },
          { type: 'docParagraph', content: [{ type: 'text', text: 'Unsaved' }] },
        ],
      }),
    },
    setCompareResult,
    setStatus: vi.fn(),
  } as unknown as ReviewContext

  await compareWithFile(ctx)

  expect(setCompareResult).toHaveBeenCalledWith({
    otherName: 'original.docx',
    entries: expect.arrayContaining([{ kind: 'removed', left: 'Unsaved' }]),
  })
})

describe('compareParagraphs', () => {
  it('reports identical documents as all same', () => {
    const entries = compareParagraphs(['a', 'b'], ['a', 'b'])
    expect(entries.every((e) => e.kind === 'same')).toBe(true)
    expect(summarize(entries)).toEqual({ added: 0, removed: 0, changed: 0 })
  })

  it('detects an added paragraph', () => {
    const entries = compareParagraphs(['a', 'c'], ['a', 'b', 'c'])
    expect(entries.map((e) => e.kind)).toEqual(['same', 'added', 'same'])
    expect(entries[1].right).toBe('b')
  })

  it('detects a removed paragraph', () => {
    const entries = compareParagraphs(['a', 'b', 'c'], ['a', 'c'])
    expect(entries.map((e) => e.kind)).toEqual(['same', 'removed', 'same'])
    expect(entries[1].left).toBe('b')
  })

  it('merges adjacent remove+add into changed', () => {
    const entries = compareParagraphs(
      ['title', 'old content', 'ending'],
      ['title', 'new content', 'ending'],
    )
    expect(entries.map((e) => e.kind)).toEqual(['same', 'changed', 'same'])
    expect(entries[1]).toMatchObject({ left: 'old content', right: 'new content' })
  })

  it('handles empty documents', () => {
    expect(compareParagraphs([], [])).toEqual([])
    expect(compareParagraphs([], ['x'])[0].kind).toBe('added')
    expect(compareParagraphs(['x'], [])[0].kind).toBe('removed')
  })
})
