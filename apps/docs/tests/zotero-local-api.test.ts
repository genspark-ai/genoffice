import { describe, expect, it, vi } from 'vitest'
import {
  listZoteroCollectionItems,
  listZoteroCollections,
  searchZoteroItems,
  zoteroCiteData,
} from '../src/main/zotero-local-api'

function reply(body: unknown, init: { status?: number; total?: number } = {}): Response {
  return new Response(JSON.stringify(body), {
    status: init.status ?? 200,
    headers: {
      'Content-Type': 'application/json',
      ...(init.total !== undefined ? { 'Total-Results': String(init.total) } : {}),
    },
  })
}

const article = {
  key: 'ABCD2345',
  library: { id: 18084754 },
  meta: { creatorSummary: 'Smith and Lee', parsedDate: '2020-05-01' },
  data: { itemType: 'journalArticle', title: 'Cognitive load in practice' },
}

describe('Zotero local API', () => {
  it('searches title, creators and year and keeps only citable items', async () => {
    const fetchImpl = vi
      .fn()
      .mockResolvedValue(
        reply([
          article,
          { key: 'PDFX2345', data: { itemType: 'attachment', title: 'Full text' } },
          { key: 'NOTE2345', data: { itemType: 'note' } },
        ]),
      )
    const result = await searchZoteroItems('smith 2020', 25, fetchImpl)
    expect(result).toEqual({
      ok: true,
      value: [
        {
          key: 'ABCD2345',
          libraryId: 18084754,
          title: 'Cognitive load in practice',
          creators: 'Smith and Lee',
          year: '2020',
          itemType: 'journalArticle',
        },
      ],
    })
    const url = String(fetchImpl.mock.calls[0][0])
    expect(url.startsWith('http://127.0.0.1:23119/api/users/0/items/top?')).toBe(true)
    expect(url).toContain('q=smith%202020')
    expect(url).toContain('qmode=titleCreatorYear')
  })

  it('asks nothing for an empty query', async () => {
    const fetchImpl = vi.fn()
    expect(await searchZoteroItems('   ', 25, fetchImpl)).toEqual({ ok: true, value: [] })
    expect(fetchImpl).not.toHaveBeenCalled()
  })

  it('tells the setting being off from Zotero not running', async () => {
    const off = await searchZoteroItems(
      'x',
      25,
      vi.fn().mockResolvedValue(reply({}, { status: 403 })),
    )
    expect(off).toMatchObject({ ok: false, errorCode: 'off' })
    const down = await searchZoteroItems(
      'x',
      25,
      vi.fn().mockRejectedValue(new TypeError('fetch failed')),
    )
    expect(down).toMatchObject({ ok: false, errorCode: 'unreachable' })
    const broken = await searchZoteroItems(
      'x',
      25,
      vi.fn().mockResolvedValue(reply({}, { status: 500 })),
    )
    expect(broken).toMatchObject({ ok: false, errorCode: 'bad-response' })
  })

  it('pages through the collections and sorts them by name', async () => {
    const page = (n: number, from: number) =>
      Array.from({ length: n }, (_, i) => ({
        key: `C${String(from + i).padStart(7, '2')}`,
        data: { name: `Collection ${String(1000 - from - i)}`, parentCollection: false },
      }))
    const fetchImpl = vi
      .fn()
      .mockResolvedValueOnce(reply(page(100, 0), { total: 130 }))
      .mockResolvedValueOnce(reply(page(30, 100), { total: 130 }))
    const result = await listZoteroCollections(fetchImpl)
    expect(fetchImpl).toHaveBeenCalledTimes(2)
    expect(String(fetchImpl.mock.calls[1][0])).toContain('start=100')
    if (!result.ok) throw new Error('expected collections')
    expect(result.value).toHaveLength(130)
    expect(result.value[0].name).toBe('Collection 1000')
    expect(result.value[1].name).toBe('Collection 871')
    expect(result.value.every((c) => c.parent === null)).toBe(true)
  })

  it('lists one collection, and refuses a key that is not one', async () => {
    const fetchImpl = vi.fn().mockResolvedValue(reply([article]))
    const result = await listZoteroCollectionItems('QRST2345', 'load', fetchImpl)
    expect(result.ok).toBe(true)
    const url = String(fetchImpl.mock.calls[0][0])
    expect(url).toContain('/collections/QRST2345/items/top?')
    expect(url).toContain('q=load')
    const bad = await listZoteroCollectionItems('../items', '', fetchImpl)
    expect(bad).toMatchObject({ ok: false, errorCode: 'bad-request' })
    expect(fetchImpl).toHaveBeenCalledTimes(1)
  })

  it('reads the CSL-JSON and citation of each item, in the order asked', async () => {
    const fetchImpl = vi.fn().mockResolvedValue(
      reply([
        {
          key: 'WXYZ6789',
          library: { id: 7 },
          csljson: JSON.stringify([{ id: '7/WXYZ6789', type: 'book' }]),
          citation: '<span>(Lee, 2021)</span>',
        },
        {
          key: 'ABCD2345',
          library: { id: 7 },
          csljson: { id: '7/ABCD2345', type: 'article-journal' },
          citation: '(Smith, 2020)',
        },
      ]),
    )
    const result = await zoteroCiteData(['ABCD2345', 'WXYZ6789'], 'ieee', fetchImpl)
    expect(result).toEqual({
      ok: true,
      value: [
        {
          key: 'ABCD2345',
          libraryId: 7,
          itemData: { id: '7/ABCD2345', type: 'article-journal' },
          citation: '(Smith, 2020)',
        },
        {
          key: 'WXYZ6789',
          libraryId: 7,
          itemData: { id: '7/WXYZ6789', type: 'book' },
          citation: '<span>(Lee, 2021)</span>',
        },
      ],
    })
    const url = String(fetchImpl.mock.calls[0][0])
    expect(url).toContain('itemKey=ABCD2345,WXYZ6789')
    expect(url).toContain('include=csljson,citation')
    expect(url).toContain('style=ieee')
  })

  it('falls back to apa for a style id that is not one, and reports missing items', async () => {
    const fetchImpl = vi.fn().mockResolvedValue(reply([]))
    const result = await zoteroCiteData(['ABCD2345'], 'http://evil/../x', fetchImpl)
    expect(String(fetchImpl.mock.calls[0][0])).toContain('style=apa')
    expect(result).toMatchObject({ ok: false, errorCode: 'bad-response' })
    expect(await zoteroCiteData(['nope'], 'apa', fetchImpl)).toMatchObject({
      ok: false,
      errorCode: 'bad-request',
    })
  })
})
