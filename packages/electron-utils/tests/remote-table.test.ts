import { describe, expect, it, vi } from 'vitest'

import {
  detectDelimiter,
  fetchRemoteTable,
  parseDelimitedRow,
  parseDelimitedTable,
  parseHtmlTable,
} from '../src/remote-table'

vi.mock('node:dns/promises', () => ({
  lookup: async () => [{ address: '93.184.215.14', family: 4 }],
}))

const htmlResponse = (body: string, contentType = 'text/html') =>
  new Response(body, { status: 200, headers: { 'content-type': contentType } })

describe('parseHtmlTable', () => {
  it('reads a table with a header row', () => {
    const table = parseHtmlTable(`
      <html><body>
        <table>
          <tr><th>Currency</th><th>Rate</th></tr>
          <tr><td>EUR</td><td>1.09</td></tr>
          <tr><td>USD</td><td>1.00</td></tr>
        </table>
      </body></html>`)
    expect(table?.headers).toEqual(['Currency', 'Rate'])
    expect(table?.rows).toEqual([
      ['EUR', '1.09'],
      ['USD', '1.00'],
    ])
    expect(table?.source).toBe('html')
  })

  it('keeps cell text that carries markup', () => {
    const table = parseHtmlTable(`
      <table>
        <tr><th>Name</th></tr>
        <tr><td><b>Bold</b> and <i>italic</i></td></tr>
      </table>`)
    expect(table?.rows[0]?.[0]).toBe('Bold and italic')
  })

  it('takes the widest table when the page leads with a layout table', () => {
    const table = parseHtmlTable(`
      <table><tr><td></td></tr></table>
      <table>
        <tr><th>a</th><th>b</th></tr>
        <tr><td>1</td><td>2</td></tr>
        <tr><td>3</td><td>4</td></tr>
      </table>`)
    expect(table?.headers).toEqual(['a', 'b'])
    expect(table?.rows).toHaveLength(2)
  })

  it('returns null for a page without a table', () => {
    expect(parseHtmlTable('<html><body><p>hello</p></body></html>')).toBeNull()
  })
})

describe('delimited parsing', () => {
  it('reads quoted cells, escaped quotes and embedded commas', () => {
    expect(parseDelimitedRow('"Doe, John","says ""hi""",30', ',')).toEqual([
      'Doe, John',
      'says "hi"',
      '30',
    ])
  })

  it('picks the delimiter the file uses', () => {
    expect(detectDelimiter('a,b,c\n1,2,3')).toBe(',')
    expect(detectDelimiter('a\tb\tc\n1\t2\t3')).toBe('\t')
    expect(detectDelimiter('a;b;c\n1;2;3')).toBe(';')
  })

  it('ignores delimiters inside quotes when sniffing', () => {
    expect(detectDelimiter('"a,b";c\n"x;y";z')).toBe(';')
  })

  it('treats the first row as the header', () => {
    const table = parseDelimitedTable('Currency,Rate\nEUR,1.09\nUSD,1.00')
    expect(table?.headers).toEqual(['Currency', 'Rate'])
    expect(table?.rows).toEqual([
      ['EUR', '1.09'],
      ['USD', '1.00'],
    ])
  })

  it('skips a byte order mark', () => {
    const table = parseDelimitedTable('﻿a,b\n1,2')
    expect(table?.headers).toEqual(['a', 'b'])
  })
})

describe('fetchRemoteTable', () => {
  it('fetches and reads an HTML table', async () => {
    const fetchImpl = vi
      .fn()
      .mockResolvedValue(
        htmlResponse('<table><tr><th>Currency</th></tr><tr><td>EUR</td></tr></table>'),
      )
    const table = await fetchRemoteTable('https://example.com/rates.html', { fetchImpl })
    expect(table?.headers).toEqual(['Currency'])
    expect(table?.rows).toEqual([['EUR']])
    expect(table?.url).toBe('https://example.com/rates.html')
  })

  it('reads a CSV answer', async () => {
    const fetchImpl = vi.fn().mockResolvedValue(htmlResponse('Currency,Rate\nEUR,1.09', 'text/csv'))
    const table = await fetchRemoteTable('https://example.com/rates.csv', { fetchImpl })
    expect(table?.source).toBe('csv')
    expect(table?.headers).toEqual(['Currency', 'Rate'])
  })

  it('gives up on a refused address instead of retrying', async () => {
    const fetchImpl = vi.fn().mockResolvedValue(new Response('nope', { status: 404 }))
    expect(await fetchRemoteTable('https://example.com/missing', { fetchImpl })).toBeNull()
  })

  it('gives up when the answer is not a table', async () => {
    const fetchImpl = vi.fn().mockResolvedValue(htmlResponse('<p>hello</p>'))
    expect(await fetchRemoteTable('https://example.com/page', { fetchImpl })).toBeNull()
  })

  it('still reads a CSV answer that arrives as text/html', async () => {
    const fetchImpl = vi
      .fn()
      .mockResolvedValue(htmlResponse('Currency,Rate\nEUR,1.09', 'text/html'))
    const table = await fetchRemoteTable('https://example.com/rates.csv', { fetchImpl })
    expect(table?.headers).toEqual(['Currency', 'Rate'])
  })
})
