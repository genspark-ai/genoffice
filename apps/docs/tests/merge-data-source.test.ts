import { describe, expect, it } from 'vitest'
import { mkdtemp, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { csvToXlsxBuffer } from '@genoffice/xlsx-gateway/gateway/csv-import'
import { readMergeDataSource } from '../src/main/merge-data-source'

/**
 * The data source of the mail merge: a local .xlsx or .csv read as a field
 * list plus records. genoffice#2097.
 */

async function tempDir(): Promise<string> {
  return mkdtemp(join(tmpdir(), 'genoffice-merge-'))
}

describe('readMergeDataSource', () => {
  it('reads a CSV with a UTF-8 BOM and quoted fields', async () => {
    const dir = await tempDir()
    const path = join(dir, 'recipients.csv')
    await writeFile(
      path,
      '﻿Name,Note\r\n"Ada Lovelace","a, b"\r\n"Grace Hopper","line1\nline2"\r\n',
      'utf8',
    )
    const source = await readMergeDataSource(path)
    expect(source.name).toBe('recipients.csv')
    expect(source.fields).toEqual(['Name', 'Note'])
    expect(source.records).toEqual([
      { Name: 'Ada Lovelace', Note: 'a, b' },
      { Name: 'Grace Hopper', Note: 'line1\nline2' },
    ])
  })

  it('reads the first sheet of an xlsx', async () => {
    const dir = await tempDir()
    const path = join(dir, 'recipients.xlsx')
    await writeFile(path, await csvToXlsxBuffer('First Name,Age\nAda,36\nGrace,45\n'))
    const source = await readMergeDataSource(path)
    expect(source.name).toBe('recipients.xlsx')
    expect(source.fields).toEqual(['First Name', 'Age'])
    expect(source.records).toEqual([
      { 'First Name': 'Ada', Age: '36' },
      { 'First Name': 'Grace', Age: '45' },
    ])
  })

  it('names an empty header column and separates a repeated one', async () => {
    const dir = await tempDir()
    const path = join(dir, 'recipients.csv')
    await writeFile(path, 'Name,,Name\nAda,x,y\n', 'utf8')
    const source = await readMergeDataSource(path)
    expect(source.fields).toEqual(['Name', 'Column2', 'Name2'])
    expect(source.records).toEqual([{ Name: 'Ada', Column2: 'x', Name2: 'y' }])
  })

  it('refuses a format it cannot read', async () => {
    const dir = await tempDir()
    const path = join(dir, 'recipients.txt')
    await writeFile(path, 'Name\nAda\n', 'utf8')
    await expect(readMergeDataSource(path)).rejects.toThrow(/unsupported data source/)
  })
})
