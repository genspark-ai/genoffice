import { readFileSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import JSZip from 'jszip'
import { expect } from 'vitest'
import { run } from './helpers'

/** ref → { value, formula? } of one worksheet part, shared and inline strings resolved */
export async function cells(
  path: string,
  sheetPart: string,
): Promise<Map<string, { value: string | number | null; formula?: string }>> {
  const zip = await JSZip.loadAsync(readFileSync(path))
  const shared = [
    ...((await zip.file('xl/sharedStrings.xml')?.async('string')) ?? '').matchAll(
      /<si>([\s\S]*?)<\/si>/g,
    ),
  ].map((m) => [...m[1]!.matchAll(/<t[^>]*>([\s\S]*?)<\/t>/g)].map((t) => t[1]).join(''))
  const xml = await zip.file(sheetPart)!.async('string')
  const out = new Map<string, { value: string | number | null; formula?: string }>()
  for (const m of xml.matchAll(/<c r="([A-Z]+\d+)"([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g)) {
    const type = /\bt="([^"]+)"/.exec(m[2]!)?.[1]
    const body = m[3] ?? ''
    const raw = /<v>([\s\S]*?)<\/v>/.exec(body)?.[1]
    const formula = /<f>([\s\S]*?)<\/f>/.exec(body)?.[1]
    const value =
      type === 's'
        ? (shared[Number(raw)] ?? null)
        : type === 'inlineStr'
          ? (/<t[^>]*>([\s\S]*?)<\/t>/.exec(body)?.[1] ?? null)
          : raw === undefined
            ? null
            : type === 'str' || type === 'e'
              ? raw
              : Number(raw)
    out.set(m[1]!, formula ? { value, formula } : { value })
  }
  return out
}

export async function book(dir: string, rows: unknown[][]): Promise<string> {
  const table = join(dir, 'table.json')
  writeFileSync(table, JSON.stringify(rows))
  const out = join(dir, 'book.xlsx')
  expect((await run(['create', '--type', 'xlsx', '--from', table, '--out', out])).code).toBe(0)
  return out
}

export async function part(path: string, name: string): Promise<string> {
  const zip = await JSZip.loadAsync(readFileSync(path))
  return zip.file(name)!.async('string')
}
