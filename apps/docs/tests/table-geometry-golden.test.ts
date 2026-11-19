/**
 * Golden pin for table width geometry (table-geometry.ts): every table of the
 * fixture corpus converts to the same PM width attrs and serialises to the same
 * table-level CSS (width / margins / colgroup) as before the width rules moved
 * out of convert.ts, extensions.ts and protected-render.ts. Declaration order
 * is not part of the pin. Hashes only; regenerate with
 * TABLE_GEOMETRY_GOLDEN_UPDATE=1 after an intentional change (then prettier
 * the golden directory). TABLE_GEOMETRY_GOLDEN_DUMP=<dir> writes the hashed
 * strings per fixture; TABLE_GEOMETRY_CORPUS=<dir> adds every .docx of a
 * directory outside the repo to the dump (never to the committed goldens).
 */
import { createHash } from 'node:crypto'
import { existsSync, mkdirSync, readdirSync, readFileSync, writeFileSync } from 'node:fs'
import { basename, dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { Editor } from '@tiptap/core'
import { DOMSerializer } from '@tiptap/pm/model'
import { parseDocx, readSections, type ParsedDoc, type TableModel } from '@genoffice/docx-engine'
import { describe, expect, it } from 'vitest'
import { blocksToPmDoc, pmDocOptions } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { renderTableSpec } from '../src/renderer/editor/protected-render'

;(globalThis as { CSS?: unknown }).CSS ??= { escape: (s: string) => s }

const HERE = dirname(fileURLToPath(import.meta.url))
const REPO = join(HERE, '../../..')
const CORPUS_DIR = join(HERE, 'pagination-corpus/docx')
const TRUTHS_DIR = join(REPO, 'packages/docx-engine/tests/fixtures/truths')
const GOLDEN_DIR = join(HERE, 'fixtures/table-geometry-golden')
const UPDATE = process.env.TABLE_GEOMETRY_GOLDEN_UPDATE === '1'
const DUMP_DIR = process.env.TABLE_GEOMETRY_GOLDEN_DUMP
const EXTERNAL = process.env.TABLE_GEOMETRY_CORPUS

const WIDTH_ATTRS = [
  'widthPx',
  'widthPct',
  'colWidthsPct',
  'tblAlign',
  'indentTwips',
  'tblAutoFit',
  'tblFixedLayout',
  'tblDxaWidth',
  'tblFloat',
  'tblFloatWidthPx',
  'bidiVisual',
]

interface Golden {
  source: string
  tables: number
  hashes: string[]
  staticTables: number
  staticHashes: string[]
}

interface Fixture {
  name: string
  source: string
  external?: boolean
}

function docxIn(dir: string, source: (f: string) => string, external = false): Fixture[] {
  if (!existsSync(dir)) return []
  return readdirSync(dir)
    .filter((f) => f.endsWith('.docx'))
    .sort()
    .map((f) => ({ name: basename(f, '.docx'), source: source(f), external }))
}

function fixtures(): Fixture[] {
  const pinned = existsSync(GOLDEN_DIR)
    ? readdirSync(GOLDEN_DIR)
        .filter((f) => f.endsWith('.json'))
        .sort()
        .map((f) => ({
          name: basename(f, '.json'),
          source: (JSON.parse(readFileSync(join(GOLDEN_DIR, f), 'utf8')) as Golden).source,
        }))
    : []
  const external = EXTERNAL ? docxIn(EXTERNAL, (f) => join(EXTERNAL, f), true) : []
  if (pinned.length > 0 && !UPDATE) return [...pinned, ...external]
  const corpus = docxIn(CORPUS_DIR, (f) => `apps/docs/tests/pagination-corpus/docx/${f}`)
  const truths = readdirSync(TRUTHS_DIR, { withFileTypes: true })
    .filter((d) => d.isDirectory() && existsSync(join(TRUTHS_DIR, d.name, 'probe.docx')))
    .map((d) => ({
      name: `truth-${d.name}`,
      source: `packages/docx-engine/tests/fixtures/truths/${d.name}/probe.docx`,
    }))
  const all = new Map([...pinned, ...corpus, ...truths].map((f) => [f.name, f]))
  return [...[...all.values()].sort((a, b) => a.name.localeCompare(b.name)), ...external]
}

const sha = (s: string) => createHash('sha1').update(s).digest('hex').slice(0, 16)

const sortedStyle = (style: string | null) =>
  (style ?? '')
    .split(';')
    .map((d) => d.trim())
    .filter(Boolean)
    .sort()
    .join(';')

function tableCss(table: HTMLElement): string {
  const cols = Array.from(table.querySelectorAll(':scope > colgroup > col'))
    .map((c) => sortedStyle(c.getAttribute('style')))
    .join(',')
  return `${sortedStyle(table.getAttribute('style'))} | cols ${cols}`
}

function pmTableStrings(parsed: ParsedDoc): string[] {
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks, readSections(parsed), pmDocOptions(parsed)) as never,
  })
  try {
    const serializer = DOMSerializer.fromSchema(editor.schema)
    const out: string[] = []
    editor.state.doc.descendants((node) => {
      if (node.type.name !== 'docTable') return true
      const dom = serializer.serializeNode(node, { document }) as HTMLElement
      const attrs = Object.fromEntries(WIDTH_ATTRS.map((k) => [k, node.attrs[k] ?? null]))
      out.push(`${JSON.stringify(attrs)} | ${tableCss(dom)}`)
      return true
    })
    return out
  } finally {
    editor.destroy()
  }
}

function staticTableStrings(parsed: ParsedDoc): string[] {
  const widthKeys: Array<keyof TableModel> = [
    'colWidthsTwips',
    'widthPct',
    'align',
    'indentTwips',
    'autoFit',
    'fixedLayout',
    'dxaWidth',
    'autoLayout',
    'layoutGrid',
  ]
  return parsed.blocks.flatMap((b) => {
    if (!b.table) return []
    const { dom } = DOMSerializer.renderSpec(document, renderTableSpec(b.table) as never)
    const tables = [
      dom as HTMLElement,
      ...Array.from((dom as HTMLElement).querySelectorAll('table')),
    ]
    const model = Object.fromEntries(widthKeys.map((k) => [k, b.table![k] ?? null]))
    return tables.map((t, i) => `${i === 0 ? JSON.stringify(model) : 'nested'} | ${tableCss(t)}`)
  })
}

describe('table geometry golden (PM width attrs + emitted table CSS)', () => {
  const all = fixtures()
  it('covers the fixture corpus', () => {
    expect(all.filter((f) => !f.external).length).toBeGreaterThanOrEqual(50)
  })
  for (const fx of all) {
    it(fx.name, async () => {
      const path = fx.external ? fx.source : join(REPO, fx.source)
      const parsed = await parseDocx(new Uint8Array(readFileSync(path)))
      const pm = pmTableStrings(parsed)
      const stat = staticTableStrings(parsed)
      if (DUMP_DIR) {
        mkdirSync(DUMP_DIR, { recursive: true })
        writeFileSync(join(DUMP_DIR, `${fx.name}.pm.txt`), pm.map((x) => `${x}\n`).join(''))
        writeFileSync(join(DUMP_DIR, `${fx.name}.static.txt`), stat.map((x) => `${x}\n`).join(''))
      }
      if (fx.external) return
      const file = join(GOLDEN_DIR, `${fx.name}.json`)
      if (UPDATE) {
        mkdirSync(GOLDEN_DIR, { recursive: true })
        const golden: Golden = {
          source: fx.source,
          tables: pm.length,
          hashes: pm.map(sha),
          staticTables: stat.length,
          staticHashes: stat.map(sha),
        }
        writeFileSync(file, `${JSON.stringify(golden, null, 2)}\n`)
      }
      const golden = JSON.parse(readFileSync(file, 'utf8')) as Golden
      expect(pm.length, 'table count').toBe(golden.tables)
      expect(pm.map(sha)).toEqual(golden.hashes)
      expect(stat.length, 'static table count').toBe(golden.staticTables)
      expect(stat.map(sha)).toEqual(golden.staticHashes)
    })
  }
})
