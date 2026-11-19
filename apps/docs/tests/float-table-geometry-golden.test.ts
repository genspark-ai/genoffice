/**
 * Golden pin for floating-table placement (w:tblpPr through anchor-geometry.ts):
 * per floating table the PM float attrs and the emitted open tag (class,
 * style, data-tblp-*), the nested-float CSS of every nested model and the
 * header/footer row placement, over the pagination corpus, the truth probes
 * and two synthetic matrices. Regenerate with FLOAT_TABLE_GOLDEN_UPDATE=1
 * after an intentional placement change; FLOAT_TABLE_GOLDEN_DUMP=<dir> writes
 * the hashed lines per fixture; FLOAT_TABLE_CORPUS=<dir> adds every .docx of
 * an outside corpus (dumped, never pinned).
 */
import { createHash } from 'node:crypto'
import { existsSync, mkdirSync, readdirSync, readFileSync, writeFileSync } from 'node:fs'
import { basename, dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { Editor } from '@tiptap/core'
import { DOMSerializer } from '@tiptap/pm/model'
import { parseDocx, readSections, type ParsedDoc, type TableModel } from '@genoffice/docx-engine'
import { describe, expect, it } from 'vitest'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc, pmDocOptions } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'
import {
  hfRowFloatLeftPx,
  hfRowGeomOf,
  hfRowStyle,
  hfStripGeom,
} from '../src/renderer/editor/hf-dom'
import { nestedFloatCss } from '../src/renderer/editor/protected-render'
import { floatTableFlipMatrixXml, floatTableMatrixXml } from './helpers/float-table-matrix'

;(globalThis as { CSS?: unknown }).CSS ??= { escape: (s: string) => s }

const HERE = dirname(fileURLToPath(import.meta.url))
const REPO = join(HERE, '../../..')
const CORPUS_DIR = join(HERE, 'pagination-corpus/docx')
const TRUTHS_DIR = join(REPO, 'packages/docx-engine/tests/fixtures/truths')
const GOLDEN_DIR = join(HERE, 'fixtures/float-table-geometry-golden')
const UPDATE = process.env.FLOAT_TABLE_GOLDEN_UPDATE === '1'
const DUMP_DIR = process.env.FLOAT_TABLE_GOLDEN_DUMP
const EXTERNAL = process.env.FLOAT_TABLE_CORPUS

const FLOAT_ATTRS = [
  'tblFloat',
  'tblFloatSource',
  'tblFloatSuppressed',
  'tblFloatXTwips',
  'tblFloatYTwips',
  'tblFloatHorzAnchor',
  'tblFloatVertAnchor',
  'tblFloatXSpec',
  'tblFloatYSpec',
  'tblFloatDistance',
  'tblFloatWidthPx',
  'widthPx',
  'indentTwips',
  'tblAlign',
]

const SYNTHETIC: Record<string, () => string> = {
  'synthetic-float-tables': floatTableMatrixXml,
  'synthetic-float-flips': floatTableFlipMatrixXml,
}

interface Golden {
  source: string
  editor: string[]
  nested: string[]
  hf: string[]
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
        .filter((f) => !f.source.startsWith('synthetic:'))
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

function openTag(el: Element): string {
  const attrs = Array.from(el.attributes)
    .map(({ name, value }) => {
      if (name === 'class') value = value.split(/\s+/).filter(Boolean).sort().join(' ')
      if (name === 'style')
        value = value
          .split(';')
          .map((d) => d.trim())
          .filter(Boolean)
          .sort()
          .join(';')
      return `${name}="${value}"`
    })
    .sort()
  return `<${el.tagName.toLowerCase()} ${attrs.join(' ')}>`
}

function editorLines(parsed: ParsedDoc): string[] {
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks, readSections(parsed), pmDocOptions(parsed)) as never,
  })
  try {
    const serializer = DOMSerializer.fromSchema(editor.schema)
    const out: string[] = []
    editor.state.doc.descendants((node) => {
      if (node.type.name !== 'docTable' || !node.attrs.tblFloatSource) return true
      const dom = serializer.serializeNode(node, { document }) as HTMLElement
      const attrs = Object.fromEntries(FLOAT_ATTRS.map((k) => [k, node.attrs[k] ?? null]))
      out.push(`${JSON.stringify(attrs)} | ${openTag(dom)}`)
      return true
    })
    return out
  } finally {
    editor.destroy()
  }
}

function nestedLines(parsed: ParsedDoc): string[] {
  const out: string[] = []
  const visit = (t: TableModel, path: string) => {
    const css = nestedFloatCss(t)
    if (css) out.push(`${path} ${css.join(';')}`)
    t.rows.forEach((row, r) =>
      row.forEach((cell, c) =>
        (cell.nestedTables ?? []).forEach((n, i) => visit(n, `${path}/${r}.${c}.${i}`)),
      ),
    )
  }
  parsed.blocks.forEach((b, i) => {
    if (b.table) visit(b.table, `${i}`)
  })
  return out
}

function hfLines(parsed: ParsedDoc): string[] {
  const section = readSections(parsed)[0]
  const strip = section ? hfStripGeom(section.settings) : undefined
  return Object.entries(parsed.hfParts ?? {})
    .sort(([a], [b]) => a.localeCompare(b))
    .flatMap(([name, part]) => {
      const kind = /footer|ftr/i.test(name) ? 'footer' : 'header'
      const geom = hfRowGeomOf(kind, strip)
      return (part.paras ?? []).flatMap((p, i) => {
        if (!p.row?.float) return []
        const left = hfRowFloatLeftPx(p.row, geom)
        return [
          `${name}#${i} ${JSON.stringify(p.row.float)} w=${p.row.widthTwips ?? null} left=${left == null ? null : left.toFixed(2)} ${JSON.stringify(hfRowStyle(p.row, geom, i === 0))}`,
        ]
      })
    })
}

function dump(name: string, kind: string, lines: string[]): void {
  if (!DUMP_DIR) return
  mkdirSync(DUMP_DIR, { recursive: true })
  writeFileSync(join(DUMP_DIR, `${name}.${kind}.txt`), lines.map((x) => `${x}\n`).join(''))
}

async function check(f: Fixture, parsed: ParsedDoc): Promise<void> {
  const editor = editorLines(parsed)
  const nested = nestedLines(parsed)
  const hf = hfLines(parsed)
  dump(f.name, 'editor', editor)
  dump(f.name, 'nested', nested)
  dump(f.name, 'hf', hf)
  if (f.external) return
  const got = { editor: editor.map(sha), nested: nested.map(sha), hf: hf.map(sha) }
  const file = join(GOLDEN_DIR, `${f.name}.json`)
  if (UPDATE) {
    // only documents that carry a floating table are pinned
    if (editor.length + nested.length + hf.length === 0) return
    mkdirSync(GOLDEN_DIR, { recursive: true })
    writeFileSync(file, `${JSON.stringify({ source: f.source, ...got }, null, 2)}\n`)
  }
  if (!existsSync(file)) {
    expect(got, `${f.name} gained floating tables: regenerate the golden`).toEqual({
      editor: [],
      nested: [],
      hf: [],
    })
    return
  }
  const golden = JSON.parse(readFileSync(file, 'utf8')) as Golden
  expect(got.editor, 'floating tables').toEqual(golden.editor)
  expect(got.nested, 'nested floats').toEqual(golden.nested)
  expect(got.hf, 'header/footer float rows').toEqual(golden.hf)
}

describe('float-table geometry golden (w:tblpPr placement of every floating table)', () => {
  for (const [name, body] of Object.entries(SYNTHETIC)) {
    it(name, async () => {
      const parsed = await parseDocx(await buildDocx({ bodyXml: body(), withImage: true }))
      expect(editorLines(parsed).length).toBeGreaterThan(0)
      await check({ name, source: `synthetic:${name}` }, parsed)
    })
  }
  for (const f of fixtures()) {
    it(f.name, async () => {
      const path = f.external ? f.source : join(REPO, f.source)
      if (!existsSync(path)) return
      await check(f, await parseDocx(new Uint8Array(readFileSync(path))))
    })
  }
})
