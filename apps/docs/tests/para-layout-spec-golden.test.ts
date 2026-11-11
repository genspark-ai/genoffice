/**
 * Golden pin for the paragraph line-box attrs (para-layout-spec.ts): every
 * paragraph-like block of the fixture corpus serialises to the same static HTML
 * as before the helpers moved out of extensions.ts (fromPmNode, through the
 * editor schema's renderHTML), and every read-only cell / textbox paragraph
 * (fromRuns, through renderTableSpec / renderTextboxSpec) to the same canonical
 * markup: attribute, class and declaration order are not part of that pin.
 * Header/footer strip paragraphs (makeGapHfEl, stacked and layout-table cell
 * lines) and footnote/endnote rows (makeGapNotesEl) are pinned the same way.
 * Hashes only, so the goldens stay small and carry no fixture text. Regenerate
 * with PARA_LAYOUT_GOLDEN_UPDATE=1 after an intentional rendering change (then
 * run prettier on the golden directory); the committed goldens define the
 * fixture set once they exist.
 * PARA_LAYOUT_GOLDEN_DUMP=<dir> also writes the hashed markup per fixture, so
 * two dumps diff to the paragraphs an intentional change moved.
 */
import { createHash } from 'node:crypto'
import { existsSync, mkdirSync, readdirSync, readFileSync, writeFileSync } from 'node:fs'
import { basename, dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { Editor } from '@tiptap/core'
import { DOMSerializer } from '@tiptap/pm/model'
import { parseDocx, readSections, type ParsedDoc } from '@genoffice/docx-engine'
import { describe, expect, it } from 'vitest'
import { blocksToPmDoc, pmDocOptions } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { makeGapHfEl } from '../src/renderer/editor/hf-dom'
import { makeGapNotesEl, noteRowMetrics } from '../src/renderer/editor/note-dom'
import { renderTableSpec, renderTextboxSpec } from '../src/renderer/editor/protected-render'
import { footnoteLineHeightPx, noteRunStyle, resolveNoteStyle } from '../src/renderer/line-metrics'

;(globalThis as { CSS?: unknown }).CSS ??= { escape: (s: string) => s }

const HERE = dirname(fileURLToPath(import.meta.url))
const REPO = join(HERE, '../../..')
const CORPUS_DIR = join(HERE, 'pagination-corpus/docx')
const TRUTHS_DIR = join(REPO, 'packages/docx-engine/tests/fixtures/truths')
const GOLDEN_DIR = join(HERE, 'fixtures/para-layout-golden')
const TRUTH_PROBES = 23
/** truth probes carrying header/footer parts (the corpus has none) */
const HF_TRUTH_PROBES = [
  'render-hf-anchor-reserve',
  'render-hf-blank-header-push',
  'render-hf-empty-footer-bottom',
  'render-hf-header-dist-push',
  'render-hf-inline-picture',
  'render-hf-numeric-style-size',
  'render-hf-rtl-neutral-size',
  'render-inherited-header-width',
]
const UPDATE = process.env.PARA_LAYOUT_GOLDEN_UPDATE === '1'
const DUMP_DIR = process.env.PARA_LAYOUT_GOLDEN_DUMP

const PARA_TYPES = new Set(['docParagraph', 'docHeading', 'docListItem'])
const RUN_PARA_SEL = '.doc-cell-para, .doc-textbox-para'
const HF_PARA_SEL = '.page-hf-para, .page-hf-cell-para'

interface Golden {
  source: string
  paragraphs: number
  hashes: string[]
  /** read-only cell paragraphs (renderTableSpec, nested tables and cell boxes included) */
  cellParagraphs?: number
  cellHashes?: string[]
  /** paragraphs of body-anchored textboxes (renderTextboxSpec) */
  boxParagraphs?: number
  boxHashes?: string[]
  /** header/footer strip paragraphs of every part (makeGapHfEl) */
  hfParagraphs?: number
  hfHashes?: string[]
  /** footnote and endnote rows (makeGapNotesEl) */
  noteParagraphs?: number
  noteHashes?: string[]
}

interface Fixture {
  name: string
  source: string
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
  if (pinned.length > 0 && !UPDATE) return pinned
  const corpus = readdirSync(CORPUS_DIR)
    .filter((f) => f.endsWith('.docx'))
    .sort()
    .map((f) => ({
      name: basename(f, '.docx'),
      source: `apps/docs/tests/pagination-corpus/docx/${f}`,
    }))
  const probes = readdirSync(TRUTHS_DIR, { withFileTypes: true })
    .filter((d) => d.isDirectory() && existsSync(join(TRUTHS_DIR, d.name, 'probe.docx')))
    .map((d) => d.name)
    .sort()
  // the pinned set is the fixture set; a regeneration only adds the HF probes
  // (and the first TRUTH_PROBES when no golden exists yet)
  const truthDirs =
    pinned.length > 0 ? HF_TRUTH_PROBES : [...probes.slice(0, TRUTH_PROBES), ...HF_TRUTH_PROBES]
  const truths = truthDirs.map((dir) => ({
    name: `truth-${dir}`,
    source: `packages/docx-engine/tests/fixtures/truths/${dir}/probe.docx`,
  }))
  const all = new Map([...pinned, ...corpus, ...truths].map((f) => [f.name, f]))
  return [...all.values()].sort((a, b) => a.name.localeCompare(b.name))
}

const sha = (html: string) => createHash('sha1').update(html).digest('hex').slice(0, 16)

type Hashed = { hashes: string[]; html: string[] }

function paragraphHashes(parsed: ParsedDoc): Hashed {
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks, readSections(parsed), pmDocOptions(parsed)) as never,
  })
  try {
    const serializer = DOMSerializer.fromSchema(editor.schema)
    const html: string[] = []
    editor.state.doc.descendants((node) => {
      if (!PARA_TYPES.has(node.type.name)) return true
      const dom = serializer.serializeNode(node, { document }) as HTMLElement
      html.push(dom.outerHTML)
      return true
    })
    return { hashes: html.map(sha), html }
  } finally {
    editor.destroy()
  }
}

/** outerHTML with attributes, class tokens and style declarations in sorted order */
function canonicalHtml(node: Node): string {
  if (node.nodeType === Node.TEXT_NODE) return node.textContent ?? ''
  if (node.nodeType !== Node.ELEMENT_NODE) return ''
  const el = node as Element
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
  const children = Array.from(el.childNodes).map(canonicalHtml).join('')
  const tag = el.tagName.toLowerCase()
  return `<${tag}${attrs.length ? ` ${attrs.join(' ')}` : ''}>${children}</${tag}>`
}

function runParagraphHtml(spec: unknown): string[] {
  const { dom } = DOMSerializer.renderSpec(document, spec as never)
  return Array.from((dom as HTMLElement).querySelectorAll(RUN_PARA_SEL)).map(canonicalHtml)
}

function cellHashes(parsed: ParsedDoc): Hashed {
  const html = parsed.blocks.flatMap((b) =>
    b.table ? runParagraphHtml(renderTableSpec(b.table)) : [],
  )
  return { hashes: html.map(sha), html }
}

/** the corpus holds no textboxes: every rich cell paragraph doubles as a
 *  textbox paragraph (same engine shape), so the textbox path sees real runs */
function boxHashes(parsed: ParsedDoc): Hashed {
  const html = parsed.blocks.flatMap((b) => [
    ...(b.textboxes ?? []).flatMap((box) => runParagraphHtml(renderTextboxSpec(box))),
    ...(b.table?.rows ?? []).flatMap((row) =>
      row.flatMap((cell) =>
        cell.richParas?.length
          ? runParagraphHtml(renderTextboxSpec({ paras: cell.richParas }))
          : [],
      ),
    ),
  ])
  return { hashes: html.map(sha), html }
}

/** strip paragraphs of every header/footer part, stacked and layout-table cell lines */
function hfHashes(parsed: ParsedDoc): Hashed {
  const html = Object.entries(parsed.hfParts ?? {})
    .sort(([a], [b]) => a.localeCompare(b))
    .flatMap(([name, part]) => {
      const el = makeGapHfEl({
        kind: /footer|ftr/i.test(name) ? 'footer' : 'header',
        value: { text: part.text, paras: part.paras },
        images: part.images,
        pageNo: 1,
        pageTotal: 1,
      })
      return Array.from(el.querySelectorAll(HF_PARA_SEL)).map(canonicalHtml)
    })
  return { hashes: html.map(sha), html }
}

/** footnote and endnote rows at the renderers' styles; the measured min-height
 *  and the localised tooltip are not part of the pin */
function noteHashes(parsed: ParsedDoc): Hashed {
  const docGrid = readSections(parsed)[0]?.settings.docGrid
  const items = [...(parsed.footnotes ?? []), ...(parsed.endnotes ?? [])].map((n, i) => {
    const style = {
      ...resolveNoteStyle(parsed, n.styleId, n.spacing),
      ...noteRunStyle(n.richParas),
    }
    return {
      no: i + 1,
      id: n.id,
      text: n.text,
      height: 0,
      ...(n.richParas ? { richParas: n.richParas } : {}),
      ...(n.noRefMark ? { noRefMark: true as const } : {}),
      ...noteRowMetrics(style, docGrid),
    }
  })
  if (items.length === 0) return { hashes: [], html: [] }
  const el = makeGapNotesEl(items, 0, 600, 100, footnoteLineHeightPx(docGrid), () => {})
  const html = Array.from(el.querySelectorAll<HTMLElement>('.page-gap-note')).map((row) => {
    row.style.removeProperty('min-height')
    row.removeAttribute('title')
    return canonicalHtml(row)
  })
  return { hashes: html.map(sha), html }
}

function dump(
  name: string,
  paras: Hashed,
  cells: Hashed,
  boxes: Hashed,
  hf: Hashed,
  notes: Hashed,
): void {
  if (!DUMP_DIR) return
  mkdirSync(DUMP_DIR, { recursive: true })
  for (const [kind, h] of [
    ['paras', paras],
    ['cells', cells],
    ['boxes', boxes],
    ['hf', hf],
    ['notes', notes],
  ] as const) {
    writeFileSync(join(DUMP_DIR, `${name}.${kind}.txt`), h.html.map((x) => `${x}\n`).join(''))
  }
}

describe('para-layout-spec golden (static HTML of every paragraph block)', () => {
  const all = fixtures()
  it('covers the fixture corpus', () => {
    expect(all.length).toBeGreaterThanOrEqual(50)
  })
  for (const fx of all) {
    it(fx.name, async () => {
      const parsed = await parseDocx(new Uint8Array(readFileSync(join(REPO, fx.source))))
      const paras = paragraphHashes(parsed)
      const cells = cellHashes(parsed)
      const boxes = boxHashes(parsed)
      const hf = hfHashes(parsed)
      const notes = noteHashes(parsed)
      dump(fx.name, paras, cells, boxes, hf, notes)
      const hashes = paras.hashes
      const file = join(GOLDEN_DIR, `${fx.name}.json`)
      if (UPDATE) {
        mkdirSync(GOLDEN_DIR, { recursive: true })
        const golden: Golden = {
          source: fx.source,
          paragraphs: hashes.length,
          hashes,
          cellParagraphs: cells.hashes.length,
          cellHashes: cells.hashes,
          boxParagraphs: boxes.hashes.length,
          boxHashes: boxes.hashes,
          hfParagraphs: hf.hashes.length,
          hfHashes: hf.hashes,
          noteParagraphs: notes.hashes.length,
          noteHashes: notes.hashes,
        }
        writeFileSync(file, `${JSON.stringify(golden, null, 2)}\n`)
      }
      const golden = JSON.parse(readFileSync(file, 'utf8')) as Golden
      expect(hashes.length, 'paragraph count').toBe(golden.paragraphs)
      expect(hashes).toEqual(golden.hashes)
      expect(cells.hashes.length, 'cell paragraph count').toBe(golden.cellParagraphs)
      expect(cells.hashes).toEqual(golden.cellHashes)
      expect(boxes.hashes.length, 'textbox paragraph count').toBe(golden.boxParagraphs)
      expect(boxes.hashes).toEqual(golden.boxHashes)
      expect(hf.hashes.length, 'strip paragraph count').toBe(golden.hfParagraphs)
      expect(hf.hashes).toEqual(golden.hfHashes)
      expect(notes.hashes.length, 'note row count').toBe(golden.noteParagraphs)
      expect(notes.hashes).toEqual(golden.noteHashes)
    })
  }
})
