/**
 * Golden pin for anchor placement output (anchor-geometry.ts): the attributes
 * and inline styles every anchored object emits (run pictures, picture blocks,
 * floating tables, floating textboxes, nested floats, header/footer floats)
 * over a synthetic matrix of wp:anchor / w:tblpPr combinations plus the truth
 * probes that carry anchors. Only the elements' own open tags are hashed
 * (canonical attribute, class and declaration order), so the pin is about the
 * geometry, not the content. Regenerate with ANCHOR_GOLDEN_UPDATE=1 after an
 * intentional placement change (then prettier the golden directory);
 * ANCHOR_GOLDEN_DUMP=<dir> writes the hashed markup per fixture for diffing.
 */
import { createHash } from 'node:crypto'
import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { Editor } from '@tiptap/core'
import { DOMSerializer } from '@tiptap/pm/model'
import {
  buildAnchoredTextboxParagraphXml,
  parseDocx,
  readSections,
  type HfImage,
  type ParsedDoc,
  type TableModel,
} from '@genoffice/docx-engine'
import { describe, expect, it } from 'vitest'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc, pmDocOptions } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'
import {
  hfFloatPagePos,
  hfRowStyle,
  hfStripGeom,
  makeGapHfEl,
  makeHfFloatImgEl,
  type HfRowGeom,
} from '../src/renderer/editor/hf-dom'
import { nestedFloatCss } from '../src/renderer/editor/protected-render'
import { floatTableMatrixXml } from './helpers/float-table-matrix'

;(globalThis as { CSS?: unknown }).CSS ??= { escape: (s: string) => s }

const HERE = dirname(fileURLToPath(import.meta.url))
const REPO = join(HERE, '../../..')
const GOLDEN_DIR = join(HERE, 'fixtures/anchor-geometry-golden')
const UPDATE = process.env.ANCHOR_GOLDEN_UPDATE === '1'
const DUMP_DIR = process.env.ANCHOR_GOLDEN_DUMP

const ANCHOR_PROBES = [
  'render-anchor-own-line-band',
  'render-bottom-float-section-break',
  'render-hf-anchor-reserve',
  'render-nested-float-table',
  'render-section-start-clears-float',
]

const PIC =
  '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
  '<pic:pic><pic:blipFill><a:blip r:embed="rId10"/></pic:blipFill>' +
  '<pic:spPr><a:xfrm ROT><a:off x="0" y="0"/><a:ext cx="CX" cy="CY"/></a:xfrm>' +
  '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic>'

interface Pic {
  relH?: string
  relV?: string
  x?: number | string
  y?: number | string
  wrap: 'None' | 'Square' | 'Tight' | 'Through' | 'TopAndBottom'
  wrapText?: 'bothSides' | 'left' | 'right'
  behind?: boolean
  cx?: number
  cy?: number
  rot?: number
  dist?: boolean
  layoutInCell?: boolean
}

function anchorXml(p: Pic): string {
  const cx = p.cx ?? 1905000
  const cy = p.cy ?? 952500
  const pos = (v: number | string | undefined) =>
    v === undefined
      ? '<wp:posOffset>0</wp:posOffset>'
      : typeof v === 'number'
        ? `<wp:posOffset>${v}</wp:posOffset>`
        : `<wp:align>${v}</wp:align>`
  const wrap =
    p.wrap === 'None'
      ? '<wp:wrapNone/>'
      : p.wrap === 'TopAndBottom'
        ? '<wp:wrapTopAndBottom/>'
        : `<wp:wrap${p.wrap} wrapText="${p.wrapText ?? 'bothSides'}"/>`
  const dist = p.dist
    ? 'distT="45720" distB="45720" distL="114300" distR="114300"'
    : 'distT="0" distB="0" distL="0" distR="0"'
  return (
    `<w:r><w:drawing><wp:anchor ${dist} simplePos="0" relativeHeight="251659264" behindDoc="${p.behind ? 1 : 0}" locked="0" layoutInCell="${p.layoutInCell === false ? 0 : 1}" allowOverlap="1">` +
    '<wp:simplePos x="0" y="0"/>' +
    `<wp:positionH relativeFrom="${p.relH ?? 'column'}">${pos(p.x)}</wp:positionH>` +
    `<wp:positionV relativeFrom="${p.relV ?? 'paragraph'}">${pos(p.y)}</wp:positionV>` +
    `<wp:extent cx="${cx}" cy="${cy}"/>${wrap}<wp:docPr id="1" name="pic"/>` +
    PIC.replace('CX', String(cx))
      .replace('CY', String(cy))
      .replace(' ROT', p.rot ? ` rot="${p.rot}"` : '') +
    '</wp:anchor></w:drawing></w:r>'
  )
}

const IND = '<w:pPr><w:ind w:left="720" w:firstLine="360"/></w:pPr>'
const text = (s: string) => `<w:r><w:t xml:space="preserve">${s}</w:t></w:r>`

const RELS_H = ['column', 'margin', 'page', 'character', 'leftMargin']
const RELS_V = ['paragraph', 'margin', 'page', 'line', 'topMargin']
const WRAPS: Pic['wrap'][] = ['None', 'Square', 'Tight', 'TopAndBottom']

function pictureMatrix(): Pic[] {
  const out: Pic[] = []
  let k = 0
  for (const wrap of WRAPS) {
    for (const relH of RELS_H) {
      for (const relV of RELS_V) {
        const x = [914400, -228600, 2743200][k % 3]
        const y = [457200, -114300, 0][(k >> 1) % 3]
        out.push({
          relH,
          relV,
          x,
          y,
          wrap,
          wrapText: k % 2 ? 'left' : 'bothSides',
          dist: k % 3 === 0,
        })
        k++
      }
    }
    // align keywords
    for (const relH of ['margin', 'page', 'column']) {
      for (const [ax, ay] of [
        ['center', 'top'],
        ['right', 'bottom'],
        ['left', 'center'],
        ['center', undefined],
        [undefined, 'top'],
      ] as const) {
        out.push({
          relH,
          relV: ay === undefined ? 'paragraph' : relH === 'column' ? 'margin' : relH,
          x: ax,
          y: ay,
          wrap,
        })
      }
    }
  }
  out.push({ wrap: 'None', behind: true, x: 457200, y: 457200 })
  out.push({ wrap: 'Square', behind: true, x: 457200, y: 457200 })
  out.push({ wrap: 'Square', rot: 5400000, x: 914400, y: 914400 })
  out.push({ wrap: 'TopAndBottom', rot: 5400000, x: 914400, y: 914400 })
  out.push({ wrap: 'None', rot: 5400000, x: 914400, y: 914400 })
  out.push({ wrap: 'Square', relV: 'page', y: 228600, x: 0 })
  out.push({ wrap: 'Square', x: 0, y: 2743200, cx: 5486400 })
  out.push({ wrap: 'TopAndBottom', x: 0, y: 914400 })
  out.push({ wrap: 'Square', x: 914400, y: -914400, layoutInCell: false })
  return out
}

function runPictureDoc(): string {
  return pictureMatrix()
    .map(
      (p, i) => `<w:p>${i % 2 ? IND : ''}${text('before ')}${anchorXml(p)}${text(' after')}</w:p>`,
    )
    .join('')
}

function pictureBlockDoc(): string {
  return pictureMatrix()
    .map((p, i) => `<w:p>${i % 2 ? IND : ''}${anchorXml(p)}</w:p>${i % 5 === 0 ? '<w:p/>' : ''}`)
    .join('')
}

function cellPictureDoc(): string {
  const cell = (inner: string) =>
    `<w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>${inner}</w:tc>`
  const rows = pictureMatrix()
    .filter((_p, i) => i % 4 === 0)
    .map(
      (p) =>
        `<w:tr>${cell(`<w:p>${text('a ')}${anchorXml(p)}</w:p>`)}${cell(`<w:p>${anchorXml(p)}</w:p>`)}</w:tr>`,
    )
    .join('')
  return `<w:tbl><w:tblPr><w:tblW w:w="8000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>${rows}</w:tbl><w:p/>`
}

function textboxDoc(): string {
  const paragraphs = [{ runs: [{ text: 'box' }], format: {} }]
  const out: string[] = []
  let id = 10
  for (const anchor of ['paragraph', 'page'] as const) {
    for (const wrap of ['topAndBottom', 'none'] as const) {
      for (const [x, y] of [
        [91440, 12700],
        [-457200, 914400],
        [4572000, -228600],
      ]) {
        out.push(
          buildAnchoredTextboxParagraphXml({
            anchor,
            wrap,
            xEmu: x,
            yEmu: y,
            widthEmu: 2743200,
            heightEmu: 914400,
            fillHex: 'DDDDDD',
            insetsEmu: { l: 91440, t: 45720, r: 91440, b: 45720 },
            zOrder: 3,
            id: id++,
            paragraphs,
            holderLineTwips: 20,
          }),
        )
      }
    }
  }
  return out.join('') + '<w:p/>'
}

const SYNTHETIC: Record<string, () => string> = {
  'synthetic-run-pictures': runPictureDoc,
  'synthetic-picture-blocks': pictureBlockDoc,
  'synthetic-cell-pictures': cellPictureDoc,
  'synthetic-float-tables': floatTableMatrixXml,
  'synthetic-textboxes': textboxDoc,
}

const ANCHOR_SEL = [
  '[data-inline-image]',
  '.doc-inline-img-anchor',
  '.doc-protected-image',
  '.doc-img-wrap',
  '.doc-img-float',
  '[class*="img-wrap-"]',
  'table',
  '.doc-protected-floating',
  '.doc-textbox',
  '.box-rotator',
  '[data-tblp-vy]',
  '[data-page-rel-v]',
  '[data-page-rel-x]',
  '[data-wrap-side]',
  '[data-band]',
  '.page-hf-float-img',
  '.page-hf-row',
].join(',')

const sha = (s: string) => createHash('sha1').update(s).digest('hex').slice(0, 16)

/** open tag with attributes, class tokens and style declarations in sorted order */
function openTag(el: Element): string {
  const attrs = Array.from(el.attributes)
    .filter(({ name }) => name !== 'src')
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
  return `<${el.tagName.toLowerCase()}${attrs.length ? ` ${attrs.join(' ')}` : ''}>`
}

function editorTags(parsed: ParsedDoc): string[] {
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks, readSections(parsed), pmDocOptions(parsed)) as never,
  })
  try {
    const serializer = DOMSerializer.fromSchema(editor.schema)
    const frag = serializer.serializeFragment(editor.state.doc.content, { document })
    const host = document.createElement('div')
    host.appendChild(frag)
    return Array.from(host.querySelectorAll(ANCHOR_SEL)).map(openTag)
  } finally {
    editor.destroy()
  }
}

/** nested float CSS of every table model, top-level and nested */
function nestedTags(parsed: ParsedDoc): string[] {
  const out: string[] = []
  const visit = (t: TableModel) => {
    const css = nestedFloatCss(t)
    if (css) out.push(css.join(';'))
    for (const row of t.rows)
      for (const cell of row) for (const n of cell.nestedTables ?? []) visit(n)
  }
  for (const b of parsed.blocks) if (b.table) visit(b.table)
  return out
}

function hfTags(parsed: ParsedDoc): string[] {
  const section = readSections(parsed)[0]
  const geom = section ? hfStripGeom(section.settings) : undefined
  return Object.entries(parsed.hfParts ?? {})
    .sort(([a], [b]) => a.localeCompare(b))
    .flatMap(([name, part]) => {
      const el = makeGapHfEl({
        kind: /footer|ftr/i.test(name) ? 'footer' : 'header',
        value: { text: part.text, paras: part.paras },
        images: part.images,
        pageNo: 1,
        pageTotal: 1,
        geom,
      })
      const floats = (part.images ?? [])
        .filter((img) => img.floating)
        .flatMap((img) => [
          openTag(makeHfFloatImgEl(img, HF_BOX, 'gap')),
          openTag(makeHfFloatImgEl(img, HF_BOX, 'lead')),
        ])
      return [...Array.from(el.querySelectorAll(ANCHOR_SEL)).map(openTag), ...floats]
    })
}

/** Letter, 1in margins, strips 0.5in from the edges (px) */
const HF_BOX = {
  pageW: 816,
  pageH: 1056,
  marginLeft: 96,
  marginRight: 96,
  marginTop: 120,
  marginBottom: 96,
  headerDist: 48,
  sectMarginTop: 96,
  sectMarginBottom: 96,
  paraOriginY: 60,
}

function hfNumericLines(): string[] {
  const out: string[] = []
  const hRels: Array<HfImage['posHRel']> = [
    undefined,
    'page',
    'margin',
    'leftMargin',
    'rightMargin',
  ]
  const vRels: Array<HfImage['posVRel']> = [
    undefined,
    'page',
    'margin',
    'paragraph',
    'topMargin',
    'bottomMargin',
  ]
  const wraps: Array<HfImage['wrap']> = [undefined, 'none', 'square']
  for (const posHRel of hRels)
    for (const posVRel of vRels)
      for (const wrap of wraps) {
        for (const [posXPx, posYPx, posH, posV] of [
          [10, 20, undefined, undefined],
          [undefined, undefined, 'center', 'center'],
          [undefined, undefined, 'right', 'bottom'],
          [undefined, 30, 'left', undefined],
          [-40, undefined, undefined, 'top'],
        ] as const) {
          const img: HfImage = { dataUrl: '', posHRel, posVRel, wrap, posXPx, posYPx, posH, posV }
          const p = hfFloatPagePos(img, HF_BOX)
          out.push(
            `${posHRel}/${posVRel}/${wrap}/${posXPx}/${posYPx}/${posH}/${posV} -> ${JSON.stringify(p)}`,
          )
        }
      }
  const geom: HfRowGeom = {
    kind: 'header',
    columnPx: 624,
    stripLeftPx: 96,
    pageWPx: 816,
    marginRightPx: 96,
    marginTopPx: 96,
    headerStripTopPx: 48,
  }
  for (const horzAnchor of ['page', 'margin', 'text'] as const)
    for (const vertAnchor of ['page', 'margin', 'text'] as const)
      for (const xSpec of [undefined, 'left', 'center', 'right'] as const)
        for (const [xTwips, yTwips] of [
          [600, 300],
          [-200, undefined],
        ] as const) {
          const row = {
            cells: [],
            widthTwips: 3000,
            float: { horzAnchor, vertAnchor, xSpec, xTwips, yTwips },
          }
          const style = hfRowStyle(row as never, geom, true)
          out.push(
            `${horzAnchor}/${vertAnchor}/${xSpec}/${xTwips}/${yTwips} -> ${JSON.stringify(style)}`,
          )
          out.push(
            `footer -> ${JSON.stringify(hfRowStyle(row as never, { ...geom, kind: 'footer' }, true))}`,
          )
        }
  return out
}

interface Golden {
  source: string
  editor: string[]
  nested: string[]
  hf: string[]
  numeric?: string
}

function dump(name: string, kind: string, lines: string[]): void {
  if (!DUMP_DIR) return
  mkdirSync(DUMP_DIR, { recursive: true })
  writeFileSync(join(DUMP_DIR, `${name}.${kind}.txt`), lines.map((x) => `${x}\n`).join(''))
}

function pin(name: string, source: string, got: Omit<Golden, 'source'>): void {
  const file = join(GOLDEN_DIR, `${name}.json`)
  if (UPDATE) {
    mkdirSync(GOLDEN_DIR, { recursive: true })
    writeFileSync(file, `${JSON.stringify({ source, ...got }, null, 2)}\n`)
  }
  const golden = JSON.parse(readFileSync(file, 'utf8')) as Golden
  expect(got.editor.length, 'editor anchor elements').toBe(golden.editor.length)
  expect(got.editor).toEqual(golden.editor)
  expect(got.nested).toEqual(golden.nested)
  expect(got.hf).toEqual(golden.hf)
  if (got.numeric !== undefined) expect(got.numeric).toBe(golden.numeric)
}

async function pinParsed(name: string, source: string, parsed: ParsedDoc, numeric?: string[]) {
  const editor = editorTags(parsed)
  const nested = nestedTags(parsed)
  const hf = hfTags(parsed)
  dump(name, 'editor', editor)
  dump(name, 'nested', nested)
  dump(name, 'hf', hf)
  if (numeric) dump(name, 'numeric', numeric)
  expect(editor.length + hf.length, 'the fixture must carry anchored objects').toBeGreaterThan(0)
  pin(name, source, {
    editor: editor.map(sha),
    nested: nested.map(sha),
    hf: hf.map(sha),
    ...(numeric ? { numeric: sha(numeric.join('\n')) } : {}),
  })
}

describe('anchor-geometry golden (anchor attrs and styles of every anchored object)', () => {
  for (const [name, body] of Object.entries(SYNTHETIC)) {
    it(name, async () => {
      const parsed = await parseDocx(await buildDocx({ bodyXml: body(), withImage: true }))
      await pinParsed(
        name,
        `synthetic:${name}`,
        parsed,
        name === 'synthetic-textboxes' ? hfNumericLines() : undefined,
      )
    })
  }
  for (const dir of ANCHOR_PROBES) {
    const source = `packages/docx-engine/tests/fixtures/truths/${dir}/probe.docx`
    it(`truth-${dir}`, async () => {
      if (!existsSync(join(REPO, source))) return
      const parsed = await parseDocx(new Uint8Array(readFileSync(join(REPO, source))))
      await pinParsed(`truth-${dir}`, source, parsed)
    })
  }
})
