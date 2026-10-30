import { existsSync, readFileSync, statSync, writeFileSync } from 'node:fs'
import { createRequire } from 'node:module'
import { basename, join, resolve } from 'node:path'
import { mkdtempSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { runInNewContext } from 'node:vm'
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { parseDocx } from '@genoffice/docx-engine'
import { openPptx } from '@genoffice/pptx-engine'
import { LANGS } from '@genoffice/i18n'
import {
  publicEntry,
  readSourceManifest,
  sourceFilesFor,
  TEMPLATE_CATEGORIES,
} from '../../tools/gen-templates.mjs'
import { TEMPLATE_GALLERY_STRINGS } from '../src/renderer/src/template-gallery-strings'

const TEMPLATES_DIR = resolve(import.meta.dirname, '../build/templates')

const require = createRequire(import.meta.url)
// Read packaging metadata without downloading Electron (the shell-new pattern).
const configModule = { exports: {} as ReturnType<typeof require> }
runInNewContext(readFileSync(resolve(import.meta.dirname, '../electron-builder.cjs'), 'utf8'), {
  module: configModule,
  __dirname: resolve(import.meta.dirname, '..'),
  process: { platform: 'linux', arch: 'x64', env: {} },
  require: (id: string) =>
    id === 'node:fs' ? { ...require(id), existsSync: () => true } : require(id),
})
const config = configModule.exports

function templateFile(name: string): Buffer {
  const path = join(TEMPLATES_DIR, name)
  expect(existsSync(path), `${name} must ship in build/templates`).toBe(true)
  return readFileSync(path)
}

/** the shipped manifest: what the IPC reads and the gallery shows */
function shippedManifest(): Array<ReturnType<typeof publicEntry>> {
  return JSON.parse(readFileSync(join(TEMPLATES_DIR, 'templates.json'), 'utf8'))
}

describe('template center — manifest matches the outputs', () => {
  it('ships at least eight templates across docx, xlsx and md', () => {
    const manifest = shippedManifest()
    expect(manifest.length).toBeGreaterThanOrEqual(8)
    const kinds = new Set(manifest.map((entry) => entry.kind))
    expect(kinds.has('docx')).toBe(true)
    expect(kinds.has('xlsx')).toBe(true)
    expect(kinds.has('md')).toBe(true)
  })

  it('lists exactly the authored source manifest, as public entries', () => {
    expect(shippedManifest()).toEqual(readSourceManifest().map(publicEntry))
  })

  it('every entry names an existing, non-empty file with a matching extension', () => {
    for (const entry of shippedManifest()) {
      expect(Object.keys(entry).sort()).toEqual([
        'category',
        'description',
        'file',
        'id',
        'kind',
        'title',
      ])
      const path = join(TEMPLATES_DIR, entry.file)
      expect(existsSync(path), `${entry.file} must exist`).toBe(true)
      expect(statSync(path).size).toBeGreaterThan(0)
      expect(entry.file.endsWith(`.${entry.kind}`)).toBe(true)
      expect(TEMPLATE_CATEGORIES).toContain(entry.category)
    }
  })

  it('carries unique ids', () => {
    const ids = shippedManifest().map((entry) => entry.id)
    expect(new Set(ids).size).toBe(ids.length)
  })

  it('every template has its authored sources under src/', () => {
    for (const entry of readSourceManifest()) {
      for (const source of sourceFilesFor(entry)) {
        expect(existsSync(source), `${basename(source)} must exist`).toBe(true)
      }
    }
  })

  it('rejects a malformed source manifest entry', () => {
    const dir = mkdtempSync(join(tmpdir(), 'tpl-src-'))
    const write = (entry: Record<string, string>) =>
      writeFileSync(join(dir, 'manifest.json'), JSON.stringify([entry]))
    write({ id: 'x', title: 'x', description: 'x', kind: 'exe', category: 'office', file: 'x.exe' })
    expect(() => readSourceManifest(dir)).toThrow(/kind must be/)
    write({
      id: 'x',
      title: 'x',
      description: 'x',
      kind: 'docx',
      category: 'office',
      file: '../evil.docx',
    })
    expect(() => readSourceManifest(dir)).toThrow(/plain file name/)
    write({
      id: 'X',
      title: 'x',
      description: 'x',
      kind: 'docx',
      category: 'office',
      file: 'X.docx',
    })
    expect(() => readSourceManifest(dir)).toThrow(/id must be/)
    write({ id: 'x', title: 'x', description: 'x', kind: 'docx', category: 'nope', file: 'x.docx' })
    expect(() => readSourceManifest(dir)).toThrow(/category must be/)
  })

  it('packages the Home desktop resources for every platform', () => {
    // desktop-ipc.ts and templates-ipc.ts read these from
    // process.resourcesPath in every packaged build, so the entries belong to
    // the top-level extraResources: as win-only entries the mac/linux installs
    // shipped with an empty wallpaper picker and template center.
    for (const name of ['templates', 'wallpapers']) {
      const resource = config.extraResources.find((entry: { to: string }) => entry.to === name)
      expect(resource, `every packaged platform must ship resources/${name}`).toBeDefined()
      expect(resource.from).toBe(`build/${name}`)
      expect(
        config.win.extraResources.find((entry: { to: string }) => entry.to === name),
        `win.extraResources must not duplicate resources/${name}`,
      ).toBeUndefined()
    }
  })
})

describe('template center — the shipped office files are real documents', () => {
  const docxEntries = () => shippedManifest().filter((entry) => entry.kind === 'docx')

  it('opens every Word template with headings, tables and zh placeholders', async () => {
    expect(docxEntries().length).toBeGreaterThanOrEqual(4)
    for (const entry of docxEntries()) {
      const doc = await parseDocx(templateFile(entry.file))
      expect(
        doc.blocks.length,
        `${entry.file}: a real template is not one paragraph`,
      ).toBeGreaterThan(8)
      expect(
        doc.blocks.some((block) => block.type === 'heading'),
        `${entry.file}: needs headings`,
      ).toBe(true)
      expect(
        doc.blocks.some((block) => block.type === 'table'),
        `${entry.file}: needs a table`,
      ).toBe(true)
      const zip = await JSZip.loadAsync(templateFile(entry.file), { checkCRC32: true })
      const document = await zip.file('word/document.xml')!.async('string')
      expect(document).toContain('〔填写')
      expect(document).not.toMatch(/lorem/i)
    }
  })

  it('ships Excel workbooks with real formulas', async () => {
    const entries = shippedManifest().filter((entry) => entry.kind === 'xlsx')
    expect(entries.length).toBeGreaterThanOrEqual(2)
    for (const entry of entries) {
      const zip = await JSZip.loadAsync(templateFile(entry.file), { checkCRC32: true })
      const workbook = await zip.file('xl/workbook.xml')!.async('string')
      expect(workbook.match(/<sheet\s/g)?.length ?? 0).toBeGreaterThan(0)
      const sheets = Object.keys(zip.files).filter((name) =>
        /^xl\/worksheets\/sheet\d+\.xml$/.test(name),
      )
      const formulas = await Promise.all(
        sheets.map(async (name) => (await zip.file(name)!.async('string')).match(/<f>/g) ?? []),
      )
      expect(
        formulas.reduce((count, found) => count + found.length, 0),
        `${entry.file}: needs live formulas`,
      ).toBeGreaterThan(5)
    }
  })

  it('opens the PowerPoint template as a designed multi-slide deck', async () => {
    const entries = shippedManifest().filter((entry) => entry.kind === 'pptx')
    expect(entries.length).toBeGreaterThanOrEqual(1)
    for (const entry of entries) {
      const opened = await openPptx(templateFile(entry.file))
      expect(
        opened.deck.slides.length,
        `${entry.file}: a deck, not one blank slide`,
      ).toBeGreaterThanOrEqual(5)
      for (const slide of opened.deck.slides) {
        expect(slide.elements.length, `${entry.file}: every slide carries content`).toBeGreaterThan(
          0,
        )
      }
      const text = opened.deck.slides
        .flatMap((slide) => slide.elements)
        .flatMap((element) => element.text?.paragraphs ?? [])
        .flatMap((paragraph) => paragraph.runs.map((run) => run.text ?? ''))
        .join('')
      expect(text).toContain('项目启动会')
      expect(text).toContain('〔填写')
    }
  })

  it('ships markdown templates as authored text with placeholders', () => {
    const entries = shippedManifest().filter((entry) => entry.kind === 'md')
    expect(entries.length).toBeGreaterThanOrEqual(2)
    for (const entry of entries) {
      const text = templateFile(entry.file).toString('utf8')
      expect(text.trim()).not.toBe('')
      expect(text).toContain('〔填写')
      expect(text).not.toMatch(/lorem/i)
    }
  })
})

describe('template gallery strings — every supported language', () => {
  it('covers exactly the LANGS set', () => {
    expect(Object.keys(TEMPLATE_GALLERY_STRINGS).sort()).toEqual([...LANGS].sort())
  })

  it('carries no empty label', () => {
    for (const [lang, strings] of Object.entries(TEMPLATE_GALLERY_STRINGS)) {
      for (const [key, value] of Object.entries(strings)) {
        if (typeof value === 'string') {
          expect(value.length, `${lang}.${key} is empty`).toBeGreaterThan(0)
        } else {
          for (const [category, label] of Object.entries(value)) {
            expect(label.length, `${lang}.categories.${category} is empty`).toBeGreaterThan(0)
          }
        }
      }
    }
  })
})
