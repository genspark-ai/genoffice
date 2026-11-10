import { mkdtemp, mkdir, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterAll, beforeAll, describe, expect, it } from 'vitest'
import { inlineImagesForSingleFile, singleFileExportBaseName } from '../src/main/single-file-html'

const PNG_BASE64 =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=='
const PNG_DATA_URL = `data:image/png;base64,${PNG_BASE64}`

let dir: string
let docPath: string

beforeAll(async () => {
  dir = await mkdtemp(join(tmpdir(), 'goff-single-file-'))
  docPath = join(dir, 'page.html')
  await writeFile(docPath, '<!doctype html>', 'utf8')
  await mkdir(join(dir, 'assets'))
  const png = Buffer.from(PNG_BASE64, 'base64')
  await writeFile(join(dir, 'assets', 'logo.png'), png)
  // "save page as, complete" assets carry no extension: the signature types them
  await writeFile(join(dir, 'assets', 'photo'), png)
  await writeFile(join(dir, 'outside.png'), png)
})

afterAll(async () => {
  await rm(dir, { recursive: true, force: true })
})

describe('inlineImagesForSingleFile', () => {
  it('inlines img srcs and CSS url() refs, preserving quote style', async () => {
    const html = [
      '<html><head><style>',
      '.hero { background-image: url(assets/logo.png); }',
      '.card { background: url("assets/logo.png") no-repeat; }',
      '</style></head><body>',
      '<img src="assets/logo.png" alt="logo" width="120">',
      `<div style="background-image: url('assets/logo.png')">x</div>`,
      '</body></html>',
    ].join('\n')
    const r = await inlineImagesForSingleFile(html, docPath)
    expect(r.inlined).toBe(4)
    expect(r.skipped).toEqual([])
    expect(r.html).toContain(`<img src="${PNG_DATA_URL}" alt="logo" width="120">`)
    expect(r.html).toContain(`background-image: url(${PNG_DATA_URL});`)
    expect(r.html).toContain(`background: url("${PNG_DATA_URL}") no-repeat;`)
    expect(r.html).toContain(`background-image: url('${PNG_DATA_URL}')`)
    expect(r.html).not.toContain('assets/logo.png')
  })

  it('types an extensionless asset from its binary signature', async () => {
    const r = await inlineImagesForSingleFile('<img src="assets/photo">', docPath)
    expect(r.inlined).toBe(1)
    expect(r.html).toContain(PNG_DATA_URL)
  })

  it('leaves external, data: and unresolvable references untouched', async () => {
    const html = [
      '<img src="https://cdn.example/pic.png">',
      '<img src="//cdn.example/pic2.png">',
      `<img src="${PNG_DATA_URL}">`,
      '<img src="assets/missing.png">',
      '<img src="../outside.png">',
      '<style>.a { background: url(https://cdn.example/bg.png); }</style>',
    ].join('\n')
    const r = await inlineImagesForSingleFile(html, docPath)
    expect(r.html).toBe(html)
    expect(r.inlined).toBe(0)
    // only the local-looking refs are reported; external URLs are expected to stay
    expect(r.skipped.sort()).toEqual(['../outside.png', 'assets/missing.png'])
  })

  it('returns an unsaved document unchanged (its images are already data URLs)', async () => {
    const html = '<img src="assets/logo.png">'
    const r = await inlineImagesForSingleFile(html, null)
    expect(r.html).toBe(html)
    expect(r.inlined).toBe(0)
  })

  it('inlines every occurrence of a repeated source', async () => {
    const html = '<img src="assets/logo.png"><img src="assets/logo.png">'
    const r = await inlineImagesForSingleFile(html, docPath)
    expect(r.inlined).toBe(2)
    expect(r.html).not.toContain('assets/logo.png')
  })

  it('inlines indented <img> tags: pretty-printed markup is not a Markdown code block', async () => {
    const html = [
      '<html>',
      '  <body>',
      '    <section>',
      '      <img src="assets/logo.png" alt="four spaces">',
      '\t<img src="assets/photo" alt="tab">',
      '    </section>',
      '    <p>![not an image](assets/logo.png)</p>',
      '  </body>',
      '</html>',
    ].join('\n')
    const r = await inlineImagesForSingleFile(html, docPath)
    expect(r.inlined).toBe(2)
    expect(r.skipped).toEqual([])
    expect(r.html).toContain(`      <img src="${PNG_DATA_URL}" alt="four spaces">`)
    expect(r.html).toContain(`\t<img src="${PNG_DATA_URL}" alt="tab">`)
    // Markdown image syntax is literal text in an HTML document
    expect(r.html).toContain('<p>![not an image](assets/logo.png)</p>')
  })
})

describe('singleFileExportBaseName', () => {
  it('suffixes the name so the default save path is never the open document', () => {
    expect(singleFileExportBaseName('landing')).toBe('landing.single')
  })

  it('does not stack the suffix when re-exporting an exported copy', () => {
    expect(singleFileExportBaseName('landing.single')).toBe('landing.single')
    expect(singleFileExportBaseName('landing.SINGLE')).toBe('landing.single')
  })
})
