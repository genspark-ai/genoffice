/**
 * Insert > Icons: the bundled icon data, the search over it, and the way an
 * inserted icon is saved (an SVG part with the bitmap shown here as fallback).
 */
import { execFileSync } from 'node:child_process'
import { existsSync, readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { gzipSync } from 'node:zlib'
import { afterEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import JSZip from 'jszip'
import {
  buildBlankDocx,
  inlinePictureDrawingXml,
  parseDocx,
  saveDocx,
} from '@genoffice/docx-engine'
import { buildDocx, TINY_PNG_BASE64 } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { insertTableAt } from '../src/renderer/components/ribbon-tabs'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'
import {
  insertPictureAtSelection,
  insertPicturesInLine,
  settleCaretAfterPicture,
} from '../src/renderer/editor/insert-picture'
import { rememberPictureVector } from '../src/renderer/editor/picture-vectors'
import {
  ICON_CATEGORIES,
  ICON_INSERT_PX,
  iconSvg,
  loadIconLibrary,
  searchIcons,
} from '../src/renderer/icon-library'

const DATA = resolve(__dirname, '../src/renderer/icon-library/icons.json')
const BUILD = resolve(__dirname, '../../../tools/docs-icon-library/build.mjs')
const PACKAGE = resolve(__dirname, '../../../node_modules/@fluentui/svg-icons/package.json')

const names = (found: Array<{ name: string }>) => found.map((icon) => icon.name)

describe('the icon data', () => {
  it('holds 400 to 500 icons, each once, each in a known category', async () => {
    const icons = await loadIconLibrary()
    expect(icons.length).toBeGreaterThanOrEqual(400)
    expect(icons.length).toBeLessThanOrEqual(500)
    expect(new Set(names(icons)).size).toBe(icons.length)
    for (const icon of icons) {
      expect(ICON_CATEGORIES, icon.name).toContain(icon.category)
      expect(icon.label.trim(), icon.name).not.toBe('')
      expect(icon.path, icon.name).toMatch(/^[Mm][MmLlHhVvCcSsQqTtAaZz0-9eE.,\s-]+$/)
    }
  })

  it('names the categories the dialog lists, none of them empty', async () => {
    const icons = await loadIconLibrary()
    const data = JSON.parse(readFileSync(DATA, 'utf8')) as { categories: string[] }
    expect(data.categories).toEqual([...ICON_CATEGORIES])
    for (const category of ICON_CATEGORIES) {
      expect(icons.filter((icon) => icon.category === category).length, category).toBeGreaterThan(5)
    }
  })

  // a worktree sharing another checkout's node_modules may not have the package
  it.skipIf(!existsSync(PACKAGE))(
    'is what the generator makes from the list and the pinned package',
    () => {
      execFileSync(process.execPath, [BUILD, '--check'], { stdio: 'pipe' })
    },
  )

  it('stays well under a megabyte', () => {
    const raw = readFileSync(DATA)
    expect(raw.length).toBeLessThan(300_000)
    expect(gzipSync(raw).length).toBeLessThan(100_000)
  })
})

describe('searching the icons', () => {
  it('finds an icon by a word of its name or of its keywords, whatever the case', async () => {
    const icons = await loadIconLibrary()
    expect(names(searchIcons(icons, 'heart'))).toEqual(expect.arrayContaining(['heart']))
    expect(names(searchIcons(icons, 'CARDIOLOGY'))).toEqual(['heart_pulse'])
    expect(names(searchIcons(icons, 'lorry'))).toEqual(
      expect.arrayContaining(['vehicle_truck', 'vehicle_truck_profile']),
    )
    expect(names(searchIcons(icons, 'wi'))).toContain('wifi_1')
  })

  it('wants every word of the query, each at the start of a word', async () => {
    const icons = await loadIconLibrary()
    expect(names(searchIcons(icons, 'pie chart'))).toEqual(['data_pie'])
    expect(names(searchIcons(icons, 'eart'))).not.toContain('heart')
    expect(searchIcons(icons, 'no such thing')).toEqual([])
  })

  it('returns everything for an empty query', async () => {
    const icons = await loadIconLibrary()
    expect(searchIcons(icons, '  ')).toBe(icons)
  })
})

describe('an icon as an SVG file', () => {
  it('is one black path in the icon square, an inch wide', async () => {
    const icon = (await loadIconLibrary()).find((candidate) => candidate.name === 'heart')!
    const svg = iconSvg(icon)
    expect(svg).toBe(
      `<svg xmlns="http://www.w3.org/2000/svg" width="96" height="96" viewBox="0 0 24 24"><path d="${icon.path}"/></svg>`,
    )
    expect(ICON_INSERT_PX).toBe(96)
    const parsed = new DOMParser().parseFromString(svg, 'image/svg+xml')
    expect(parsed.querySelector('parsererror')).toBeNull()
    expect(parsed.documentElement.querySelectorAll('path')).toHaveLength(1)
  })
})

const PNG_A =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=='
const PNG_B =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=='
const SVG =
  '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"><path d="M2 2h20v20H2z"/></svg>'

const picture = (base64: string) => ({
  dataUrl: `data:image/png;base64,${base64}`,
  base64,
  mime: 'image/png',
  widthPx: ICON_INSERT_PX,
  heightPx: ICON_INSERT_PX,
  label: 'Square',
  altText: 'Square',
})

const liveEditors: Editor[] = []
afterEach(() => {
  for (const editor of liveEditors.splice(0)) editor.destroy()
})

async function openBlank() {
  const parsed = await parseDocx(await buildBlankDocx())
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
  liveEditors.push(editor)
  return { editor, parsed }
}

async function save(editor: Editor, parsed: Awaited<ReturnType<typeof parseDocx>>) {
  const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
  const bytes = await saveDocx(parsed, plan.saveBlocks, { newPictures: plan.newPictures })
  const zip = await JSZip.loadAsync(bytes)
  const media = Object.keys(zip.files).filter((f) => /^word\/media\/[^/]+$/.test(f))
  return { bytes, zip, media, xml: await zip.file('word/document.xml')!.async('string') }
}

describe('saving a picture inserted from a vector', () => {
  it('writes the SVG beside the bitmap, with the alt text, and leaves other pictures plain', async () => {
    rememberPictureVector(PNG_A, btoa(SVG))
    const { editor, parsed } = await openBlank()
    for (const base64 of [PNG_A, PNG_B]) {
      insertPictureAtSelection(editor, picture(base64))
      settleCaretAfterPicture(editor)
    }
    const { zip, media, xml } = await save(editor, parsed)
    expect(media.filter((f) => f.endsWith('.svg'))).toHaveLength(1)
    expect(media.filter((f) => f.endsWith('.png'))).toHaveLength(2)
    expect(await zip.file(media.find((f) => f.endsWith('.svg'))!)!.async('string')).toBe(SVG)
    expect(xml.match(/<w:drawing>/g)).toHaveLength(2)
    expect(xml.match(/<asvg:svgBlip /g)).toHaveLength(1)
    expect(xml.match(/<wp:extent cx="914400" cy="914400"\/>/g)).toHaveLength(2)
    expect(xml).toContain('descr="Square"')
  })

  const firstCell = (editor: Editor) => {
    let cell = -1
    editor.state.doc.descendants((node, pos) => {
      if (cell < 0 && node.type.name === 'docTableCell') cell = pos + 2
      return cell < 0
    })
    return cell
  }
  const tableOf = (xml: string) => xml.slice(xml.indexOf('<w:tbl>'), xml.indexOf('</w:tbl>'))

  it('does the same for a picture in a table cell, with the name it was remembered under', async () => {
    rememberPictureVector(PNG_A, btoa(SVG), 'Square & circle')
    const { editor, parsed } = await openBlank()
    insertTableAt(editor, 1, 1)
    editor.commands.setTextSelection(firstCell(editor))
    insertPictureAtSelection(editor, picture(PNG_A))
    const { media, xml } = await save(editor, parsed)
    const table = tableOf(xml)
    expect(table).toMatch(/<asvg:svgBlip [^>]*r:embed="rId\d+"\/>/)
    expect(table).toMatch(/<wp:docPr id="\d+" name="Picture \d+" descr="Square &amp; circle"\/>/)
    expect(xml).not.toContain('aidocs-new-pic-')
    expect(media.filter((f) => f.endsWith('.svg'))).toHaveLength(1)
  })

  it('writes no description for a cell picture that came with no name', async () => {
    rememberPictureVector(PNG_A, btoa(SVG))
    const { editor, parsed } = await openBlank()
    insertTableAt(editor, 1, 1)
    editor.commands.setTextSelection(firstCell(editor))
    insertPictureAtSelection(editor, picture(PNG_A))
    insertPictureAtSelection(editor, picture(PNG_B))
    const table = tableOf((await save(editor, parsed)).xml)
    expect(table.match(/<wp:docPr /g)).toHaveLength(2)
    expect(table).not.toContain('descr=')
  })

  it('leaves a cell picture that came from the file as it was, whatever is remembered for its bytes', async () => {
    const drawing = inlinePictureDrawingXml({ rId: 'rId10', docPrId: 7, widthPx: 40, heightPx: 40 })
    const cellXml = `<w:p><w:r>${drawing}</w:r></w:p>`
    const source = await buildDocx({
      withImage: true,
      bodyXml:
        '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>' +
        `<w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>${cellXml}</w:tc></w:tr></w:tbl>` +
        '<w:p><w:r><w:t>below</w:t></w:r></w:p>',
    })
    rememberPictureVector(TINY_PNG_BASE64, btoa(SVG), 'Remembered')
    const parsed = await parseDocx(source)
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks) as never,
    })
    liveEditors.push(editor)
    let after = -1
    editor.state.doc.descendants((node, pos) => {
      if (node.type.name === 'docInlineImage') after = pos + node.nodeSize
      return true
    })
    expect(after).toBeGreaterThan(0)
    // an edit in the cell makes the save write the cell again
    editor.commands.setTextSelection(after)
    editor.commands.insertContent('x')
    const { media, xml } = await save(editor, parsed)
    const table = tableOf(xml)
    expect(table).toMatch(/>x<\/w:t>/)
    expect(table).toContain(drawing)
    expect(table).not.toContain('descr=')
    expect(table).not.toContain('svgBlip')
    expect(media).toEqual(['word/media/image1.png'])
  })

  it('keeps both pictures and the vector through a reopen and a second save', async () => {
    rememberPictureVector(PNG_A, btoa(SVG))
    const { editor, parsed } = await openBlank()
    insertPictureAtSelection(editor, picture(PNG_A))
    settleCaretAfterPicture(editor)
    const first = await save(editor, parsed)
    const reopened = await parseDocx(first.bytes)
    const pictures = reopened.blocks.filter((block) => block.type === 'image')
    expect(pictures).toHaveLength(1)
    expect(pictures[0].imageDataUrl).toBe(`data:image/png;base64,${PNG_A}`)
    const again = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(reopened.blocks) as never,
    })
    liveEditors.push(again)
    const second = await save(again, reopened)
    expect(second.xml.match(/<asvg:svgBlip /g)).toHaveLength(1)
    expect(second.media.sort()).toEqual(first.media.sort())
  })
})

describe('inserting several icons at once', () => {
  const VECTOR = `data:image/svg+xml;base64,${btoa(SVG)}`
  /** each top-level block as its text, a picture run standing as [alt] */
  const lines = (editor: Editor) => {
    const out: string[] = []
    editor.state.doc.forEach((block) => {
      let line = ''
      block.descendants((node) => {
        if (node.isText) line += node.text
        if (node.type.name === 'docInlineImage') line += '[icon]'
        return true
      })
      out.push(block.type.name === 'docProtected' ? '[block picture]' : line)
    })
    return out
  }
  const icons = (count: number) =>
    Array.from({ length: count }, () => ({ ...picture(PNG_A), vectorUrl: VECTOR }))

  it('puts them side by side in the caret paragraph in one transaction, so one Undo takes all of them out', async () => {
    const { editor } = await openBlank()
    editor.commands.insertContent('before')
    let changes = 0
    editor.on('transaction', ({ transaction }) => {
      if (transaction.docChanged) changes++
    })
    expect(insertPicturesInLine(editor, icons(3))).toBe(true)
    expect(changes).toBe(1)
    expect(lines(editor)).toEqual(['before[icon][icon][icon]'])

    editor.commands.undo()
    expect(lines(editor).join('')).not.toContain('[icon]')
    editor.commands.redo()
    expect(lines(editor)).toEqual(['before[icon][icon][icon]'])
    // the caret is after the last icon
    editor.commands.insertContent('after')
    expect(lines(editor)).toEqual(['before[icon][icon][icon]after'])
  })

  it('saves them as picture runs of that one paragraph, each with its SVG', async () => {
    rememberPictureVector(PNG_A, btoa(SVG), 'Square')
    const { editor, parsed } = await openBlank()
    editor.commands.insertContent('before')
    insertPicturesInLine(editor, icons(2))
    const { xml, media } = await save(editor, parsed)
    const paragraph = /<w:p\b(?:(?!<\/w:p>)[\s\S])*before[\s\S]*?<\/w:p>/.exec(xml)![0]
    expect(paragraph.match(/<wp:inline /g)).toHaveLength(2)
    expect(paragraph.match(/<asvg:svgBlip /g)).toHaveLength(2)
    expect(paragraph.match(/descr="Square"/g)).toHaveLength(2)
    expect(xml).not.toContain('aidocs-new-pic-')
    // the same bytes twice are one bitmap part and one SVG part
    expect(media.filter((f) => f.endsWith('.svg'))).toHaveLength(1)
    const reopened = await parseDocx((await save(editor, parsed)).bytes)
    const runs = reopened.blocks[0].runs?.filter((run) => run.image) ?? []
    expect(runs.map((run) => run.image?.vectorUrl)).toEqual([VECTOR, VECTOR])
  })

  it('keeps the name of each icon on its picture run, where Alt Text reads it and the save writes it', async () => {
    const { editor, parsed } = await openBlank()
    insertPicturesInLine(editor, [
      { ...picture(PNG_A), altText: 'Heart' },
      { ...picture(PNG_B), altText: 'Rocket & star' },
    ])
    const names: unknown[] = []
    editor.state.doc.descendants((node) => {
      if (node.type.name === 'docInlineImage') names.push(node.attrs.altText)
      return true
    })
    expect(names).toEqual(['Heart', 'Rocket & star'])
    const { xml } = await save(editor, parsed)
    expect(xml).toContain('descr="Heart"')
    expect(xml).toContain('descr="Rocket &amp; star"')
  })

  it('draws an icon from its SVG until its bitmap is another one', async () => {
    const { editor } = await openBlank()
    insertPicturesInLine(editor, icons(1))
    const img = () => editor.view.dom.querySelector('img.doc-inline-img') as HTMLImageElement
    expect(img().getAttribute('src')).toBe(VECTOR)
    let at = -1
    editor.state.doc.descendants((node, pos) => {
      if (node.type.name === 'docInlineImage') at = pos
      return true
    })
    const other = `data:image/png;base64,${PNG_B}`
    editor.view.dispatch(editor.state.tr.setNodeAttribute(at, 'dataUrl', other))
    expect(img().getAttribute('src')).toBe(other)
  })

  it('gives each a paragraph of its own where there is no paragraph to take them', async () => {
    const { editor } = await openBlank()
    insertPictureAtSelection(editor, picture(PNG_B))
    // the block picture is selected: no text position to put a run at
    expect(editor.state.selection.$from.parent.isTextblock).toBe(false)
    expect(insertPicturesInLine(editor, icons(2))).toBe(true)
    expect(
      lines(editor).filter((line) => line === '[block picture]').length,
    ).toBeGreaterThanOrEqual(2)
  })

  it('inserts nothing for an empty pick', async () => {
    const { editor } = await openBlank()
    expect(insertPicturesInLine(editor, [])).toBe(false)
    expect(lines(editor).join('')).toBe('')
  })
})
