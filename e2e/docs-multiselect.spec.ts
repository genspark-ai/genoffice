/**
 * Ctrl/⌘+click adds a non-contiguous selection range (genoffice#2013) through
 * the real click path: the modifier is held via the keyboard, the clicks are
 * trusted Chromium input, and the added ranges are read off the live editor.
 *
 * ProseMirror's own selectNodeModifier (⌘/ctrl+click selects the clicked node)
 * must not win here: without the plugin's suppression the first click turns
 * the paragraph into a NodeSelection, and the next click stores that whole
 * paragraph as a "range".
 */
import { test, expect, type Page } from '@playwright/test'
import { mkdtempSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import JSZip from 'jszip'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'

/** the multi-select modifier for this platform, as prosemirror-view spells it */
const MOD = process.platform === 'darwin' ? 'Meta' : 'Control'

interface AidocsWindow {
  __aidocs?: {
    editor?: {
      state: {
        selection: { from: number; to: number; node?: unknown }
        doc: { textBetween: (a: number, b: number, sep: string) => string }
      }
    }
  }
}

async function minimalDocx(paras: string[]): Promise<Buffer> {
  const zip = new JSZip()
  zip.file(
    '[Content_Types].xml',
    '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>',
  )
  zip.file(
    '_rels/.rels',
    '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
  )
  const body = paras
    .map((t) => `<w:p><w:r><w:t xml:space="preserve">${t}</w:t></w:r></w:p>`)
    .join('')
  zip.file(
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}</w:body></w:document>`,
  )
  return zip.generateAsync({ type: 'nodebuffer' })
}

interface WordBox {
  left: number
  right: number
  centerX: number
  centerY: number
}

async function wordBox(page: Page, word: string): Promise<WordBox> {
  return page.evaluate((w) => {
    const root = document.querySelector('.ProseMirror')
    if (!root) throw new Error('no editor')
    const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT)
    let node: Text | null
    while ((node = walker.nextNode() as Text | null)) {
      const i = node.data.indexOf(w)
      if (i < 0) continue
      const r = document.createRange()
      r.setStart(node, i)
      r.setEnd(node, i + w.length)
      const rect = r.getBoundingClientRect()
      if (rect.width === 0) continue
      return {
        left: rect.left,
        right: rect.right,
        centerX: rect.left + rect.width / 2,
        centerY: rect.top + rect.height / 2,
      }
    }
    throw new Error(`word not found: ${w}`)
  }, word)
}

/**
 * A user's ctrl/⌘+click: the modifier goes down first, and the button is held
 * 80 ms so any mousedown-side capture would have fired long before mouseup.
 */
async function modClick(page: Page, at: { centerX: number; centerY: number }): Promise<void> {
  await page.keyboard.down(MOD)
  await page.mouse.click(at.centerX, at.centerY, { delay: 80 })
  await page.keyboard.up(MOD)
}

interface Probe {
  kind: 'text' | 'node'
  text: string
  ranges: string[]
}

async function probe(page: Page): Promise<Probe> {
  return page.evaluate(() => {
    const ed = (window as unknown as AidocsWindow).__aidocs!.editor!
    const s = ed.state.selection
    return {
      kind: s.node ? 'node' : 'text',
      text: ed.state.doc.textBetween(s.from, s.to, '\n'),
      ranges: Array.from(document.querySelectorAll('.doc-multi-range')).map((e) => e.textContent),
    }
  })
}

test.describe('docs non-contiguous selection', () => {
  let dir: string
  let docPath: string

  test.beforeEach(async () => {
    dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-multiselect-')))
    docPath = join(dir, 'words.docx')
    writeFileSync(docPath, await minimalDocx(['echo foxtrot golf hotel india juliet kilo lima']))
  })

  test.afterEach(() => {
    rmSync(dir, { recursive: true, force: true })
  })

  test('ctrl/cmd+click builds a non-contiguous selection, the way Word does', async () => {
    test.setTimeout(180_000)
    const launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'docs-multiselect',
      openFile: docPath,
    })
    const { app } = launched
    try {
      const page = await waitForPageWithUrl(app, '://docs/')
      await page.locator('.doc-page p').first().waitFor()
      // word boxes are only stable once the webfonts have settled
      await page.evaluate(() => document.fonts.ready)

      // a fresh ⌘+click adds the word under the cursor — not the paragraph:
      // PM's selectNodeModifier would make the click a NodeSelection of the
      // whole paragraph, which the next click would then store as a "range"
      await modClick(page, await wordBox(page, 'golf'))
      let p = await probe(page)
      expect(p.kind).toBe('text')
      expect(p.text).toBe('golf')
      expect(p.ranges).toEqual(['golf'])

      // a second ⌘+click adds a second, disjoint range; the live selection
      // still follows the last click
      await modClick(page, await wordBox(page, 'kilo'))
      p = await probe(page)
      expect(p.kind).toBe('text')
      expect(p.text).toBe('kilo')
      expect(p.ranges).toEqual(['golf', 'kilo'])

      // a plain click clears every added range and just places the caret
      const hotel = await wordBox(page, 'hotel')
      await page.mouse.click(hotel.centerX, hotel.centerY)
      p = await probe(page)
      expect(p.ranges).toEqual([])

      // drag a range, then ⌘+click elsewhere: the dragged range is kept
      // verbatim and the clicked word joins it
      const echo = await wordBox(page, 'echo')
      const juliet = await wordBox(page, 'juliet')
      await page.mouse.move(echo.left + 1, echo.centerY)
      await page.mouse.down()
      await page.mouse.move(juliet.right - 1, juliet.centerY, { steps: 8 })
      await page.mouse.up()
      const dragged = await probe(page)
      expect(dragged.text.startsWith('echo foxtrot golf')).toBe(true)
      expect(dragged.ranges).toEqual([])
      await modClick(page, await wordBox(page, 'kilo'))
      p = await probe(page)
      expect(p.text).toBe('kilo')
      expect(p.ranges).toEqual([dragged.text, 'kilo'])

      // a ⌘+double-click adds its word once — the second click runs on
      // ProseMirror's doubleClick path (no handleClick), so it must not stack
      // another copy of the word or the paragraph
      const lima = await wordBox(page, 'lima')
      await page.keyboard.down(MOD)
      await page.mouse.dblclick(lima.centerX, lima.centerY, { delay: 80 })
      await page.keyboard.up(MOD)
      p = await probe(page)
      expect(p.text).toBe('lima')
      expect(p.ranges).toEqual([dragged.text, 'kilo', 'lima'])
    } finally {
      await closeAndSaveVideo(launched, 'docs-multiselect')
    }
  })
})
