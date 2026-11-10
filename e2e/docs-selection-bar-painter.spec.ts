import { test, expect, type Page } from '@playwright/test'
import { mkdtempSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import JSZip from 'jszip'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'

/**
 * Two Word mouse behaviors in Docs:
 * - the left page margin is the selection bar: click = line, drag = lines,
 *   double-click = paragraph, triple-click = document, Shift+click extends;
 * - the Format Painter brushes the clicked WORD, double-clicking its button
 *   locks it on until Esc, and copy/paste-formatting chords exist.
 */

interface AidocsWindow {
  __aidocs?: { editor?: unknown }
}

const LONG = Array.from({ length: 70 }, (_, i) => `lorem${i}`).join(' ')
const PARAS = ['alpha one two', 'beta four five', 'gamma six seven', LONG, 'delta eight nine']

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
  const body = paras.map((t) => `<w:p><w:r><w:t>${t}</w:t></w:r></w:p>`).join('')
  zip.file(
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}</w:body></w:document>`,
  )
  return zip.generateAsync({ type: 'nodebuffer' })
}

interface Sel {
  from: number
  to: number
  docSize: number
  text: string
  focused: boolean
}

async function selection(editor: Page): Promise<Sel> {
  return editor.evaluate(() => {
    const ed = (window as unknown as AidocsWindow).__aidocs!.editor as any
    const s = ed.state.selection
    return {
      from: s.from,
      to: s.to,
      docSize: ed.state.doc.content.size,
      text: ed.state.doc.textBetween(s.from, s.to, '\n'),
      focused: !!document.activeElement?.closest('.ProseMirror'),
    }
  })
}

async function boldText(editor: Page): Promise<string[]> {
  return editor.evaluate(() => {
    const ed = (window as unknown as AidocsWindow).__aidocs!.editor as any
    const out: string[] = []
    ed.state.doc.descendants(
      (n: { isText: boolean; text?: string; marks: { type: { name: string } }[] }) => {
        if (n.isText && n.marks.some((m) => m.type.name === 'bold')) out.push(n.text ?? '')
      },
    )
    return out
  })
}

async function setBold(editor: Page, word: string): Promise<void> {
  await editor.evaluate((w) => {
    const ed = (window as unknown as AidocsWindow).__aidocs!.editor as any
    let at = -1
    ed.state.doc.descendants((n: { isText: boolean; text?: string }, pos: number) => {
      if (at >= 0) return false
      const i = n.isText ? (n.text ?? '').indexOf(w) : -1
      if (i >= 0) at = pos + i
      return at < 0
    })
    ed.chain()
      .setTextSelection({ from: at, to: at + w.length })
      .setMark('bold')
      .run()
  }, word)
}

/** client center of `word` inside the paragraph containing it */
async function wordCenter(editor: Page, word: string): Promise<{ x: number; y: number }> {
  return editor.evaluate((w) => {
    const walker = document.createTreeWalker(
      document.querySelector('.doc-page')!,
      NodeFilter.SHOW_TEXT,
    )
    let tn: Node | null
    while ((tn = walker.nextNode())) {
      const i = (tn.textContent ?? '').indexOf(w)
      if (i < 0) continue
      const r = document.createRange()
      r.setStart(tn, i)
      r.setEnd(tn, i + w.length)
      const b = r.getBoundingClientRect()
      return { x: b.left + b.width / 2, y: b.top + b.height / 2 }
    }
    throw new Error(`word not found: ${w}`)
  }, word)
}

test.describe('docs selection bar and format painter', () => {
  let dir: string
  let docPath: string

  test.beforeEach(async () => {
    dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-selbar-')))
    docPath = join(dir, 'selbar.docx')
    writeFileSync(docPath, await minimalDocx(PARAS))
  })

  test.afterEach(() => {
    rmSync(dir, { recursive: true, force: true })
  })

  test('selection bar selects line, lines, paragraph and document', async () => {
    test.setTimeout(180_000)
    const launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'docs-selection-bar',
      openFile: docPath,
    })
    try {
      const editor = await waitForPageWithUrl(launched.app, '://docs/')
      await editor.waitForFunction(
        () => Boolean((window as unknown as AidocsWindow).__aidocs?.editor),
        undefined,
        { timeout: 30_000 },
      )
      const page = editor.locator('.doc-page')
      await expect(page).toContainText('delta eight nine')
      const pageBox = (await page.boundingBox())!
      const barX = pageBox.x + 30
      const rowY = async (text: string, dy = 8) => {
        const box = (await editor.locator('.doc-page p', { hasText: text }).first().boundingBox())!
        return box.y + dy
      }

      await editor.mouse.click(barX, await rowY('alpha one'))
      let sel = await selection(editor)
      expect(sel.text).toBe('alpha one two')
      expect(sel.focused).toBe(true)

      // first visual line of a wrapped paragraph, not the whole paragraph
      const longY = await rowY('lorem0 ')
      await editor.mouse.click(barX, longY)
      sel = await selection(editor)
      expect(sel.text.startsWith('lorem0 ')).toBe(true)
      expect(sel.text.length).toBeLessThan(LONG.length)
      expect(sel.text.length).toBeGreaterThan(20)

      await editor.mouse.dblclick(barX, longY)
      sel = await selection(editor)
      expect(sel.text).toBe(LONG)

      await editor.mouse.click(barX, longY, { clickCount: 3 })
      sel = await selection(editor)
      expect(sel.from).toBe(0)
      expect(sel.to).toBe(sel.docSize)

      // Shift+click extends the anchor to the clicked line
      await editor.mouse.click(barX, await rowY('alpha one'))
      await editor.keyboard.down('Shift')
      await editor.mouse.click(barX, await rowY('gamma six'))
      await editor.keyboard.up('Shift')
      sel = await selection(editor)
      expect(sel.text).toBe('alpha one two\nbeta four five\ngamma six seven')

      // drag down the bar
      await editor.mouse.move(barX, await rowY('beta four'))
      await editor.mouse.down()
      await editor.mouse.move(barX, await rowY('gamma six'), { steps: 4 })
      await editor.mouse.up()
      sel = await selection(editor)
      expect(sel.text).toBe('beta four five\ngamma six seven')

      // clicking the text area still places a caret (the bar is only the margin)
      const w = await wordCenter(editor, 'five')
      await editor.mouse.click(w.x, w.y)
      await expect.poll(async () => (await selection(editor)).text).toBe('')
    } finally {
      await closeAndSaveVideo(launched, 'docs-selection-bar')
    }
  })

  test('format painter brushes the clicked word, locks on double-click, has copy/paste chords', async () => {
    test.setTimeout(180_000)
    const launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'docs-format-painter',
      openFile: docPath,
    })
    try {
      const editor = await waitForPageWithUrl(launched.app, '://docs/')
      await editor.waitForFunction(
        () => Boolean((window as unknown as AidocsWindow).__aidocs?.editor),
        undefined,
        { timeout: 30_000 },
      )
      await expect(editor.locator('.doc-page')).toContainText('delta eight nine')
      await setBold(editor, 'beta')
      expect(await boldText(editor)).toEqual(['beta'])
      const painter = editor.getByRole('button', { name: /^Format Painter/ })
      // ProseMirror syncs a pointer-placed caret on a deferred selectionchange;
      // wait for it so the next chord reads the clicked position
      const clickWord = async (word: string) => {
        const c = await wordCenter(editor, word)
        await editor.mouse.click(c.x, c.y)
        await expect.poll(async () => (await selection(editor)).text).toBe('')
        await editor.waitForTimeout(50)
      }

      // single click: one stroke, on the word (with its trailing space), not the sentence
      await clickWord('beta')
      await painter.click()
      await expect(painter).toHaveAttribute('aria-pressed', 'true')
      await clickWord('six')
      await expect.poll(() => boldText(editor)).toEqual(['beta', 'six '])
      await expect(painter).toHaveAttribute('aria-pressed', 'false')

      // double-click: locked until Esc
      await clickWord('beta')
      await painter.dblclick()
      await expect(painter).toHaveAttribute('aria-pressed', 'true')
      await clickWord('one')
      await expect.poll(() => boldText(editor)).toEqual(['one ', 'beta', 'six '])
      await expect(painter).toHaveAttribute('aria-pressed', 'true')
      await clickWord('eight')
      await expect.poll(() => boldText(editor)).toEqual(['one ', 'beta', 'six ', 'eight '])
      await editor.keyboard.press('Escape')
      await expect(painter).toHaveAttribute('aria-pressed', 'false')
      await clickWord('nine')
      await editor.waitForTimeout(400)
      expect(await boldText(editor)).toEqual(['one ', 'beta', 'six ', 'eight '])

      // copy formatting at the caret, paste it onto the word under the caret
      const mac = process.platform === 'darwin'
      await clickWord('beta')
      await editor.keyboard.press(mac ? 'Meta+Shift+c' : 'Control+Shift+c')
      await clickWord('seven')
      await editor.keyboard.press(mac ? 'Meta+Shift+v' : 'Control+Alt+Shift+v')
      await expect.poll(() => boldText(editor)).toEqual(['one ', 'beta', 'six seven', 'eight '])
    } finally {
      await closeAndSaveVideo(launched, 'docs-format-painter')
    }
  })
})
