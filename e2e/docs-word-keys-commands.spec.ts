/**
 * The editing commands behind chords of Word for Mac, pressed on a real
 * document: the selection mode of F8, locking, unlocking and unlinking fields
 * and moving between them, promoting and demoting headings, the ends of a
 * table row and column, the Spike and Go Back, with the saved file read back
 * for the ones that edit.
 */
import { test, expect, type Page } from '@playwright/test'
import { mkdtempSync, readFileSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import JSZip from 'jszip'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl, type LaunchedApp } from './helpers'
import { docxWith, para, saveFromMenu } from './docs-objects-fixture'

const VIDEO_DIR = 'docs-word-keys-commands'
const IS_MAC = process.platform === 'darwin'
const SKIP = 'these chords are Word for Mac’s; other platforms keep their own'

const field = (instr: string, result: string, attrs = '') =>
  `<w:r><w:fldChar w:fldCharType="begin"${attrs}/></w:r>` +
  `<w:r><w:instrText xml:space="preserve"> ${instr} </w:instrText></w:r>` +
  '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
  `<w:r><w:t>${result}</w:t></w:r>` +
  '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
const run = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`
const cell = (text: string) => `<w:tc><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`
const row = (n: number) =>
  `<w:tr>${cell(`A${n}`)}${cell(`B${n} more words`)}${cell(`C${n}`)}</w:tr>`
const TABLE =
  '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>' +
  '<w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid>' +
  `${row(1)}${row(2)}${row(3)}</w:tbl>`

const SENTENCES = 'One two three. Four five six.'
const FIELDS =
  `<w:p>${run('Time ')}${field('TIME', 'old-time')}${run(' date ')}` +
  `${field('DATE', 'keep-date', ' w:fldLock="1"')}${run(' file ')}${field('FILENAME', 'old-name')}${run(' end')}</w:p>`
const BODY =
  para('Chapter') +
  para(SENTENCES) +
  para('Section') +
  para('alpha beta gamma') +
  FIELDS +
  TABLE +
  para('Last paragraph')

interface FieldSnap {
  text: string
  instr: string
  locked: boolean
}

interface Snap {
  blocks: string[]
  /** `h2` for a heading of level 2, `p` otherwise, per block */
  kinds: string[]
  /** index of the block the selection's head is in */
  head: number
  headOffset: number
  selected: string
  cells: boolean
  /** how many ranges the selection has: one per cell for a cell selection */
  ranges: number
  /** the node selected as a whole, when one is */
  node: string | null
  fields: FieldSnap[]
}

const snap = (page: Page): Promise<Snap> =>
  page.evaluate(() => {
    interface Node {
      isTextblock: boolean
      isText: boolean
      text?: string
      textContent: string
      type: { name: string }
      attrs: Record<string, unknown>
      marks: { type: { name: string }; attrs: Record<string, unknown> }[]
    }
    const { editor } = (
      window as never as {
        __aidocs: {
          editor: {
            state: {
              doc: {
                descendants(visit: (node: Node, pos: number) => void): void
                textBetween(from: number, to: number, sep: string): string
              }
              selection: {
                head: number
                from: number
                to: number
                $anchorCell?: unknown
                ranges: unknown[]
                node?: { type: { name: string } }
              }
            }
          }
        }
      }
    ).__aidocs
    const { doc, selection } = editor.state
    const starts: number[] = []
    const blocks: string[] = []
    const kinds: string[] = []
    const fields: { text: string; instr: string; locked: boolean }[] = []
    doc.descendants((node, pos) => {
      if (node.isTextblock) {
        starts.push(pos + 1)
        blocks.push(node.textContent)
        kinds.push(node.type.name === 'docHeading' ? `h${String(node.attrs.level)}` : 'p')
      }
      const mark = node.isText ? node.marks.find((m) => m.type.name === 'instrField') : undefined
      if (mark)
        fields.push({
          text: node.text ?? '',
          instr: String(mark.attrs.instr).trim(),
          locked: mark.attrs.locked === true,
        })
    })
    let head = 0
    for (let i = 0; i < starts.length; i++) if (starts[i] <= selection.head) head = i
    return {
      blocks,
      kinds,
      head,
      headOffset: selection.head - starts[head],
      selected: doc.textBetween(selection.from, selection.to, '\n'),
      cells: selection.$anchorCell !== undefined,
      ranges: selection.ranges.length,
      node: selection.node?.type.name ?? null,
      fields,
    }
  })

/** a caret (or `length` characters selected) at `offset` of `needle` in the block that contains it */
const selectAt = (page: Page, needle: string, offset = 0, length = 0) =>
  page.evaluate(
    ({ needle, offset, length }) => {
      interface Node {
        isTextblock: boolean
        textContent: string
      }
      const { editor } = (
        window as never as {
          __aidocs: {
            editor: {
              state: { doc: { descendants(visit: (node: Node, pos: number) => void): void } }
              chain(): {
                focus(): {
                  setTextSelection(range: { from: number; to: number }): { run(): boolean }
                }
              }
            }
          }
        }
      ).__aidocs
      let at = -1
      editor.state.doc.descendants((node, pos) => {
        const found = node.isTextblock ? node.textContent.indexOf(needle) : -1
        if (at < 0 && found >= 0) at = pos + 1 + found
      })
      if (at < 0) throw new Error(`no paragraph contains "${needle}"`)
      editor
        .chain()
        .focus()
        .setTextSelection({ from: at + offset, to: at + offset + length })
        .run()
    },
    { needle, offset, length },
  )

const userData = () => realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-')))

async function open(docPath: string, userDataDir: string) {
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: VIDEO_DIR,
    openFile: docPath,
    userDataDir,
  })
  const page = await waitForPageWithUrl(launched.app, '://docs/')
  await page.waitForFunction(
    () => Boolean((window as unknown as { __aidocs?: { editor?: unknown } }).__aidocs?.editor),
    undefined,
    { timeout: 30_000 },
  )
  await expect(page.locator('.doc-page p', { hasText: 'Last paragraph' })).toBeVisible({
    timeout: 30_000,
  })
  return { launched, page }
}

async function newDoc() {
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-word-keys-')))
  const docPath = join(dir, 'fields.docx')
  writeFileSync(docPath, await docxWith(BODY))
  return { dir, docPath }
}

async function close(launched: LaunchedApp, ...dirs: string[]) {
  await closeAndSaveVideo(launched, VIDEO_DIR)
  for (const dir of dirs) rmSync(dir, { recursive: true, force: true })
}

const documentXml = async (docPath: string) =>
  (await JSZip.loadAsync(readFileSync(docPath))).file('word/document.xml')!.async('string')

const fieldTexts = async (page: Page) => (await snap(page)).fields.map((f) => f.text)

test('F8 extends the selection by unit, by arrow, by character and by click until Escape', async () => {
  test.skip(!IS_MAC, SKIP)
  test.setTimeout(180_000)
  const userDataDir = userData()
  const { dir, docPath } = await newDoc()
  const { launched, page } = await open(docPath, userDataDir)
  try {
    await selectAt(page, 'two', 1)
    await page.keyboard.press('F8')
    expect((await snap(page)).selected).toBe('')
    await page.keyboard.press('F8')
    expect((await snap(page)).selected).toBe('two')
    await page.keyboard.press('F8')
    expect((await snap(page)).selected).toBe('One two three. ')
    await page.keyboard.press('F8')
    expect((await snap(page)).selected).toBe(SENTENCES)
    await page.keyboard.press('F8')
    let now = await snap(page)
    expect(now.selected.startsWith('Chapter\n')).toBe(true)
    expect(now.selected.endsWith('Last paragraph')).toBe(true)
    // Shift+F8 goes back unit by unit, down to the caret
    await page.keyboard.press('Shift+F8')
    expect((await snap(page)).selected).toBe(SENTENCES)
    await page.keyboard.press('Shift+F8')
    await page.keyboard.press('Shift+F8')
    expect((await snap(page)).selected).toBe('two')
    await page.keyboard.press('Shift+F8')
    now = await snap(page)
    expect(now.selected).toBe('')
    expect([now.head, now.headOffset]).toEqual([1, 5])

    // the mode is still on: arrows extend without Shift, a character extends to its next occurrence
    await page.keyboard.press('ArrowRight')
    await page.keyboard.press('ArrowRight')
    expect((await snap(page)).selected).toBe('wo')
    await page.keyboard.press('Alt+ArrowRight')
    expect((await snap(page)).selected).toBe('wo three')
    await page.keyboard.press('ArrowLeft')
    expect((await snap(page)).selected).toBe('wo thre')
    await page.keyboard.press('f')
    now = await snap(page)
    expect(now.selected).toBe('wo three. Four f')
    expect(now.blocks[1]).toBe(SENTENCES)
    // the caret moves of the key table extend too
    await page.keyboard.press('End')
    expect((await snap(page)).selected).toBe('wo three. Four five six.')
    await page.keyboard.press('Meta+ArrowDown')
    expect((await snap(page)).selected).toBe('wo three. Four five six.\n')

    // a click extends to where it lands
    const target = page.locator('.doc-page p', { hasText: 'alpha beta gamma' })
    const box = (await target.boundingBox())!
    await page.mouse.click(box.x + 2, box.y + box.height / 2)
    now = await snap(page)
    expect(now.selected).toBe('wo three. Four five six.\nSection\n')
    expect(now.head).toBe(3)

    // Escape leaves the mode with the selection kept; an arrow then collapses it
    await page.keyboard.press('Escape')
    expect((await snap(page)).selected).toBe('wo three. Four five six.\nSection\n')
    await page.keyboard.press('ArrowRight')
    await expect.poll(async () => (await snap(page)).selected).toBe('')
    now = await snap(page)
    expect(now.blocks.slice(0, 4)).toEqual(['Chapter', SENTENCES, 'Section', 'alpha beta gamma'])

    // in a table cell the units are the cell's own
    await selectAt(page, 'B2 more', 4)
    await page.keyboard.press('F8')
    await page.keyboard.press('F8')
    expect((await snap(page)).selected).toBe('more')
    await page.keyboard.press('F8')
    await page.keyboard.press('F8')
    expect((await snap(page)).selected).toBe('B2 more words')
    // typing after an edit key ends the mode: Delete removes the selection and the next letter is typed
    await page.keyboard.press('Shift+F8')
    await page.keyboard.press('Shift+F8')
    expect((await snap(page)).selected).toBe('more')
    await page.keyboard.press('Delete')
    await page.keyboard.type('less')
    now = await snap(page)
    expect(now.blocks).toContain('B2 less words')
    expect(now.selected).toBe('')
    await page.keyboard.press('Meta+KeyZ')
    await page.keyboard.press('Meta+KeyZ')
    await expect.poll(async () => (await snap(page)).blocks).toContain('B2 more words')
  } finally {
    await close(launched, dir, userDataDir)
  }
})

test('fields lock, unlock and unlink from the keyboard, and a locked field keeps its result through an update, a save and a reopen', async () => {
  test.skip(!IS_MAC, SKIP)
  test.setTimeout(300_000)
  const userDataDir = userData()
  const { dir, docPath } = await newDoc()
  let { launched, page } = await open(docPath, userDataDir)
  try {
    // the lock the file carries is read
    expect((await snap(page)).fields).toEqual([
      { text: 'old-time', instr: 'TIME', locked: false },
      { text: 'keep-date', instr: 'DATE', locked: true },
      { text: 'old-name', instr: 'FILENAME', locked: false },
    ])

    // F11 / Shift+F11 select the next and the previous field
    await selectAt(page, 'Chapter')
    await page.keyboard.press('F11')
    expect((await snap(page)).selected).toBe('old-time')
    await page.keyboard.press('F11')
    expect((await snap(page)).selected).toBe('keep-date')
    await page.keyboard.press('F11')
    expect((await snap(page)).selected).toBe('old-name')
    await page.keyboard.press('F11')
    expect((await snap(page)).selected).toBe('old-name')
    await page.keyboard.press('Shift+F11')
    expect((await snap(page)).selected).toBe('keep-date')

    // Cmd+F11 locks the file name; F9 then updates the time alone
    await selectAt(page, 'old-name', 3)
    await page.keyboard.press('Meta+F11')
    expect((await snap(page)).fields.map((f) => f.locked)).toEqual([false, true, true])
    await page.keyboard.press('F9')
    await expect.poll(async () => (await fieldTexts(page))[0]).not.toBe('old-time')
    expect((await fieldTexts(page)).slice(1)).toEqual(['keep-date', 'old-name'])
    const time = (await fieldTexts(page))[0]

    // Shift+Cmd+F11 unlocks the date, which F9 then updates; Cmd+3 locks it again and Cmd+Z undoes that
    await selectAt(page, 'keep-date', 2)
    await page.keyboard.press('Shift+Meta+F11')
    expect((await snap(page)).fields.map((f) => f.locked)).toEqual([false, false, true])
    await page.keyboard.press('F9')
    await expect.poll(async () => (await fieldTexts(page))[1]).not.toBe('keep-date')
    const date = (await fieldTexts(page))[1]
    expect((await fieldTexts(page))[2]).toBe('old-name')
    await selectAt(page, date, 1)
    await page.keyboard.press('Meta+Digit3')
    expect((await snap(page)).fields.map((f) => f.locked)).toEqual([false, true, true])
    await page.keyboard.press('Meta+KeyZ')
    expect((await snap(page)).fields.map((f) => f.locked)).toEqual([false, false, true])
    await page.keyboard.press('Shift+Meta+KeyZ')
    expect((await snap(page)).fields.map((f) => f.locked)).toEqual([false, true, true])
    // Cmd+4 unlocks and Cmd+3 locks, as the F11 pair does
    await page.keyboard.press('Meta+Digit4')
    expect((await snap(page)).fields.map((f) => f.locked)).toEqual([false, false, true])
    await page.keyboard.press('Meta+Digit3')

    // Shift+Cmd+F9 turns the time into its text; Cmd+Z brings the field back, Cmd+6 unlinks again
    await selectAt(page, time, 1)
    await page.keyboard.press('Shift+Meta+F9')
    let now = await snap(page)
    expect(now.fields.map((f) => f.instr)).toEqual(['DATE', 'FILENAME'])
    expect(now.blocks[4]).toBe(`Time ${time} date ${date} file old-name end`)
    await page.keyboard.press('Meta+KeyZ')
    expect((await snap(page)).fields.map((f) => f.instr)).toEqual(['TIME', 'DATE', 'FILENAME'])
    await selectAt(page, time, 1)
    await page.keyboard.press('Meta+Digit6')
    now = await snap(page)
    expect(now.fields.map((f) => f.instr)).toEqual(['DATE', 'FILENAME'])
    // unlinking never updates the others
    expect(now.fields.map((f) => f.text)).toEqual([date, 'old-name'])

    await saveFromMenu(launched.app)
    await expect
      .poll(async () => (await documentXml(docPath)).match(/w:fldLock="true"/g)?.length ?? 0, {
        timeout: 30_000,
      })
      .toBe(2)
    const xml = await documentXml(docPath)
    expect(xml).not.toContain(' TIME ')
    expect(xml).toContain(time)
    expect(xml).toMatch(/w:fldLock="true"\/><\/w:r><w:r><w:instrText[^>]*> ?DATE ?</)
    expect(xml).toMatch(/w:fldLock="true"\/><\/w:r><w:r><w:instrText[^>]*> ?FILENAME ?</)
    expect(xml).toContain('old-name')

    // reopened: both locks are read back and an update leaves both results alone
    await closeAndSaveVideo(launched, VIDEO_DIR)
    ;({ launched, page } = await open(docPath, userDataDir))
    expect((await snap(page)).fields).toEqual([
      { text: date, instr: 'DATE', locked: true },
      { text: 'old-name', instr: 'FILENAME', locked: true },
    ])
    await selectAt(page, 'Chapter')
    await page.keyboard.press('F9')
    await page.keyboard.press('Alt+Shift+Meta+KeyU')
    expect(await fieldTexts(page)).toEqual([date, 'old-name'])
    expect((await snap(page)).blocks[4]).toBe(`Time ${time} date ${date} file old-name end`)
  } finally {
    await close(launched, dir, userDataDir)
  }
})

test('headings promote and demote, table keys reach the ends of the row and column, the Spike collects and Go Back returns', async () => {
  test.skip(!IS_MAC, SKIP)
  test.setTimeout(300_000)
  const userDataDir = userData()
  const { dir, docPath } = await newDoc()
  const { launched, page } = await open(docPath, userDataDir)
  try {
    // ---- outline levels ----
    // body text with no heading before it becomes Heading 1 on Promote
    await selectAt(page, 'Chapter', 2)
    await page.keyboard.press('Control+Shift+ArrowLeft')
    expect((await snap(page)).kinds.slice(0, 4)).toEqual(['h1', 'p', 'p', 'p'])
    // Demote on body text: one level below the heading before it
    await selectAt(page, 'Section', 2)
    await page.keyboard.press('Control+Shift+ArrowRight')
    expect((await snap(page)).kinds.slice(0, 4)).toEqual(['h1', 'p', 'h2', 'p'])
    await page.keyboard.press('Control+Shift+ArrowRight')
    expect((await snap(page)).kinds[2]).toBe('h3')
    await page.keyboard.press('Control+Shift+ArrowLeft')
    await page.keyboard.press('Control+Shift+ArrowLeft')
    expect((await snap(page)).kinds[2]).toBe('h1')
    await page.keyboard.press('Control+Shift+ArrowLeft')
    expect((await snap(page)).kinds[2]).toBe('h1')
    await page.keyboard.press('Meta+KeyZ')
    let now = await snap(page)
    expect(now.kinds.slice(0, 4)).toEqual(['h1', 'p', 'h2', 'p'])
    expect(now.blocks.slice(0, 4)).toEqual(['Chapter', SENTENCES, 'Section', 'alpha beta gamma'])

    // Outline view: Control+Shift+- folds what is under the heading, Control+Shift+= opens it
    await selectAt(page, 'Chapter', 2)
    await page.keyboard.press('Control+Shift+Minus')
    await expect(page.locator('.outline-collapsed')).toHaveCount(0)
    await page.keyboard.press('Alt+Meta+KeyO')
    await expect(page.locator('.outline-handle').first()).toBeVisible()
    await selectAt(page, 'Chapter', 2)
    await page.keyboard.press('Control+Shift+Minus')
    await expect(page.locator('.outline-collapsed')).not.toHaveCount(0)
    await page.keyboard.press('Control+Shift+Equal')
    await expect(page.locator('.outline-collapsed')).toHaveCount(0)
    await page.keyboard.press('Alt+Meta+KeyP')
    await expect(page.locator('.outline-handle')).toHaveCount(0)
    expect((await snap(page)).kinds.slice(0, 4)).toEqual(['h1', 'p', 'h2', 'p'])

    // ---- table ----
    const cellAt = async () => {
      const s = await snap(page)
      return [s.blocks[s.head], s.selected, s.cells] as const
    }
    await selectAt(page, 'B2 more', 4)
    await page.keyboard.press('Control+Home')
    expect(await cellAt()).toEqual(['A2', '', false])
    await page.keyboard.press('Control+End')
    expect(await cellAt()).toEqual(['C2', '', false])
    await page.keyboard.press('Control+PageUp')
    expect(await cellAt()).toEqual(['C1', '', false])
    await page.keyboard.press('Control+PageDown')
    expect(await cellAt()).toEqual(['C3', '', false])
    await page.keyboard.press('Alt+Shift+Home')
    expect(await cellAt()).toEqual(['A3', '', false])
    await page.keyboard.press('Alt+Shift+PageUp')
    expect(await cellAt()).toEqual(['A1', '', false])
    await page.keyboard.press('Alt+Shift+End')
    expect(await cellAt()).toEqual(['C1', '', false])
    await page.keyboard.press('Alt+Shift+PageDown')
    expect(await cellAt()).toEqual(['C3', '', false])
    // Option+Clear, the keypad's key, selects the table
    await selectAt(page, 'B2 more', 4)
    await page.keyboard.press('Alt+Numpad5')
    expect((await snap(page)).node).toBe('docTable')
    // outside a table the keys move nothing
    await selectAt(page, 'Last paragraph', 4)
    await page.keyboard.press('Control+PageUp')
    now = await snap(page)
    expect([now.blocks[now.head], now.headOffset]).toEqual(['Last paragraph', 4])

    // ---- the Spike ----
    await selectAt(page, 'alpha beta gamma', 0, 6)
    await page.keyboard.press('Meta+F3')
    await selectAt(page, 'beta gamma', 5, 5)
    await page.keyboard.press('Meta+F3')
    expect((await snap(page)).blocks[3]).toBe('beta ')
    await selectAt(page, 'A1', 2)
    await page.keyboard.press('Meta+F3')
    await page.keyboard.press('Shift+Meta+F3')
    now = await snap(page)
    expect(now.blocks).toContain('A1alpha ')
    expect(now.blocks).toContain('gamma')
    // emptied: a second press inserts nothing, and it never changed the case of anything
    const before = now.blocks
    await page.keyboard.press('Shift+Meta+F3')
    expect((await snap(page)).blocks).toEqual(before)
    await page.keyboard.press('Meta+KeyZ')
    now = await snap(page)
    expect(now.blocks).toContain('A1')
    expect(now.blocks).not.toContain('gamma')
    // into the last paragraph for the saved file
    await selectAt(page, 'beta ', 0, 5)
    await page.keyboard.press('Meta+F3')
    await selectAt(page, 'Last paragraph', 14)
    await page.keyboard.press('Shift+Meta+F3')
    expect((await snap(page)).blocks).toContain('Last paragraphbeta ')

    // ---- Go Back ----
    await selectAt(page, 'One two', 3)
    await page.keyboard.type('!')
    await selectAt(page, 'C3', 2)
    await page.keyboard.type('?')
    await selectAt(page, 'Chapter')
    await page.keyboard.press('Shift+F5')
    now = await snap(page)
    expect([now.blocks[now.head], now.headOffset]).toEqual(['C3?', 3])
    await page.keyboard.press('Alt+Meta+KeyZ')
    now = await snap(page)
    expect([now.blocks[now.head], now.headOffset]).toEqual(['One! two three. Four five six.', 4])
    await page.keyboard.press('Shift+F5')
    expect((await snap(page)).blocks[(await snap(page)).head]).toBe('Last paragraphbeta ')
    await page.keyboard.press('Shift+F5')
    now = await snap(page)
    expect([now.blocks[now.head], now.headOffset]).toEqual(['Chapter', 0])

    await saveFromMenu(launched.app)
    await expect
      .poll(async () => (await documentXml(docPath)).includes('Last paragraphbeta'), {
        timeout: 30_000,
      })
      .toBe(true)
    const xml = await documentXml(docPath)
    const styleOf = (text: string) =>
      /<w:pStyle w:val="([^"]+)"/.exec(
        xml.slice(xml.lastIndexOf('<w:p', xml.indexOf(`>${text}<`)), xml.indexOf(`>${text}<`)),
      )?.[1] ?? null
    expect(styleOf('Chapter')).toMatch(/^Heading\s?1$|^1$/i)
    expect(styleOf('Section')).toMatch(/^Heading\s?2$|^2$/i)
    expect(styleOf('Last paragraphbeta ')).toBeNull()
    expect(xml).toContain('One! two three. Four five six.')
    expect(xml).not.toContain('alpha')
  } finally {
    await close(launched, dir, userDataDir)
  }
})

test('a Customize Keyboard assignment takes a chord from the new commands', async () => {
  test.skip(!IS_MAC, SKIP)
  test.setTimeout(180_000)
  const userDataDir = userData()
  // as Customize Keyboard stores it: Word Count takes F8, Lock Fields moves off Cmd+3
  writeFileSync(
    join(userDataDir, 'docs-keyboard.json'),
    JSON.stringify({
      'word-count': ['F8'],
      'extend-selection:grow': [],
      'field-lock:lock': ['Cmd+F11'],
      'table-cell-go:row-end-extend': ['Ctrl+Shift+End'],
    }),
  )
  const { dir, docPath } = await newDoc()
  const { launched, page } = await open(docPath, userDataDir)
  try {
    await selectAt(page, 'two', 1)
    await page.keyboard.press('F8')
    const wordCount = page.locator('.word-count-dialog')
    await expect(wordCount).toBeVisible()
    await page.keyboard.press('Escape')
    await expect(wordCount).toBeHidden()
    await selectAt(page, 'two', 1)
    await page.keyboard.press('ArrowRight')
    expect((await snap(page)).selected).toBe('')
    // the chord Lock Fields gave up does nothing; the one it kept and its neighbours work
    await selectAt(page, 'old-name', 2)
    await page.keyboard.press('Meta+Digit3')
    expect((await snap(page)).fields.map((f) => f.locked)).toEqual([false, true, false])
    await page.keyboard.press('Meta+F11')
    expect((await snap(page)).fields.map((f) => f.locked)).toEqual([false, true, true])
    await page.keyboard.press('Meta+Digit4')
    expect((await snap(page)).fields.map((f) => f.locked)).toEqual([false, true, false])
    // selecting to the end of the row has no chord of its own: the one given here runs it
    await selectAt(page, 'B2 more', 4)
    await page.keyboard.press('Control+Shift+End')
    const now = await snap(page)
    expect([now.cells, now.ranges, now.blocks[now.head]]).toEqual([true, 2, 'C2'])
  } finally {
    await close(launched, dir, userDataDir)
  }
})
