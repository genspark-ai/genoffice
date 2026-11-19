/**
 * The ribbon's main tabs greyed or live as Word has them on a blank document
 * with the caret in body text; the Review buttons that stay live with nothing
 * tracked; a comment on an empty paragraph, on the page and in the saved file.
 */
import { test, expect, type ElectronApplication, type Page } from '@playwright/test'
import { rmSync } from 'node:fs'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'
import { para, saveFromMenu } from './docs-objects-fixture'
import { open, part, ribbonTab, savedBody } from './docs-word-acceptance-kit'

const VIDEO_DIR = 'docs-ribbon-states-like-word'

const menuEnabled = (app: ElectronApplication, path: string[]) =>
  app.evaluate(({ Menu }, path) => {
    let items = Menu.getApplicationMenu()?.items
    let item: Electron.MenuItem | undefined
    for (const label of path) {
      item = items?.find((candidate) => candidate.label === label)
      items = item?.submenu?.items
    }
    return item ? item.enabled : null
  }, path)

const big = (page: Page, name: string | RegExp) =>
  page.locator('[data-ribbon-body] button:is(.rb-big, .rb-small)', { hasText: name })
const status = (page: Page) => page.locator('.status-bar .status-msg')
const dirty = (page: Page) =>
  page.locator('.qa-btn[data-tip^="Save"]').evaluate((b) => !(b as HTMLButtonElement).disabled)

/** each tab's controls on a blank document: true where Word has the control live */
const BLANK: Array<[string, Array<[string | RegExp, boolean, number?]>]> = [
  [
    'Insert',
    [
      [/^Comment$/, true],
      ['Drop Cap', false],
    ],
  ],
  ['Layout', [[/^Align$/, true]]],
  [
    'References',
    [
      ['Show Notes', false],
      [/^Update Table$/, true, 0],
      [/^Update Table$/, false, 1],
      [/^Update Index$/, false],
      [/^Update Table$/, false, 2],
    ],
  ],
  [
    'Mailings',
    [
      ['Start Mail Merge', true],
      ['Select Recipients', true],
      ['Edit Recipient List', false],
      ['Filter Recipients', false],
      ['Rules', false],
      ['Preview Results', false],
      ['Find Recipient', false],
      ['Finish & Merge', false],
    ],
  ],
  [
    'Review',
    [
      ['New Comment', true],
      ['Show Comments', false],
      [/^Accept$/, true],
      [/^Reject$/, true],
      ['Previous Change', true],
      ['Next Change', true],
    ],
  ],
  ['View', [['Merge All Windows', false]]],
]

test('a blank document greys and enables the ribbon as Word does, and the Review buttons say there is nothing tracked', async () => {
  test.setTimeout(180_000)
  const launched = await launchShell({ onboardingSeen: true, videoDir: VIDEO_DIR })
  const { page: shell, app } = launched
  try {
    await shell.evaluate(() =>
      (window as never as { aiOffice: { newDoc(): Promise<void> } }).aiOffice.newDoc(),
    )
    const page = await waitForPageWithUrl(app, '://docs/')
    await page.locator('main.editor-scroll .ProseMirror').first().click()

    for (const [tab, controls] of BLANK) {
      await ribbonTab(page, tab)
      for (const [name, live, nth] of controls) {
        const control = big(page, name).nth(nth ?? 0)
        await expect(control, `${tab} > ${name} ${nth ?? ''}`).toBeVisible()
        if (live) await expect(control, `${tab} > ${name} ${nth ?? ''}`).toBeEnabled()
        else await expect(control, `${tab} > ${name} ${nth ?? ''}`).toBeDisabled()
      }
    }
    expect(await menuEnabled(app, ['Insert', 'Comment'])).toBe(true)
    expect(await menuEnabled(app, ['Format', 'Drop Cap…'])).toBe(false)

    // Layout > Align opens with nothing to line up: its align items are greyed, the frame ones not
    await ribbonTab(page, 'Layout')
    await big(page, /^Align$/).click()
    const items = page.locator('[data-rb-panel].layout-menu button')
    for (const name of ['Align Left', 'Align Middle', 'Distribute Horizontally'])
      await expect(items.filter({ hasText: name })).toBeDisabled()
    await expect(items.filter({ hasText: 'Align to Page' })).toBeEnabled()
    await expect(items.filter({ hasText: 'Align Selected Objects' })).toBeDisabled()
    await page.screenshot({ path: test.info().outputPath('layout-align.png') })
    await page.keyboard.press('Escape')

    // Review: the four are live with nothing tracked, and say so without touching the document
    await ribbonTab(page, 'Review')
    expect(await dirty(page)).toBe(false)
    await big(page, 'Next Change').click()
    await expect(status(page)).toContainText('No tracked changes')
    await big(page, /^Accept$/).click()
    const accept = page.locator('[data-rb-panel].layout-menu button')
    await expect(accept.filter({ hasText: 'Accept and Move to Next' })).toBeEnabled()
    await expect(accept.filter({ hasText: 'Accept All Changes Shown' })).toBeDisabled()
    await expect(accept.filter({ hasText: 'Accept All in Selection' })).toBeDisabled()
    await page.screenshot({ path: test.info().outputPath('review-accept.png') })
    await accept.filter({ hasText: /^Accept All Changes$/ }).click()
    await expect(status(page)).toContainText('No tracked changes to process')
    await big(page, /^Reject$/).click()
    await page
      .locator('[data-rb-panel].layout-menu button', { hasText: 'Reject and Move to Next' })
      .click()
    await expect(status(page)).toContainText('No tracked changes to process')
    await big(page, 'Previous Change').click()
    await expect(status(page)).toContainText('No tracked changes')
    expect(await dirty(page)).toBe(false)

    // Drop Cap comes alive once the paragraph has a letter to drop
    await page.locator('main.editor-scroll .ProseMirror').first().click()
    await page.keyboard.type('Words')
    await ribbonTab(page, 'Insert')
    await expect(big(page, 'Drop Cap')).toBeEnabled()
    expect(await menuEnabled(app, ['Format', 'Drop Cap…'])).toBe(true)
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR)
  }
})

/** an empty paragraph holding the range of comment `id` (and of its replies) around nothing */
const EMPTY_RANGE = (id: string) =>
  new RegExp(
    '<w:p>(?:<w:pPr>(?:(?!</w:pPr>).)*</w:pPr>)?(?:<w:commentRangeStart w:id="\\d+"/>)*' +
      `<w:commentRangeStart w:id="${id}"/>(?:<w:commentRangeStart w:id="\\d+"/>)*` +
      '(?:<w:commentRangeEnd w:id="\\d+"/><w:r><w:commentReference w:id="\\d+"/></w:r>)+</w:p>',
  )

test('a comment on an empty paragraph is shown beside it, saved as a range around nothing and read back', async () => {
  test.setTimeout(240_000)
  let session = await open(
    VIDEO_DIR,
    para('above') + '<w:p/>' + para('below'),
    'empty-comment.docx',
  )
  const { dir, docPath } = session
  try {
    let { app, page } = session
    await expect(page.locator('.doc-page p', { hasText: 'below' })).toBeVisible({
      timeout: 30_000,
    })
    const empty = (p: Page) => p.locator('.doc-page[contenteditable="true"] > p').nth(1)
    const caretInEmpty = (p: Page) =>
      p.evaluate(() => {
        const { selection } = (
          window as never as {
            __aidocs: {
              editor: {
                state: {
                  selection: {
                    empty: boolean
                    $from: { parent: { content: { size: number } }; index(d: number): number }
                  }
                }
              }
            }
          }
        ).__aidocs.editor.state
        return (
          selection.empty && selection.$from.parent.content.size === 0 && selection.$from.index(0)
        )
      })
    await empty(page).click()
    await expect.poll(() => caretInEmpty(page)).toBe(1)

    // the three entry points are live by the one rule
    await ribbonTab(page, 'Insert')
    await expect(big(page, /^Comment$/)).toBeEnabled()
    expect(await menuEnabled(app, ['Insert', 'Comment'])).toBe(true)
    await empty(page).click({ button: 'right' })
    await expect(
      page.locator('.ctx-menu button.ctx-item', { hasText: 'New Comment' }),
    ).toBeEnabled()
    await page.keyboard.press('Escape')
    await expect.poll(() => caretInEmpty(page)).toBe(1)

    await ribbonTab(page, 'Review')
    await expect(big(page, 'Show Comments')).toBeDisabled()
    await big(page, 'New Comment').click()
    const pane = page.locator('.comments-pane')
    await pane.locator('.comment-compose textarea').fill('why is this line empty')
    await pane.locator('.comment-compose-actions button.primary').click()
    const thread = pane.locator('.comment-thread').first()
    await expect(thread).toContainText('why is this line empty')
    await expect(big(page, 'Show Comments')).toBeEnabled()
    // its icon stands in the margin at the empty paragraph's line
    const icon = page.locator('.page-margin-annotations .comment-balloon').first()
    await expect(icon).toBeVisible()
    const line = await empty(page).boundingBox()
    const mark = await icon.boundingBox()
    expect(Math.abs(mark!.y - line!.y)).toBeLessThan(line!.height)
    await page.screenshot({ path: test.info().outputPath('empty-paragraph-comment.png') })

    await thread.locator('.comment-reply-input').fill('a spacer, keep it')
    await thread.locator('.comment-reply-input').press('Enter')
    await expect(thread).toContainText('a spacer, keep it')

    // Enter in that paragraph: the comment stays with the first of the two, on the page and in the file
    await empty(page).click()
    await page.keyboard.press('Enter')
    await expect(page.locator('.doc-page[contenteditable="true"] > p')).toHaveCount(4)
    await expect(page.locator('.page-margin-annotations .comment-balloon')).toHaveCount(1)
    await expect
      .poll(async () =>
        Math.abs((await icon.boundingBox())!.y - (await empty(page).boundingBox())!.y),
      )
      .toBeLessThan(line!.height)

    await saveFromMenu(app)
    await expect
      .poll(
        async () => {
          const body = await savedBody(docPath)
          const comments = await part(docPath, 'word/comments.xml')
          const id =
            /<w:comment\b[^>]*w:id="(\d+)"[^>]*>(?:(?!<\/w:comment>).)*why is this line empty/.exec(
              comments,
            )?.[1]
          return {
            comment: !!id,
            reply: comments.includes('a spacer, keep it'),
            aroundNothing: !!id && EMPTY_RANGE(id).test(body),
            starts: (body.match(/<w:commentRangeStart\b/g) ?? []).length,
            references: (body.match(/<w:commentReference\b/g) ?? []).length,
            text: body.includes('above') && body.includes('below'),
          }
        },
        { timeout: 30_000 },
      )
      .toEqual({
        comment: true,
        reply: true,
        aroundNothing: true,
        starts: 2,
        references: 2,
        text: true,
      })

    // reopened: the thread is there at the same paragraph, and deleting it takes the range out
    await closeAndSaveVideo(session.launched, VIDEO_DIR)
    const launched = await launchShell({
      onboardingSeen: true,
      videoDir: VIDEO_DIR,
      openFile: docPath,
    })
    session = { ...session, launched }
    app = launched.app
    page = await waitForPageWithUrl(app, '://docs/')
    await expect(page.locator('.doc-page p', { hasText: 'below' })).toBeVisible({
      timeout: 30_000,
    })
    const reopened = page.locator('.page-margin-annotations .comment-bubble').first()
    await expect(reopened).toContainText('why is this line empty')
    await expect(reopened).toContainText('a spacer, keep it')
    await expect(page.locator('.page-margin-annotations .comment-balloon')).toHaveCount(1)
    await page.screenshot({ path: test.info().outputPath('empty-paragraph-comment-reopened.png') })

    // text typed into the paragraph is saved inside the range
    await empty(page).click()
    await page.keyboard.type('filled')
    await saveFromMenu(app)
    await expect
      .poll(
        async () =>
          /<w:commentRangeStart w:id="\d+"\/>.*filled.*<w:commentRangeEnd\b/.test(
            await savedBody(docPath),
          ),
        {
          timeout: 30_000,
        },
      )
      .toBe(true)

    await ribbonTab(page, 'Review')
    await big(page, /^Delete$/).click()
    await page
      .locator('[data-rb-panel].layout-menu button', { hasText: 'Delete All Comments' })
      .click()
    await expect(page.locator('.page-margin-annotations .comment-bubble')).toHaveCount(0)
    await expect(big(page, 'Show Comments')).toBeDisabled()
    await saveFromMenu(app)
    await expect
      .poll(
        async () => {
          const body = await savedBody(docPath)
          return [/<w:comment(RangeStart|RangeEnd|Reference)\b/.test(body), body.includes('filled')]
        },
        { timeout: 30_000 },
      )
      .toEqual([false, true])
  } finally {
    await closeAndSaveVideo(session.launched, VIDEO_DIR).catch(() => {})
    rmSync(dir, { recursive: true, force: true })
  }
})
