import { test, expect, type Page } from '@playwright/test'
import { existsSync, mkdtempSync, readFileSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import JSZip from 'jszip'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'

/**
 * Pins what every ribbon tab shows (groups, control names, which are disabled
 * or pressed) so a refactor of how ribbon controls are wired can prove the
 * ribbon itself did not change. Regenerate with UPDATE_E2E_BASELINES=1 when a
 * change is intended.
 */

const BASELINE = join(__dirname, 'baselines', 'docs-ribbon-structure.json')

interface RibbonControl {
  name: string
  tag: string
  disabled?: true
  pressed?: true
  popup?: true
}
type RibbonTab = { group: string; controls: RibbonControl[] }[]

async function minimalDocx(): Promise<Buffer> {
  const cell = (t: string) => `<w:tc><w:p><w:r><w:t>${t}</w:t></w:r></w:p></w:tc>`
  const zip = new JSZip()
  zip.file(
    '[Content_Types].xml',
    '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>',
  )
  zip.file(
    '_rels/.rels',
    '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
  )
  zip.file(
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Before</w:t></w:r></w:p><w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid><w:tr>${cell('A1')}${cell('B1')}</w:tr><w:tr>${cell('A2')}${cell('B2')}</w:tr></w:tbl><w:p><w:r><w:t>After</w:t></w:r></w:p></w:body></w:document>`,
  )
  return zip.generateAsync({ type: 'nodebuffer' })
}

const readTab = (page: Page): Promise<RibbonTab> =>
  page.evaluate(() => {
    const shown = (el: Element) => {
      const r = el.getBoundingClientRect()
      const s = getComputedStyle(el)
      return r.width > 0 && r.height > 0 && s.visibility !== 'hidden' && s.display !== 'none'
    }
    // a screen tip ends with the platform's shortcut chord, which is not structure
    const noChord = (s: string) => s.replace(/\s*\([^()]*\)\s*$/, '').trim()
    const textOf = (el: Element) => {
      const clone = el.cloneNode(true) as Element
      // icons carry no name, and a <small> subtitle shows current state, not what the control is
      clone
        .querySelectorAll(
          'svg, small, input, select, [aria-hidden=true], [class*=icon], [class*=glyph], [class*=caret]',
        )
        .forEach((n) => n.remove())
      const parts: string[] = []
      const walker = document.createTreeWalker(clone, NodeFilter.SHOW_TEXT)
      while (walker.nextNode()) parts.push(walker.currentNode.textContent ?? '')
      return parts.join(' ').replace(/\s+/g, ' ').trim()
    }
    const ownName = (el: Element) =>
      noChord(
        el.getAttribute('aria-label')?.trim() ||
          textOf(el) ||
          el.getAttribute('data-tip') ||
          el.getAttribute('title') ||
          '',
      )
    const nameOf = (el: Element): string => {
      const own = ownName(el)
      if (own) return own
      // a field named by the label that wraps it
      const label = el.closest('label')
      if (label) return noChord(textOf(label) || label.getAttribute('data-tip') || '')
      // the caret half of a split button, named after the half that carries the label
      const main = [...(el.closest('.rb-split-wrap')?.querySelectorAll('button') ?? [])].find(
        (b) => b !== el && ownName(b),
      )
      return main ? `${ownName(main)} menu` : ''
    }
    const body = document.querySelector('[data-ribbon-body]')
    if (!body) return []
    return [...body.querySelectorAll('.ribbon-group, .table-tool-group')].map((group) => ({
      group:
        group.querySelector('.ribbon-group-label, [class*=group-label]')?.textContent?.trim() ?? '',
      controls: [...group.querySelectorAll('button, select, input, [role=button], [role=combobox]')]
        // how many gallery cards fit depends on text metrics, not on structure
        .filter((el) => shown(el) && !el.matches('.table-style-card, .style-card'))
        .map((el) => ({
          name: nameOf(el),
          tag: el.tagName.toLowerCase(),
          ...((el as HTMLButtonElement).disabled || el.getAttribute('aria-disabled') === 'true'
            ? { disabled: true as const }
            : {}),
          ...(el.getAttribute('aria-pressed') === 'true' ? { pressed: true as const } : {}),
          ...(el.hasAttribute('aria-haspopup') || el.hasAttribute('aria-expanded')
            ? { popup: true as const }
            : {}),
        })),
    }))
  })

async function showTab(page: Page, name: string): Promise<void> {
  const tab = page.locator('.ribbon-tab', { hasText: new RegExp(`^${name}$`) }).first()
  // clicking the active tab collapses the ribbon, so only click an inactive one
  if (!/\bactive\b/.test((await tab.getAttribute('class')) ?? '')) await tab.click()
  await expect(page.locator('[data-ribbon-body] button').first()).toBeVisible()
}

test.describe('docs ribbon structure', () => {
  let dir: string

  test.beforeEach(() => {
    dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-ribbon-structure-')))
  })
  test.afterEach(() => {
    rmSync(dir, { recursive: true, force: true })
  })

  test('every ribbon tab matches its baseline', async () => {
    test.setTimeout(180_000)
    const docPath = join(dir, 'ribbon.docx')
    writeFileSync(docPath, await minimalDocx())
    const launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'docs-ribbon-structure',
      openFile: docPath,
      // the Zotero tab follows what the machine has installed; pin it on
      env: { GENOFFICE_ZOTERO: '1' },
    })
    const { app } = launched
    try {
      // one fixed width: the ribbon folds groups to fit
      await app.evaluate(({ BrowserWindow }) => {
        for (const win of BrowserWindow.getAllWindows()) win.setContentSize(1600, 900)
      })
      const page = await waitForPageWithUrl(app, '://docs/')
      await page.waitForFunction(
        () => Boolean((window as never as { __aidocs?: { editor?: unknown } }).__aidocs?.editor),
        undefined,
        { timeout: 30_000 },
      )
      const result: Record<string, RibbonTab> = {}
      await page.locator('.doc-page p', { hasText: 'Before' }).first().click()
      // Cut follows the selection a frame after the click: a caret alone leaves nothing to cut
      await expect(page.locator('[data-ribbon-body] button[aria-label^="Cut"]')).toBeDisabled()
      // the File tab and Home's Editing group exist only off macOS, where the
      // native menu bar does not carry them; the baseline is the macOS shape
      const standard = (await page.locator('.ribbon-tab').allTextContents())
        .map((t) => t.trim())
        .filter((name) => name !== 'File')
      for (const name of standard) {
        await showTab(page, name)
        result[name] = (await readTab(page)).filter((g) => g.group !== 'Editing')
      }
      const cell = page.locator('.doc-page .doc-table').first().locator('td, th').first()
      for (const name of ['Table Design', 'Table Layout']) {
        await cell.click()
        await showTab(page, name)
        result[name] = await readTab(page)
      }
      const actual = JSON.stringify(result, null, 2) + '\n'
      if (process.env.UPDATE_E2E_BASELINES || !existsSync(BASELINE)) writeFileSync(BASELINE, actual)
      expect(actual).toBe(readFileSync(BASELINE, 'utf8'))
    } finally {
      await closeAndSaveVideo(launched, 'docs-ribbon-structure')
    }
  })

  test('the Zotero tab stays away where Zotero is not installed', async () => {
    const docPath = join(dir, 'ribbon.docx')
    writeFileSync(docPath, await minimalDocx())
    const launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'docs-ribbon-no-zotero',
      openFile: docPath,
      env: { GENOFFICE_ZOTERO: '0' },
    })
    try {
      const page = await waitForPageWithUrl(launched.app, '://docs/')
      await expect(page.locator('.ribbon-tab', { hasText: /^View$/ })).toBeVisible()
      await expect(page.locator('.ribbon-tab', { hasText: /^Zotero$/ })).toHaveCount(0)
    } finally {
      await closeAndSaveVideo(launched, 'docs-ribbon-no-zotero')
    }
  })
})
