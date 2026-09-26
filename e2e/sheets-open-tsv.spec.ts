import { test, expect } from '@playwright/test'
import { mkdtemp, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import type { Page } from '@playwright/test'
import { closeAndSaveVideo, launchShell, screenshotPath, waitForPageWithUrl } from './helpers'

/**
 * Issue #1140: a .tsv dragged into the app must land in Sheets as a table.
 * The sample is UniProt-shaped on purpose — its annotation fields are
 * comma-separated lists, so a delimiter sniffer counting raw occurrences
 * picks the comma over the tabs and shatters every row.
 */
const TSV =
  'Entry\tProtein names\tGene names\tLength\r\n' +
  'P38398\tBreast cancer protein 1, RING finger E3 ubiquitin protein ligase, BRC1, BRCA1, PARSUM1\tBRCA1, RNF3, c.1125dupT\t1915\r\n' +
  'Q92766\tSentrin/SUMO-specific protease 6, GCP-2, PIASY\tSENP6, SENP7, GCP2\t650\r\n'

async function cell(sheets: Page, reference: string): Promise<unknown> {
  return sheets.evaluate((a1) => {
    const debug = (window as unknown as Record<string, unknown>).__genofficeDebug as {
      univerAPI: {
        getActiveWorkbook(): {
          getActiveSheet(): { getRange(a: string): { getValue(): unknown } } | null
        } | null
      }
    }
    const workbook = debug.univerAPI.getActiveWorkbook()
    return workbook?.getActiveSheet()?.getRange(a1)?.getValue() ?? null
  }, reference)
}

async function waitForWorkbook(page: Page): Promise<void> {
  await page.waitForFunction(() => document.body.textContent?.includes('Sheet1'), null, {
    timeout: 30_000,
  })
  await page.waitForTimeout(1_500)
}

test.describe('sheets: opening a .tsv', () => {
  test('splits on tabs, keeps comma-heavy annotations in one cell, and saves as .xlsx', async () => {
    const scratch = await mkdtemp(join(tmpdir(), 'genoffice-tsv-open-e2e-'))
    const source = join(scratch, 'uniprot.tsv')
    await writeFile(source, TSV)

    const launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'sheets-open-tsv',
      openFile: source,
      env: { GENOFFICE_DEBUG_HOOKS: '1' },
    })
    try {
      const sheets = await waitForPageWithUrl(launched.app, '://sheets/')
      await waitForWorkbook(sheets)

      // four columns, exactly where the tabs are
      expect(await cell(sheets, 'A1')).toBe('Entry')
      expect(await cell(sheets, 'B1')).toBe('Protein names')
      expect(await cell(sheets, 'C1')).toBe('Gene names')
      expect(await cell(sheets, 'D1')).toBe('Length')
      // the comma-separated list survives whole in its own cell
      expect(await cell(sheets, 'C2')).toBe('BRCA1, RNF3, c.1125dupT')
      expect(await cell(sheets, 'B3')).toBe('Sentrin/SUMO-specific protease 6, GCP-2, PIASY')
      // nothing spilled past column D: the comma sniff would have shattered here
      expect(await cell(sheets, 'E1')).toBeFalsy()
      expect(await cell(sheets, 'E2')).toBeFalsy()
      // values-only import still types plain numbers as numbers
      expect(Number(await cell(sheets, 'D2'))).toBe(1915)

      const shell = await waitForPageWithUrl(launched.app, 'shell/out')
      await expect(shell.locator('.tab-title').filter({ hasText: 'uniprot.tsv' })).toBeVisible()
      await expect(sheets.locator('.workbook-status')).not.toContainText('Error invoking')
      await expect(sheets.locator('.status-bar')).not.toContainText('Error invoking')

      await sheets.screenshot({ path: screenshotPath('sheets-open-tsv') })
    } finally {
      await closeAndSaveVideo(launched, 'sheets-open-tsv')
    }
  })
})
