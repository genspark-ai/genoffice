import { readFile } from 'node:fs/promises'
import JSZip from 'jszip'
import { assertZipInflatesWithinLimits } from '@genoffice/zip-gate'

/**
 * JSZip over bytes the user supplied, behind the metered-inflate gate.
 *
 * The CLI reads archives directly by named entry (`xl/styles.xml`,
 * `xl/workbook.xml`, revision parts…), outside the engines' gated open
 * paths — the app's docx/xlsx loaders run `assertZipInflatesWithinLimits`
 * before anything inflates, these helpers did not. A bomb behind
 * honest-looking declarations would therefore materialise in full here.
 * The gate inflates each part one byte past its claim first, so the cost
 * tracks the claim instead of the payload (see #759).
 *
 * Only user-supplied archives go through here: generated ones
 * (`blankXlsxBuffer`, the CLI's own saves) are honest by construction.
 */
export async function loadUserZip(source: string | Uint8Array): Promise<JSZip> {
  const bytes = typeof source === 'string' ? await readFile(source) : source
  await assertZipInflatesWithinLimits(bytes)
  return JSZip.loadAsync(bytes)
}
