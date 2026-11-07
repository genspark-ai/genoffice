import { mkdtemp, readFile, rm } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { exportDocsHeadless } from '../../../docs/src/main/docs-main'

/**
 * The browser tab cannot print its live DOM (printToPDF belongs to an Electron
 * WebContents), and only the docs editor prints that way — every other app
 * exports through hidden windows already. The delegate re-renders the saved
 * file through the exact hidden-window pipeline the CLI uses, so browser and
 * CLI PDF output cannot drift. Unsaved documents have nothing to re-render
 * from and say so.
 */
export async function webPrintToPDF(
  sourceUrl: string,
  _options: Record<string, unknown>,
): Promise<Buffer> {
  let path: string | null
  try {
    path = new URL(sourceUrl, 'http://localhost').searchParams.get('path')
  } catch {
    path = null
  }
  if (!path) {
    throw new Error(
      'Save the document (Ctrl+S) before exporting to PDF in the browser — the export renders the saved file server-side',
    )
  }
  const dir = await mkdtemp(join(tmpdir(), 'genoffice-web-pdf-'))
  try {
    const out = join(dir, 'export.pdf')
    await exportDocsHeadless(path, out, 'pdf')
    return await readFile(out)
  } finally {
    await rm(dir, { recursive: true, force: true })
  }
}
