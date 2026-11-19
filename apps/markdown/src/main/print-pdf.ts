import { mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { BrowserWindow } from 'electron'

/** Shared pagination for PDF and PNG export. Document scripts remain disabled. */
export async function printMarkdownPdf(html: string): Promise<Buffer> {
  const workDir = await mkdtemp(join(tmpdir(), 'genoffice-md-print-'))
  let printWin: BrowserWindow | undefined
  try {
    printWin = new BrowserWindow({
      show: false,
      webPreferences: { sandbox: true, javascript: false },
    })
    const htmlPath = join(workDir, 'print.html')
    await writeFile(htmlPath, html, 'utf8')
    await printWin.loadFile(htmlPath)
    return await printWin.webContents.printToPDF({
      pageSize: 'A4',
      printBackground: true,
      margins: { top: 0.6, bottom: 0.6, left: 0.6, right: 0.6 },
    })
  } finally {
    printWin?.destroy()
    await rm(workDir, { recursive: true, force: true })
  }
}
