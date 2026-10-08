import { printHtmlDocument, type PrintDialogWindow } from '@genoffice/electron-utils'

export type PrintWindow = PrintDialogWindow

/** Cancel comes back as a bare `{ ok: false }` so the renderer keeps its print dialog open. */
export async function printSlidesHtml(
  html: string,
  win: PrintWindow,
  platform: NodeJS.Platform = process.platform,
): Promise<{ ok: boolean; error?: string }> {
  const outcome = await printHtmlDocument({
    html,
    window: win,
    fileName: 'slides.html',
    dirPrefix: 'genoffice-slides-print-',
    platform,
  })
  if (!outcome.ok) return { ok: false, error: outcome.error }
  if ('canceled' in outcome) return { ok: false }
  return { ok: true }
}
