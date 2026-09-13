/**
 * Headless export mode: the renderer is hosted in a hidden window created by
 * `exportSlidesPdfHeadless` (main), with no user to open the deck, wait for
 * its pictures to decode or pick a save path. Only the waiting lives here —
 * the export itself is the File menu's own `exportPdf`, so the GUI and the
 * CLI cannot drift.
 */
import {
  documentFontsSettled,
  headlessSleep,
  pollUntilReady,
  runHeadlessRendererExport,
  type HeadlessExportReport,
  type HeadlessWaitOptions,
} from '@genoffice/electron-utils/headless-export'

export interface DeckReadiness {
  /** pages built from a deck that came from disk */
  slideCount: number
  /** pictures/fills still decoding; negative until the deck's image effect has run */
  pendingImages: number
  /** the boot open landed on the untitled blank deck: the input is unreadable */
  failed: boolean
}

const UNREADABLE = 'the input deck did not open (corrupt, unsupported or password-protected)'

/**
 * Resolves once the deck is open, every referenced image has decoded and the
 * document fonts have settled. Rasterizing earlier bakes missing pictures or
 * fallback faces into the exported PNGs.
 */
export async function waitForDeckSettled(
  readiness: () => DeckReadiness,
  options: HeadlessWaitOptions = {},
): Promise<void> {
  const { timeoutMs = 180_000, pollMs = 200 } = options
  const deadline = Date.now() + timeoutMs
  // A corrupt input is answered with a blank untitled deck, not an error;
  // waiting out the full budget for it would only delay the same verdict.
  await pollUntilReady(
    () => {
      const { slideCount, pendingImages, failed } = readiness()
      if (failed) throw new Error(UNREADABLE)
      return slideCount > 0 && pendingImages === 0
    },
    () => (readiness().slideCount === 0 ? UNREADABLE : 'images never finished decoding'),
    { timeoutMs, pollMs },
  )
  // The private Office FontFaces are registered asynchronously after the deck
  // settles (doc-fonts.ts); metafiles rasterized before that keep fallback faces.
  while (window.__genofficeDocFontsSynced === false && Date.now() <= deadline) {
    await headlessSleep(pollMs)
  }
  await documentFontsSettled()
  // Font registration can queue freshly rasterized metafiles: settle once more.
  const settleDeadline = Math.min(deadline, Date.now() + 15_000)
  while (readiness().pendingImages !== 0 && Date.now() <= settleDeadline) {
    await headlessSleep(pollMs)
  }
}

/** Wait for the deck, then run the GUI's own PDF export against a fixed path. */
export function runHeadlessPdfExport(
  outPath: string,
  readiness: () => DeckReadiness,
  exportPdf: (outPath: string) => Promise<boolean>,
  options?: HeadlessWaitOptions,
): Promise<HeadlessExportReport> {
  return runHeadlessRendererExport(outPath, () => waitForDeckSettled(readiness, options), exportPdf)
}
