/**
 * Headless export mode: the renderer is hosted in a hidden window created by
 * `exportDocsHeadless` (main), with no user to open the document, watch
 * pagination settle or pick a save path. This module supplies only the
 * waiting — the export itself is the exact same `exportPdf` / `exportHtml`
 * the File menu runs, so the GUI and the CLI cannot drift.
 */
import {
  documentFontsSettled,
  headlessSleep,
  pollUntilReady,
  runHeadlessRendererExport,
  type HeadlessExportReport,
  type HeadlessWaitOptions,
} from '@genoffice/electron-utils/headless-export'

/** Pagination publishes its PageCut signature here on every remeasure (App.tsx). */
interface PageDebug {
  pageCutSig?: string
}

const pageCutSig = (): string => {
  const debug = (window as unknown as { __pageDebug?: PageDebug }).__pageDebug
  return typeof debug?.pageCutSig === 'string' ? debug.pageCutSig : ''
}

export interface DocumentReadiness {
  /** a document loaded from disk is mounted in the editor */
  opened: boolean
  /** the boot open landed on the untitled blank fallback: the input is unreadable */
  failed: boolean
}

const UNREADABLE = 'the input document did not open (corrupt, unsupported or password-protected)'

/**
 * Resolves once a document is open and its pagination has produced the same
 * page cuts (start, end, lead, header/footer variant, window facts) twice in
 * a row, so headless exports the layout the editor would have shown. A page
 * count alone can repeat while the pass chain still moves a cut.
 */
export async function waitForDocumentSettled(
  readiness: () => DocumentReadiness,
  options: HeadlessWaitOptions = {},
): Promise<void> {
  const { timeoutMs = 180_000, pollMs = 300 } = options
  const deadline = Date.now() + timeoutMs
  // A corrupt input is answered with a blank untitled document, not an error;
  // waiting out the full budget for it would only delay the same verdict.
  await pollUntilReady(
    () => {
      if (readiness().failed) throw new Error(UNREADABLE)
      return readiness().opened
    },
    UNREADABLE,
    { timeoutMs, pollMs },
  )
  // Async @font-face loads change line-break points, so pagination is only
  // trustworthy once the faces the document asked for have settled.
  await documentFontsSettled()
  let stable = 0
  let last = ''
  while (Date.now() <= deadline) {
    const sig = pageCutSig()
    if (sig !== '') {
      stable = sig === last ? stable + 1 : 0
      last = sig
      if (stable >= 2) return
    }
    await headlessSleep(pollMs)
  }
  // Out of budget: cuts that never stabilized are still exportable, a
  // document that produced none is not.
  if (last !== '') return
  throw new Error('pagination produced no pages')
}

/** Wait for the document, then run one of the GUI's own exports against a fixed path. */
export function runHeadlessDocumentExport(
  outPath: string,
  readiness: () => DocumentReadiness,
  exportFile: (outPath: string) => Promise<boolean>,
  options?: HeadlessWaitOptions,
): Promise<HeadlessExportReport> {
  return runHeadlessRendererExport(
    outPath,
    () => waitForDocumentSettled(readiness, options),
    exportFile,
  )
}
