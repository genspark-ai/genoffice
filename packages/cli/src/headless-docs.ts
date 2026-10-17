/**
 * The headless document path, for a host that needs to read a file and validate
 * operations without an editor — the shell's script kind opens a document here to
 * dry-run a feature's script against it.
 *
 * It is the same code the command line runs, so a script's plan is judged by the
 * engine that will one day apply it, and a dry run means the same thing in both
 * places.
 */
export {
  applyDocOps,
  closeDocument,
  describeDocument,
  openDocument,
  PREVIEW_CHARS,
  saveDocument,
} from './formats/docx'
export type { BlockSummary, DocOpsOutcome, OpenDocument } from './formats/docx'
export { readInput } from './fs'
export type { PathContext } from './fs'
