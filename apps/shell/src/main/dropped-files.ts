/// Payload handling for OS file drops (installDropOpenBridge in preloads).
/// Kept in its own module so the decision logic is unit-testable without
/// booting the whole shell main process.
import { partitionDropPayload } from '@genoffice/electron-utils'

export interface DroppedFilesDeps {
  /** the normal File > Open pipeline; false = the path didn't open */
  openDocumentPath: (path: string) => boolean
  /** bring the shell window forward (the drop may land in a detached editor) */
  revealShellWindow: () => void
  /** the shared warning box; receives a localized message */
  showWarning: (message: string) => void
  /** localized template for one-or-more unsupported extensions */
  unsupportedMessage: (exts: string[]) => string
}

/**
 * Open each openable path (last drop wins activation), reveal the shell if
 * anything opened, and surface one combined warning for known-but-unsupported
 * formats instead of letting them disappear silently.
 */
export function handleDroppedFiles(raw: unknown, deps: DroppedFilesDeps): void {
  const { supported, unsupportedExts } = partitionDropPayload(raw)
  let opened = false
  for (const path of supported) opened = deps.openDocumentPath(path) || opened
  if (opened) deps.revealShellWindow()
  if (unsupportedExts.length > 0) deps.showWarning(deps.unsupportedMessage(unsupportedExts))
}
