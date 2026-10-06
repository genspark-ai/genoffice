/**
 * Reading the cells the reader withheld from the model, from the workbook
 * package part that records them (`xl/gxRedactions.json`).
 *
 * ## Why this is a three-state answer
 *
 * The part is optional — a workbook that never withheld anything has none, and
 * that is the common case — but its absence and a part this version cannot
 * parse mean opposite things:
 *
 * - **absent** — nothing is withheld. Reads go through unchanged.
 * - **unreadable** — the part exists and says something this version does not
 *   understand. `[]` here would be a lie in the worst direction: the model
 *   would receive exactly the values the reader hid, and nothing on screen
 *   would say so.
 *
 * So the failure is carried to the renderer as its own state instead of being
 * smoothed into an empty list, and the renderer withholds rather than reveals
 * (see `renderer/ai/redact-load.ts`).
 */
import { z } from 'zod'

import {
  REDACTION_PART_PATH,
  parseRedactionPart,
  type SheetRedactionState,
} from '@genoffice/xlsx-gateway/gateway/xlsx-redaction'
import {
  readArchiveEntryText,
  type ArchiveClient,
} from '@genoffice/xlsx-gateway/gateway/xlsx-package-io'

/// Only the entry names are needed, so this stays structural rather than
/// re-validating every CRC and size the save path checks.
const manifestNamesSchema = z.object({ entries: z.array(z.object({ name: z.string() })) })

export type RedactionPartRead =
  | { readonly status: 'absent' }
  | { readonly status: 'ok'; readonly states: readonly SheetRedactionState[] }
  | { readonly status: 'unreadable'; readonly error: string }

/**
 * Read the redaction part out of an open session's snapshot.
 *
 * Existence is checked through the archive manifest rather than by catching the
 * sidecar's "Workbook is missing …" failure: that string is a user-facing
 * message, not a contract, and the manifest is the same list the save path
 * diffs against. The snapshot is read rather than the live path so the marks
 * match the bytes this session's pending edits were made against.
 *
 * Any failure once the part is known to exist — unparseable JSON, an unknown
 * part version, a sidecar error mid-read — is `unreadable`. Not knowing what is
 * withheld is the one state that must never be read as "nothing is".
 */
export async function readWorkbookRedactionPart(
  client: ArchiveClient,
  snapshotPath: string,
): Promise<RedactionPartRead> {
  const manifest = manifestNamesSchema.parse(await client.archiveManifest(snapshotPath))
  if (!manifest.entries.some((entry) => entry.name === REDACTION_PART_PATH)) {
    return { status: 'absent' }
  }
  try {
    const text = await readArchiveEntryText(client, snapshotPath, REDACTION_PART_PATH)
    return { status: 'ok', states: parseRedactionPart(text) }
  } catch (error) {
    return {
      status: 'unreadable',
      error: error instanceof Error ? error.message : String(error),
    }
  }
}
