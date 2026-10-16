/**
 * Reading the cells the reader withheld from the model, from the workbook
 * package part that records them (`xl/gxRedactions.json`), or — when that part
 * is not there — from the copy of the same marks written into the workbook's
 * defined names.
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
 *
 * ## Why there is a second place to read the marks from
 *
 * The part survives Excel and GenOffice, and not everything else. A reader that
 * rebuilds the package from its own model drops parts it does not model along
 * with their relationships and content-type overrides — `soffice
 * --convert-to xlsx` leaves none of the three. So every save also writes the
 * marks into defined names, which every spreadsheet program has to keep
 * working (`applyRedactionNameMirror` in the gateway).
 *
 * The part stays the record. The mirror is consulted only when the part cannot
 * be read at all, which is exactly the case where the alternative is
 * withholding the whole workbook or — worse, for a file that was merely
 * re-saved elsewhere — reading the values the reader chose to hide. Opening
 * such a workbook and saving it restores the part from the mirror, so the
 * fallback is a repair, not a permanent second-class mode.
 */
import { z } from 'zod'

import {
  REDACTION_PART_PATH,
  parseRedactionNameMirror,
  parseRedactionPart,
  type SheetRedactionState,
} from '@genoffice/xlsx-gateway/gateway/xlsx-redaction'
import { parseSheetElements } from '@genoffice/xlsx-gateway/gateway/xlsx-sheets'
import {
  readArchiveEntryText,
  type ArchiveClient,
} from '@genoffice/xlsx-gateway/gateway/xlsx-package-io'

/// Only the entry names are needed, so this stays structural rather than
/// re-validating every CRC and size the save path checks.
const manifestNamesSchema = z.object({ entries: z.array(z.object({ name: z.string() })) })

const WORKBOOK_PART_PATH = 'xl/workbook.xml'

/**
 * `source` says which copy answered, so the renderer can tell the reader that
 * the file came back from somewhere that drops the part. It never changes
 * which cells are withheld — both copies are written from the same state on
 * the same save — it only decides whether the workbook is fully intact.
 */
export type RedactionPartRead =
  | { readonly status: 'absent' }
  | {
      readonly status: 'ok'
      readonly states: readonly SheetRedactionState[]
      readonly source: 'part' | 'mirror'
    }
  | { readonly status: 'unreadable'; readonly error: string }

function messageOf(error: unknown): string {
  return error instanceof Error ? error.message : String(error)
}

/**
 * Read the redaction part out of an open session's snapshot, falling back to
 * the defined-name mirror when there is nothing to read there.
 *
 * Existence is checked through the archive manifest rather than by catching the
 * sidecar's "Workbook is missing …" failure: that string is a user-facing
 * message, not a contract, and the manifest is the same list the save path
 * diffs against. The snapshot is read rather than the live path so the marks
 * match the bytes this session's pending edits were made against.
 *
 * Any failure once the part is known to exist — unparseable JSON, an unknown
 * part version, a sidecar error mid-read — is `unreadable`. Not knowing what is
 * withheld is the one state that must never be read as "nothing is". It is the
 * state the mirror exists to rescue: the mirror holds the same marks, written
 * on the same save, so it gets a turn before the answer is given up on.
 */
export async function readWorkbookRedactionPart(
  client: ArchiveClient,
  snapshotPath: string,
): Promise<RedactionPartRead> {
  const manifest = manifestNamesSchema.parse(await client.archiveManifest(snapshotPath))
  if (manifest.entries.some((entry) => entry.name === REDACTION_PART_PATH)) {
    try {
      const text = await readArchiveEntryText(client, snapshotPath, REDACTION_PART_PATH)
      return { status: 'ok', states: parseRedactionPart(text), source: 'part' }
    } catch (error) {
      // The mirror is a rescue, not a second diagnosis: only an answer it can
      // actually give replaces the part's. When it cannot, the part's error is
      // the one worth showing — it names what is actually wrong with the file.
      const fromMirror = await readRedactionNameMirror(client, snapshotPath)
      if (fromMirror.status === 'ok') return fromMirror
      return { status: 'unreadable', error: messageOf(error) }
    }
  }
  return readRedactionNameMirror(client, snapshotPath)
}

/**
 * Rebuild the marks from the workbook's defined names.
 *
 * A workbook that never withheld anything has no mirrored entries, which is
 * the `absent` this returns — the same answer the missing part gives, and for
 * the same reason. A mirrored entry this version cannot read is `unreadable`
 * for the reason the part is: a record of what was withheld that fails to
 * parse must not be read as "nothing was withheld".
 */
async function readRedactionNameMirror(
  client: ArchiveClient,
  snapshotPath: string,
): Promise<RedactionPartRead> {
  try {
    const workbookXml = await readArchiveEntryText(client, snapshotPath, WORKBOOK_PART_PATH)
    // The mirror records each sheet by its position, so the positions are read
    // back out of the same file — the one the marks have to be resolved against.
    const states = parseRedactionNameMirror(
      workbookXml,
      parseSheetElements(workbookXml).map((sheet) => sheet.name),
    )
    return states.length === 0 ? { status: 'absent' } : { status: 'ok', states, source: 'mirror' }
  } catch (error) {
    return { status: 'unreadable', error: messageOf(error) }
  }
}
