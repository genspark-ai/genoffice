/// Withholding cell values from the model: the marks live in a custom package
/// part, `xl/gxRedactions.json`.
///
/// Why a part of our own, rather than a feature Excel already has:
///
/// - **`<protectedRanges>`** means "the user may not edit this", which is the
///   opposite promise, and it is inert without `<sheetProtection>` — a carrier
///   that only works when a feature it does not want is switched off.
/// - **Cell notes** would rewrite the whole comment set and the VML drawing,
///   destroying the reader's own notes.
/// - **`<definedNames>`** is the closest fit — a name plus a range — but Excel's
///   name grammar is `^[\p{L}_\\][\p{L}\p{N}_.\\]*$` (no spaces, no hyphens) and a
///   name may not repeat within a sheet. A label like "Client Phone" cannot be
///   written, and the label is the one string the model is allowed to see, so
///   mangling it is not an option.
///
/// A part of our own is inert (Excel ignores a relationship type it does not
/// know), keeps the label byte-exact, has no uniqueness constraint, and — the
/// decisive point — is never parsed by the cell pipeline. The worksheet writer
/// fully regenerates an edited cell and drops attributes it does not model
/// (see `patchCellStyleOnly`), so any per-cell carrier would be fragile; this
/// one lives beside the cells and no cell operation can touch it.
///
/// The save side is preservation by construction: `save_archive` copies every
/// entry it was not told to replace byte-for-byte (`archive.rs:209`), and
/// `readArchiveEntryText` reads any entry back, so no engine API was needed.

import { nextFreeRelationshipId } from './xlsx-sheets'

export class RedactionError extends Error {}
/** One withheld rectangle, in zero-based screen coordinates. */
export interface RedactionMark {
  readonly startRow: number
  readonly endRow: number
  readonly startColumn: number
  readonly endColumn: number
  readonly label: string
  /**
   * The fill the cell carried before the mark, so clearing can put it back.
   *
   * A mark is visible — it tints its cells, or nothing on the grid would say
   * which ones are withheld — and a tint written over the reader's own colour
   * would destroy it. Carrying the previous value in the mark makes clearing
   * restore rather than blank, which is the whole reason the mark is allowed
   * to touch formatting at all. `null` means the cell had no fill.
   *
   * Absent (older parts, and marks whose cells were never tinted) reads the
   * same as `null`: there is nothing to put back, so clearing removes the fill.
   */
  // `| undefined` explicitly: the repo compiles with exactOptionalPropertyTypes,
  // and the zod schema on the save path infers exactly that for a
  // `.nullish()` field.
  readonly previousFill?: string | null | undefined
}

/** Every mark on one sheet. Keyed by name, like `SheetProtectedRangesState`. */
export interface SheetRedactionState {
  readonly sheetName: string
  readonly marks: readonly RedactionMark[]
}

export const REDACTION_PART_PATH = 'xl/gxRedactions.json'
const REDACTION_CONTENT_TYPE = 'application/json'
const REDACTION_REL_TYPE = 'https://schemas.genspark.ai/genoffice/2026/relationships/redactions'
const WORKBOOK_RELS_PATH = 'xl/_rels/workbook.xml.rels'
const CONTENT_TYPES_PATH = '[Content_Types].xml'
const PART_VERSION = 1

/// The slice of `PackageEditor` this module needs. Declared structurally (as
/// `xlsx-notes.ts` does for its own) so importing the editor here would not
/// close a cycle with `xlsx-gateway.ts`, which calls into this file.
export interface RedactionPackage {
  has(path: string): Promise<boolean>
  readText(path: string): Promise<string>
  write(path: string, content: string): void
  add(path: string, content: string): void
  remove(path: string): void
}

function isNonNegativeInteger(value: unknown): value is number {
  return typeof value === 'number' && Number.isInteger(value) && value >= 0
}

/**
 * Parse the part, refusing anything this version does not understand.
 *
 * A mark that fails to parse must not degrade into "no marks": the part is the
 * only record of what the reader withheld, and quietly treating a damaged part
 * as an empty one would hand the model exactly the values the reader hid. So
 * every failure throws, and the save fails rather than losing the protection.
 */
export function parseRedactionPart(text: string): SheetRedactionState[] {
  let parsed: unknown
  try {
    parsed = JSON.parse(text)
  } catch (err) {
    throw new RedactionError(
      `${REDACTION_PART_PATH} is not valid JSON — refusing to save, because reading it as ` +
        `"no withheld cells" would reveal the values it records. (${err instanceof Error ? err.message : String(err)})`,
    )
  }
  if (typeof parsed !== 'object' || parsed === null || Array.isArray(parsed)) {
    throw new RedactionError(`${REDACTION_PART_PATH} must hold an object.`)
  }
  const root = parsed as Record<string, unknown>
  if (root.version !== PART_VERSION) {
    throw new RedactionError(
      `${REDACTION_PART_PATH} has version ${JSON.stringify(root.version)}, but this ` +
        `version of GenOffice only understands ${PART_VERSION}. Refusing to save.`,
    )
  }
  const sheets = root.sheets
  if (!Array.isArray(sheets)) {
    throw new RedactionError(`${REDACTION_PART_PATH} has no "sheets" array.`)
  }
  return sheets.map((entry, sheetIndex) => {
    if (typeof entry !== 'object' || entry === null) {
      throw new RedactionError(`${REDACTION_PART_PATH} sheets[${sheetIndex}] is not an object.`)
    }
    const sheet = entry as Record<string, unknown>
    if (typeof sheet.name !== 'string' || sheet.name === '') {
      throw new RedactionError(`${REDACTION_PART_PATH} sheets[${sheetIndex}] has no sheet name.`)
    }
    // Hoisted out of the map callback below: the property narrowing above does
    // not survive into a closure, where TypeScript assumes it may have changed.
    const sheetName = sheet.name
    const marks = sheet.marks
    if (!Array.isArray(marks)) {
      throw new RedactionError(`${REDACTION_PART_PATH} sheet "${sheetName}" has no "marks" array.`)
    }
    return {
      sheetName,
      marks: marks.map((mark, markIndex) => parseMark(sheetName, mark, markIndex)),
    }
  })
}

function parseMark(sheetName: string, mark: unknown, markIndex: number): RedactionMark {
  const where = `${REDACTION_PART_PATH} sheet "${sheetName}" marks[${markIndex}]`
  if (typeof mark !== 'object' || mark === null) {
    throw new RedactionError(`${where} is not an object.`)
  }
  const record = mark as Record<string, unknown>
  // Narrowed one field at a time: a loop over the key union would not carry the
  // narrowing past the iteration.
  const { startRow, endRow, startColumn, endColumn } = record
  if (!isNonNegativeInteger(startRow))
    throw new RedactionError(`${where} has an invalid "startRow".`)
  if (!isNonNegativeInteger(endRow)) throw new RedactionError(`${where} has an invalid "endRow".`)
  if (!isNonNegativeInteger(startColumn)) {
    throw new RedactionError(`${where} has an invalid "startColumn".`)
  }
  if (!isNonNegativeInteger(endColumn)) {
    throw new RedactionError(`${where} has an invalid "endColumn".`)
  }
  if (endRow < startRow || endColumn < startColumn) {
    throw new RedactionError(`${where} ends before it starts.`)
  }
  if (typeof record.label !== 'string' || record.label === '') {
    throw new RedactionError(`${where} has no label. The label is what the model reads.`)
  }
  // An unparseable previousFill is not fatal: the mark still withholds, and
  // clearing falls back to "no fill" rather than refusing to save. It is
  // dropped rather than kept as a wrong value, so a corrupt field can never
  // paint the wrong colour back onto the reader's cell.
  const previousFill = record.previousFill
  return {
    startRow,
    endRow,
    startColumn,
    endColumn,
    label: record.label,
    ...(typeof previousFill === 'string' || previousFill === null ? { previousFill } : {}),
  }
}

export function serializeRedactionPart(states: readonly SheetRedactionState[]): string {
  // Marks are sorted so that re-saving an unchanged workbook produces an
  // identical part: a byte-different part would show up as a touched entry on
  // every save and make "did anything change?" unanswerable.
  const sheets = [...states]
    .filter((state) => state.marks.length > 0)
    .sort((left, right) =>
      left.sheetName < right.sheetName ? -1 : left.sheetName > right.sheetName ? 1 : 0,
    )
    .map((state) => ({
      name: state.sheetName,
      marks: [...state.marks]
        .sort(
          (left, right) =>
            left.startRow - right.startRow ||
            left.startColumn - right.startColumn ||
            left.endRow - right.endRow ||
            left.endColumn - right.endColumn,
        )
        .map((mark) => ({
          startRow: mark.startRow,
          endRow: mark.endRow,
          startColumn: mark.startColumn,
          endColumn: mark.endColumn,
          label: mark.label,
          // omitted rather than written as null: a mark that had no fill to
          // restore must not make every other mark's bytes differ
          ...('previousFill' in mark ? { previousFill: mark.previousFill ?? null } : {}),
        })),
    }))
  return `${JSON.stringify({ version: PART_VERSION, sheets }, null, 2)}\n`
}

/**
 * Re-key the marks for the sheet edits this save also performs.
 *
 * A sheet renamed in the same save would otherwise orphan its marks: the part
 * is keyed by name, the new name carries no marks, and the reader's withheld
 * values would silently become visible to the model. A removed sheet's marks go
 * with it — the cells they described no longer exist.
 */
export function rekeyRedactionStates(
  states: readonly SheetRedactionState[],
  renames: readonly { readonly sheetName: string; readonly newName: string }[],
  removals: readonly string[],
): SheetRedactionState[] {
  const renamed = new Map(renames.map((rename) => [rename.sheetName, rename.newName]))
  const removed = new Set(removals)
  return states
    .filter((state) => !removed.has(state.sheetName))
    .map((state) => {
      const newName = renamed.get(state.sheetName)
      return newName === undefined ? state : { ...state, sheetName: newName }
    })
}

/**
 * Write the part, declaring it the way a package part must be declared.
 *
 * The three declarations are what make the part survive outside this app: the
 * part itself, a workbook relationship (an orphan part is legal but readers are
 * free to discard it), and a content-type override. `ensureDynamicArrayMetadata`
 * does the same for `xl/metadata.xml`; the recipe is unchanged.
 */
/**
 * Undo everything `applyRedactionPart` added: the part, the workbook
 * relationship, and the content-type override.
 *
 * All three, or the file is worse than before — an orphan part with a dangling
 * override is a difference the reader's file now carries forever, and Excel
 * complains about it on open.
 */
async function removeRedactionPart(
  pkg: RedactionPackage,
  touchedEntries: Set<string>,
): Promise<void> {
  if (!(await pkg.has(REDACTION_PART_PATH))) return
  pkg.remove(REDACTION_PART_PATH)
  touchedEntries.add(REDACTION_PART_PATH)

  const relationships = await pkg.readText(WORKBOOK_RELS_PATH)
  const withoutRel = relationships.replace(
    new RegExp(`<Relationship\\b[^>]*Type="${REDACTION_REL_TYPE}"[^>]*/?>`),
    '',
  )
  if (withoutRel !== relationships) {
    pkg.write(WORKBOOK_RELS_PATH, withoutRel)
    touchedEntries.add(WORKBOOK_RELS_PATH)
  }

  const contentTypes = await pkg.readText(CONTENT_TYPES_PATH)
  const withoutOverride = contentTypes.replace(
    new RegExp(`<Override\\b[^>]*PartName="/${REDACTION_PART_PATH}"[^>]*/>`),
    '',
  )
  if (withoutOverride !== contentTypes) {
    pkg.write(CONTENT_TYPES_PATH, withoutOverride)
    touchedEntries.add(CONTENT_TYPES_PATH)
  }
}

export async function applyRedactionPart(
  pkg: RedactionPackage,
  touchedEntries: Set<string>,
  states: readonly SheetRedactionState[],
): Promise<void> {
  // A workbook with nothing withheld must not grow the part: an empty part
  // would be a difference from the file the reader opened, forever after.
  //
  // And a workbook that *used* to withhold something must not keep it. An
  // early return here left the part, its relationship and its content-type
  // override in the package, so clearing the last mark did nothing: the next
  // open read the stale part and the span came back. Clearing is the one
  // operation that has to undo what setting did.
  if (states.every((state) => state.marks.length === 0)) {
    await removeRedactionPart(pkg, touchedEntries)
    return
  }

  const content = serializeRedactionPart(states)
  if (await pkg.has(REDACTION_PART_PATH)) {
    // Re-save of a workbook that already carries marks. `add` would land the
    // path in MutationPlan.added and trip assertManifestPreserved's
    // "should have created … but it already existed".
    pkg.write(REDACTION_PART_PATH, content)
  } else {
    pkg.add(REDACTION_PART_PATH, content)
  }
  touchedEntries.add(REDACTION_PART_PATH)

  const relationships = await pkg.readText(WORKBOOK_RELS_PATH)
  if (!relationships.includes(`Type="${REDACTION_REL_TYPE}"`)) {
    const relationship =
      `<Relationship Id="${nextFreeRelationshipId(relationships)}" ` +
      `Type="${REDACTION_REL_TYPE}" Target="gxRedactions.json"/>`
    pkg.write(
      WORKBOOK_RELS_PATH,
      relationships.replace('</Relationships>', `${relationship}</Relationships>`),
    )
    touchedEntries.add(WORKBOOK_RELS_PATH)
  }

  const contentTypes = await pkg.readText(CONTENT_TYPES_PATH)
  if (!contentTypes.includes(`PartName="/${REDACTION_PART_PATH}"`)) {
    const override = `<Override PartName="/${REDACTION_PART_PATH}" ContentType="${REDACTION_CONTENT_TYPE}"/>`
    pkg.write(CONTENT_TYPES_PATH, contentTypes.replace('</Types>', `${override}</Types>`))
    touchedEntries.add(CONTENT_TYPES_PATH)
  }
}
