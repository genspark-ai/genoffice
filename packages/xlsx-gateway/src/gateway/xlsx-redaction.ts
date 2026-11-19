import { shiftCellArea, type StructuralOp } from './xlsx-structure'

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
/// The defined-name *name* really is that narrow, but the formula behind it is
/// free text, so `applyRedactionNameMirror` below writes the label there. That
/// mirror is not the record — the part is, and stays authoritative whenever it
/// is readable — it is what survives a reader that rebuilds the package and
/// drops what it does not model. See its own comment for what it does and does
/// not promise.
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
import { appendDefinedNames, escapeXmlText, unescapeXml } from './xlsx-defined-names'

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
  const sheets = statesInWriteOrder(states).map((state) => ({
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
  structuralOps: readonly {
    readonly sheetName: string
    readonly ops: readonly StructuralOp[]
  }[] = [],
): SheetRedactionState[] {
  const renamed = new Map(renames.map((rename) => [rename.sheetName, rename.newName]))
  const removed = new Set(removals)
  return states
    .filter((state) => !removed.has(state.sheetName))
    .map((state) => {
      const newName = renamed.get(state.sheetName)
      const renamedState = newName === undefined ? state : { ...state, sheetName: newName }
      const ops = structuralOps.find((entry) => entry.sheetName === renamedState.sheetName)?.ops
      if (!ops || ops.length === 0) return renamedState
      // The marks name cells the file is about to move: this same save
      // applies the ops to the grid XML, so a rectangle left where it was
      // would end up describing whatever shifted into its place. A mark whose
      // row or column the ops delete goes with them — the cell it withheld is
      // no longer there to withhold.
      const marks = renamedState.marks.flatMap((mark) => {
        const moved = shiftCellArea(mark, ops)
        return moved === null ? [] : [{ ...mark, ...moved }]
      })
      return { ...renamedState, marks }
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

/// The mirror: the same marks, written into `<definedNames>` as well.
///
/// ## Why a second copy
///
/// The part survives Excel and GenOffice. It does not survive everything. A
/// reader that rebuilds the package from its own model discards parts it does
/// not model, along with their relationships and content-type overrides, and
/// LibreOffice does exactly that: `soffice --convert-to xlsx` over a workbook
/// with marks leaves none of the three behind.
///
/// A defined name is the one carrier every spreadsheet program has to keep
/// working, because names are how print areas, data validation sources, and
/// navigation are expressed. So the marks go there too: when the part is
/// readable it wins and these are ignored, and when it is gone the marks are
/// rebuilt from them.
///
/// ## What that buys, and what it costs
///
/// It buys the coordinates, which are what withhold a cell. It does not buy
/// the sheet *name*: the sheet is recorded by position in workbook order, so a
/// workbook reordered outside GenOffice can point a mark at the wrong table.
/// That is the trade, and the direction it fails in is the safe one — a mark
/// that lands on the wrong table withholds something the reader never hid,
/// which they can undo, where a mark that is not recorded at all hands the
/// model the values they chose to withhold.
///
/// Two properties keep the mirror out of the reader's way. The names are
/// hidden, so they never appear in the Name Manager or the name box, and so
/// `applyDefinedNamesState` leaves them alone (it preserves hidden entries
/// verbatim). They are workbook-scoped rather than sheet-scoped, so deleting a
/// sheet cannot renumber every `localSheetId` behind it.
export const REDACTION_MIRROR_PREFIX = '_gxRedaction_'
const MIRROR_VERSION = 1
/**
 * Excel caps a defined name's value at 255 characters, and a name over the
 * limit is not a warning — the file is damaged. A label long enough to reach
 * it is dropped by `fitLabel`, never by skipping the entry.
 */
const MIRROR_FORMULA_LIMIT = 255
const MIRROR_TRUNCATION_MARK = '…'

/** One mirrored mark. `sheet` is a position in workbook sheet order, not a name. */
interface MirrorPayload {
  readonly v: number
  readonly sheet: number
  readonly from: readonly [number, number]
  readonly to: readonly [number, number]
  readonly label: string
}

/**
 * `"…"` — a formula string literal — is the one thing a defined name can hold
 * that every spreadsheet program parses without complaint, and its text is not
 * subject to the name grammar, which is what the label needed.
 *
 * Inside a formula string literal a quote is written doubled, so JSON's own
 * quotes are doubled on the way out and folded back on the way in. JSON never
 * emits two quotes in a row, so that folding is unambiguous.
 */
function mirrorFormula(payload: MirrorPayload): string {
  return `"${JSON.stringify(payload).replaceAll('"', '""')}"`
}

function unmirrorFormula(body: string): string {
  const text = body.trim()
  return text.length >= 2 && text.startsWith('"') && text.endsWith('"')
    ? text.slice(1, -1).replaceAll('""', '"')
    : text
}

/**
 * Trim a label until the entry fits Excel's 255-character cap.
 *
 * The entry is never dropped over a long label: the coordinates are what
 * withhold the cell, and an unrecorded mark is the leak this file exists to
 * prevent. The label only tells the model *what* was withheld, so it is the
 * part that gives way, and it says so with a trailing ellipsis rather than
 * passing itself off as the whole string.
 */
function labelFitter(payload: Omit<MirrorPayload, 'label'>): (label: string) => string {
  const overhead = Array.from(mirrorFormula({ ...payload, label: '' })).length
  const room = MIRROR_FORMULA_LIMIT - overhead
  return (label) => {
    const characters = Array.from(label)
    if (characters.length <= room) return label
    const kept = Math.max(0, room - 1)
    return characters.slice(0, kept).join('') + MIRROR_TRUNCATION_MARK
  }
}

/** The same ordering `serializeRedactionPart` uses, so the two agree mark for mark. */
function statesInWriteOrder(states: readonly SheetRedactionState[]): SheetRedactionState[] {
  return [...states]
    .filter((state) => state.marks.length > 0)
    .sort((left, right) =>
      left.sheetName < right.sheetName ? -1 : left.sheetName > right.sheetName ? 1 : 0,
    )
}

/**
 * Replace every mirrored entry in `workbookXml` with one per mark in `states`.
 *
 * Rewriting the whole set rather than adding to it is what makes clearing the
 * last mark work: there is nothing left behind describing a cell the reader
 * un-hid, in the part or in here. Entries for a sheet that is not in
 * `sheetOrder` are dropped — that sheet is gone from the file, which is the
 * one case where its marks should not be mirrored.
 */
export function applyRedactionNameMirror(
  workbookXml: string,
  states: readonly SheetRedactionState[],
  sheetOrder: readonly string[],
): string {
  const stripped = workbookXml.replace(MIRROR_ELEMENT, '')
  const entries: string[] = []
  for (const state of statesInWriteOrder(states)) {
    const sheet = sheetOrder.indexOf(state.sheetName)
    if (sheet < 0) continue
    for (const mark of state.marks) {
      const payload = {
        v: MIRROR_VERSION,
        sheet,
        from: [mark.startRow, mark.startColumn],
        to: [mark.endRow, mark.endColumn],
      } as const
      const formula = mirrorFormula({ ...payload, label: labelFitter(payload)(mark.label) })
      entries.push(
        `<definedName name="${REDACTION_MIRROR_PREFIX}${entries.length + 1}" hidden="1">` +
          `${escapeXmlText(formula)}</definedName>`,
      )
    }
  }
  if (entries.length === 0) return stripped
  return appendDefinedNames(stripped, entries.join(''))
}

const MIRROR_ELEMENT = new RegExp(
  `<definedName\\b[^>]*\\bname="${REDACTION_MIRROR_PREFIX}\\d+"[^>]*>` +
    `(?:[\\s\\S]*?</definedName>|)`,
  'g',
)

/**
 * Rebuild the marks a lost part recorded, from the mirrored defined names.
 *
 * Returns an empty list when the mirror is absent — that is a workbook that
 * never withheld anything, and the common case. Anything present but not
 * understood throws, for the reason `parseRedactionPart` throws: an unreadable
 * record of what was withheld must never be read as "nothing was withheld".
 */
export function parseRedactionNameMirror(
  workbookXml: string,
  sheetOrder: readonly string[],
): SheetRedactionState[] {
  const bySheet = new Map<string, RedactionMark[]>()
  for (const match of workbookXml.matchAll(MIRROR_ELEMENT)) {
    const element = match[0]
    const name = /\bname="([^"]*)"/.exec(element)?.[1] ?? '(unnamed)'
    const body = /<definedName\b[^>]*>([\s\S]*?)<\/definedName>/.exec(element)?.[1] ?? ''
    let payload: unknown
    try {
      payload = JSON.parse(unmirrorFormula(unescapeXml(body)))
    } catch (err) {
      throw new RedactionError(
        `Defined name "${name}" does not hold a redaction mirror this version can read — ` +
          `refusing to open the workbook as if nothing were withheld. (${
            err instanceof Error ? err.message : String(err)
          })`,
      )
    }
    const mark = mirrorMarkOf(name, payload, sheetOrder)
    const sheetName = mark[0]
    const marks = bySheet.get(sheetName)
    if (marks) marks.push(mark[1])
    else bySheet.set(sheetName, [mark[1]])
  }
  return [...bySheet].map(([sheetName, marks]) => ({ sheetName, marks }))
}

/** `[sheetName, mark]` — the sheet comes from the payload's position, validated. */
function mirrorMarkOf(
  name: string,
  payload: unknown,
  sheetOrder: readonly string[],
): [string, RedactionMark] {
  const where = `Defined name "${name}"`
  if (typeof payload !== 'object' || payload === null || Array.isArray(payload)) {
    throw new RedactionError(`${where} does not hold a redaction mirror.`)
  }
  const record = payload as Record<string, unknown>
  if (record.v !== MIRROR_VERSION) {
    throw new RedactionError(
      `${where} holds mirror version ${JSON.stringify(record.v)}, but this version of ` +
        `GenOffice only understands ${MIRROR_VERSION}. Refusing to open the workbook.`,
    )
  }
  if (!isNonNegativeInteger(record.sheet)) {
    throw new RedactionError(`${where} has no sheet position.`)
  }
  const sheetName = sheetOrder[record.sheet]
  if (sheetName === undefined) {
    throw new RedactionError(
      `${where} names sheet position ${record.sheet}, but the workbook has ` +
        `${sheetOrder.length} sheet(s) — the mirror and the file disagree about their shape.`,
    )
  }
  const from = mirrorCell(record.from, `${where} "from"`)
  const to = mirrorCell(record.to, `${where} "to"`)
  if (to[0] < from[0] || to[1] < from[1]) {
    throw new RedactionError(`${where} ends before it starts.`)
  }
  if (typeof record.label !== 'string' || record.label === '') {
    throw new RedactionError(`${where} has no label. The label is what the model reads.`)
  }
  return [
    sheetName,
    {
      startRow: from[0],
      startColumn: from[1],
      endRow: to[0],
      endColumn: to[1],
      label: record.label,
    },
  ]
}

function mirrorCell(value: unknown, where: string): [number, number] {
  if (
    !Array.isArray(value) ||
    value.length !== 2 ||
    !isNonNegativeInteger(value[0]) ||
    !isNonNegativeInteger(value[1])
  ) {
    throw new RedactionError(`${where} is not a cell corner.`)
  }
  return [value[0], value[1]]
}
