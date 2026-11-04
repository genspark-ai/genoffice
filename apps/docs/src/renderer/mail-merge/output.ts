import { mergeOutputFileName, type MergeRecord } from '@genoffice/docx-engine'
import type { MergeOutputMode } from '../../shared/ipc'
import { buildDocBytes, type FileActionContext } from '../file-actions'
import type { PmNode } from '../editor/convert'
import {
  countNextFields,
  hasNextField,
  mergeAllRecords,
  mergeNextPages,
  substituteMergeFields,
  substituteWithCursor,
} from './template'
import {
  cellContentOf,
  labelSheetDoc,
  layoutSection,
  type EnvelopePresetId,
  type MergePageLayout,
} from './layout'

/**
 * The output half of the mail merge: the documents the merge writes.
 *
 * The letters are the template with one record's values in it, serialized the
 * way the document itself is serialized — so a merged letter keeps the
 * template's styles, sections, headers and pictures, and carries no field
 * code any more. genoffice#2097.
 */

/** a slice of the recipient list the merge walks, by record index (inclusive) */
export interface MergeRange {
  /** first record index (0-based) */
  start: number
  /** last record index (0-based, inclusive) */
  end: number
}

/** what the user asked the merge to write */
export interface MailMergeRequest {
  /** single / perRecord / labels / envelopes (labels/envelopes require layout) */
  mode: MergeOutputMode
  /** the records to merge */
  records: readonly MergeRecord[]
  /** the column names of the data source */
  fields: readonly string[]
  /** the field the file names come from; empty takes the first field */
  namePattern: string
  /** the slice of records the merge walks; defaults to the whole list */
  range?: MergeRange
  /** the paper + label grid for labels, or envelope size for envelopes */
  layout?: MergePageLayout
  /** which label stock the user picked (Avery id / 'custom') */
  labelStock?: 'avery5160' | 'custom'
  /** custom label grid parameters (only used when labelStock === 'custom') */
  customGrid?: { paper: 'letter' | 'a4'; columns: number; rows: number }
  /** envelope size preset (only used when mode === 'envelopes') */
  envelopePreset?: EnvelopePresetId
}

/** one merged document, ready to be written */
export interface MailMergeFile {
  /** the file name, without a folder */
  name: string
  /** the .docx bytes */
  bytes: Uint8Array
}

/** the answer to a merge: what was written, or what went wrong */
export interface MailMergeResponse {
  ok: boolean
  /** how many documents were written */
  count?: number
  /** the file that was written (one document for all the records) */
  path?: string
  /** the folder the documents went into (one document per record) */
  folder?: string
  /** the user closed the dialog without choosing */
  canceled?: boolean
  error?: string
}

/**
 * The documents a merge writes.
 *
 * The records are first trimmed to the requested range. Then there are two
 * families:
 *
 *  - **plain** (mode single / perRecord): letters, or envelopes, one per record.
 *    The NEXT-driven variants follow at the bottom.
 *  - **labels / envelopes**: multi-up layouts on one page. Labels rides one
 *    continuous table — a cell per record, rows flowing onto as many pages as
 *    the records take — so the table's own pagination does the page breaks.
 *    Envelopes build one letter per record on the envelope paper, one file each.
 *
 * genoffice#2097, genoffice#2374.
 */
export async function buildMailMergeFiles(
  ctx: FileActionContext,
  request: MailMergeRequest,
): Promise<MailMergeFile[]> {
  const editor = ctx.editor
  if (!editor || request.records.length === 0) return []
  const template = editor.getJSON() as PmNode
  const start = request.range?.start ?? 0
  const end = request.range?.end ?? request.records.length - 1
  const subset = request.records.slice(
    Math.max(start, 0),
    Math.min(end, request.records.length - 1) + 1,
  )
  if (subset.length === 0) return []

  // --- labels / envelopes: the layout-driven branches -------------------
  if (request.mode === 'labels' || request.mode === 'envelopes') {
    const layout = request.layout
    if (!layout) return []
    const layoutCtx: FileActionContext = {
      ...ctx,
      section: layoutSection(layout, ctx.section),
      sectionDirty: true,
    }
    if (request.mode === 'labels') {
      const cells = subset.map((record) =>
        cellContentOf(substituteMergeFields(template, record)),
      )
      const bytes = await buildDocBytes(layoutCtx, labelSheetDoc(cells, layout))
      if (!bytes) return []
      return [{ name: mergedDocumentName(ctx), bytes }]
    }
    // envelopes: one file per record on the envelope's paper
    const files: MailMergeFile[] = []
    for (const record of subset) {
      const bytes = await buildDocBytes(layoutCtx, substituteMergeFields(template, record))
      if (!bytes) continue
      files.push({
        name: `${mergeOutputFileName(request.namePattern, record, request.fields)}.docx`,
        bytes,
      })
    }
    return files
  }

  // --- plain letters: split on whether the template carries a NEXT field -
  if (!hasNextField(template)) {
    if (request.mode === 'perRecord') {
      const files: MailMergeFile[] = []
      for (const record of subset) {
        const bytes = await buildDocBytes(ctx, substituteMergeFields(template, record))
        if (!bytes) continue
        files.push({
          name: `${mergeOutputFileName(request.namePattern, record, request.fields)}.docx`,
          bytes,
        })
      }
      return files
    }
    const bytes = await buildDocBytes(ctx, mergeAllRecords(template, subset))
    if (!bytes) return []
    return [{ name: mergedDocumentName(ctx), bytes }]
  }

  // plain with NEXT: cursor walks one page per slotsPerPage (NEXT count + 1)
  const nextCount = countNextFields(template)
  const slotsPerPage = nextCount + 1

  if (request.mode === 'perRecord') {
    const files: MailMergeFile[] = []
    for (let from = 0; from < subset.length; from += slotsPerPage) {
      const endIdx = from + slotsPerPage - 1
      const bytes = await buildDocBytes(
        ctx,
        substituteWithCursor(template, subset, from, endIdx),
      )
      if (!bytes) continue
      const name = `${mergeOutputFileName(
        request.namePattern,
        subset[from] ?? {},
        request.fields,
      )}.docx`
      files.push({ name, bytes })
    }
    return files
  }
  const bytes = await buildDocBytes(
    ctx,
    mergeNextPages(template, subset, slotsPerPage),
  )
  if (!bytes) return []
  return [{ name: mergedDocumentName(ctx), bytes }]
}

/** the name of the document that holds every letter */
function mergedDocumentName(ctx: FileActionContext): string {
  const base = (ctx.doc?.fileName ?? 'document').replace(/\.docx$/i, '')
  return `${mergeOutputFileName(base, {}, [])}-merged.docx`
}
