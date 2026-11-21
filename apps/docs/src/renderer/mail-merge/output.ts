import { joinLabelSheets, labelTableIndex } from './label-sheet'
import { mergeOutputFileName, type MergeRecord } from '@genoffice/docx-engine'
import type { MergeOutputMode } from '../../shared/ipc'
import { buildDocBytes, type FileActionContext } from '../file-actions'
import type { PmNode } from '../editor/convert'
import {
  joinMergedDocuments,
  mergeDocuments,
  type MergePrompt,
  type MergeRunOptions,
} from './template'
import { cellContentOf, labelSheetDoc, layoutSection, type MergePageLayout } from './layout'

/**
 * The output half of the mail merge: the documents the merge writes.
 *
 * The letters are the template with one record's values in it, serialized the
 * way the document itself is serialized — so a merged letter keeps the
 * template's styles, sections, headers and pictures, and carries no field
 * code any more. genoffice#2097.
 */

/** what the user asked the merge to write */
export interface MailMergeRequest {
  /** one document for all the records, or one document per record */
  mode: MergeOutputMode
  /** the records to merge */
  records: readonly MergeRecord[]
  /** the column names of the data source */
  fields: readonly string[]
  /** the field the file names come from; empty takes the first field */
  namePattern: string
  /** the paper the labels or envelopes print on; the label and envelope modes need it */
  layout?: MergePageLayout
  /** the place in the list of the first record, counted from 1 (Merge Record # counts from it) */
  firstRecord?: number
}

/** puts a question of the merge to the user; the answer, or the default when the user gives none */
export type MergeAsk = (prompt: MergePrompt) => Promise<string>

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
 * One document per record is the letter case: every record gets its own
 * letter, named after the record. One document for all the records is the
 * directory case: the letters follow each other, each on its own page. The
 * label case is one label-sheet document with the records in a grid; the
 * envelope case is one document per record on the envelope's paper.
 */
export async function buildMailMergeFiles(
  ctx: FileActionContext,
  request: MailMergeRequest,
  ask?: MergeAsk,
): Promise<MailMergeFile[]> {
  const editor = ctx.editor
  if (!editor || request.records.length === 0) return []
  const run: MergeRunOptions = { firstRecord: request.firstRecord ?? 1 }
  // what the merge will ask and write is planned on the document as it stands; each
  // file is then made from the live document the save hands over
  const snapshot = editor.getJSON() as PmNode
  let planned = mergeDocuments(snapshot, request.records, run)
  if (planned.prompts.length > 0 && ask) {
    const answers = new Map<string, string>()
    for (const prompt of planned.prompts) answers.set(prompt.key, await ask(prompt))
    run.answers = answers
    planned = mergeDocuments(snapshot, request.records, run)
  }
  if (planned.documents.length === 0) return []
  const mergedOf = (live: PmNode) =>
    mergeDocuments(live, request.records, run).documents.map((document) => document.doc)
  const paper =
    (request.mode === 'labels' || request.mode === 'envelopes') && request.layout
      ? layoutSection(request.layout, ctx.section)
      : undefined
  if (request.mode === 'labels' || request.mode === 'envelopes') {
    // the output prints on the layout's paper, not the template's
    if (!request.layout || !paper) return []
    const sheetAt = request.mode === 'labels' ? labelTableIndex(snapshot) : -1
    if (sheetAt >= 0) {
      // the document is the label sheet itself: it merges as it stands, on its own paper
      const bytes = await buildDocBytes(
        ctx,
        (live) => joinLabelSheets(mergedOf(live), sheetAt) ?? live,
      )
      if (!bytes) return []
      return [{ name: mergedDocumentName(ctx), bytes }]
    }
    if (request.mode === 'labels') {
      const { layout } = request
      const sheet = (live: PmNode) => labelSheetDoc(mergedOf(live).map(cellContentOf), layout)
      const bytes = await buildDocBytes(ctx, sheet, paper)
      if (!bytes) return []
      return [{ name: mergedDocumentName(ctx), bytes }]
    }
  }
  if (request.mode === 'perRecord' || request.mode === 'envelopes') {
    const files: MailMergeFile[] = []
    for (const [index, document] of planned.documents.entries()) {
      const bytes = await buildDocBytes(ctx, (live) => mergedOf(live)[index] ?? document.doc, paper)
      if (!bytes) continue
      files.push({
        name: `${mergeOutputFileName(request.namePattern, document.record, request.fields)}.docx`,
        bytes,
      })
    }
    return files
  }
  const bytes = await buildDocBytes(ctx, (live) => joinMergedDocuments(live, mergedOf(live)))
  if (!bytes) return []
  return [{ name: mergedDocumentName(ctx), bytes }]
}

/** the name of the document that holds every letter */
function mergedDocumentName(ctx: FileActionContext): string {
  const base = (ctx.doc?.fileName ?? 'document').replace(/\.docx$/i, '')
  return `${mergeOutputFileName(base, {}, [])}-merged.docx`
}
