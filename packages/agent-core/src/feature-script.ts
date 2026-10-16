import type { FeatureRecord } from './features-skill'

/**
 * The script kind's contract: the source a feature carries, and the shape of one
 * read-only run.
 *
 * A script is the body of a function that receives the API as its first argument
 * and the input object as its second. It runs in the host's sandbox, it can read
 * a document's blocks and *propose* operations, and — in this first slice — it
 * cannot write anything: the plan is validated and shown, and the file on disk is
 * untouched. The native assistant is the caller; the mod assistant only writes
 * the script.
 */

export const SCRIPT_LIMITS = {
  /** bytes of source one feature may carry */
  maxSourceBytes: 16 * 1024,
  /** operations one run may propose */
  maxOps: 5000,
  /** wall-clock budget for one run */
  defaultTimeoutMs: 5000,
  maxTimeoutMs: 30_000,
} as const

/**
 * One contract, one vocabulary per format: the document's extension decides which
 * API a script sees ($.docs, $.sheets, $.text, $.deck, $.pdf). A feature may
 * declare the format it was written for; at run time the host refuses a script
 * whose declared format does not match the document it was pointed at.
 */
export const SCRIPT_FORMATS = ['docs', 'sheets', 'text', 'deck', 'pdf'] as const

export type ScriptFormat = (typeof SCRIPT_FORMATS)[number]

export interface FeatureScript {
  /** the body of the function; it receives ($, input) */
  source: string
  /** the format this script targets, when its assistant declared one */
  format?: ScriptFormat
}

/** One read-only run, as the tool and the panel report it. */
export interface ScriptRunOutcome {
  /** proposed operations that the engine accepted */
  applied: number
  /** proposed operations, one line each: "3: findReplace: replaced 2 runs" */
  plan: string[]
  /** proposed operations the engine rejected */
  failures: string[]
  /** lines the script wrote with $.log */
  log: string[]
  /** milliseconds the sandbox was busy */
  ms: number
  /**
   * what a read-only script returned as its report ($.pdf today), JSON-encoded
   * and clipped by the host; absent when the script returned nothing reportable
   */
  result?: string
}

export interface ScriptValidationError {
  field: string
  message: string
}

export type ScriptValidation =
  { ok: true; script: FeatureScript } | { ok: false; errors: ScriptValidationError[] }

/** Module syntax is refused early with a readable message; the sandbox would only say "SyntaxError". */
const MODULE_SYNTAX = /^[ \t]*(export|import)[ \t]/m

/** Validate a script payload from the assistant or a mod file. */
export function validateScript(value: unknown): ScriptValidation {
  const errors: ScriptValidationError[] = []
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    return { ok: false, errors: [{ field: 'script', message: 'needs a script object' }] }
  }
  const input = value as Record<string, unknown>
  if (typeof input.source !== 'string' || input.source.trim().length === 0) {
    errors.push({ field: 'script.source', message: 'needs the code the script runs' })
    return { ok: false, errors }
  }
  const bytes = byteLength(input.source)
  if (bytes > SCRIPT_LIMITS.maxSourceBytes) {
    errors.push({
      field: 'script.source',
      message: 'is ' + bytes + ' bytes, more than the ' + SCRIPT_LIMITS.maxSourceBytes + ' allowed',
    })
  }
  if (MODULE_SYNTAX.test(input.source)) {
    errors.push({
      field: 'script.source',
      message:
        'must be the body of a function, not a module: drop import/export and write plain statements',
    })
  }
  let format: ScriptFormat | undefined
  if (input.format !== undefined) {
    if (
      typeof input.format !== 'string' ||
      !(SCRIPT_FORMATS as readonly string[]).includes(input.format)
    ) {
      errors.push({
        field: 'script.format',
        message:
          'must be one of ' + SCRIPT_FORMATS.join(', ') + ' (got ' + String(input.format) + ')',
      })
    } else {
      format = input.format as ScriptFormat
    }
  }
  if (errors.length) return { ok: false, errors }
  return { ok: true, script: format ? { source: input.source, format } : { source: input.source } }
}

function byteLength(text: string): number {
  let bytes = 0
  for (let i = 0; i < text.length; i++) {
    const code = text.charCodeAt(i)
    bytes += code < 0x80 ? 1 : code < 0x800 ? 2 : 3
  }
  return bytes
}

/** The API summary the assistant is told about; kept beside the contract so they cannot drift. */
export const SCRIPT_API_SUMMARY =
  'A script is the body of a function that receives ($, input). It may read the document, ' +
  'propose operations, write lines to the run log ($.log(...)), read $.input and watch ' +
  '$.budget.remainingMs. Which API it sees follows the document it runs against, so declare ' +
  'the format the script targets ({ source, format }) and use that namespace: ' +
  '"docs" for .docx ($.docs.blocks(range?, maxChars?), $.docs.text(), ' +
  '$.docs.propose([{ op: "findReplace", find, replace }, ...])); ' +
  '"sheets" for .xlsx ($.sheets.sheets(), $.sheets.usedRange(sheet), ' +
  '$.sheets.range("Sheet1!A2:F200") as a 2-D value grid, $.sheets.formulas(sheet?), ' +
  '$.sheets.formats(range), $.sheets.find(query), and $.sheets.propose(ops) with the ' +
  'workbook ops set_cell, set_formula, format_range, find_replace, sort_range, insert_rows, ' +
  'add_sheet, ... — every read is capped and refuses over the cap with a reason); ' +
  '"text" for .md and .html ($.text.get(), $.text.lines(); return the new full text from the ' +
  'function — one whole-document edit, never a partial patch); ' +
  '"deck" for .pptx ($.deck.slides() gives each slide and the box, text and id of every ' +
  'element, $.deck.propose(ops) with ' +
  'the same op vocabulary the slides app uses — setText, addElement, addPicture, ...); ' +
  '"pdf" is read-only ($.pdf.info(), $.pdf.pages(from, to, maxChars?)): return a report object ' +
  'and the run hands it back as the result. Every run is atomic: a refused operation rolls the ' +
  'whole batch back. No import, no export, no require, no network, no filesystem — the API is ' +
  'all it has.'

/** The scripts a registry holds, for the assistant to see what can be run. */
export function scriptsIn(features: readonly FeatureRecord[]): FeatureRecord[] {
  return features.filter((feature) => feature.script)
}
