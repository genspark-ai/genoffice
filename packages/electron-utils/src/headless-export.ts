/**
 * Argv parsing and the stdout envelope for the headless export entry:
 *
 *   <app binary> --headless-export <input-file> --to <format> --out <path> [--json]
 *
 * Pure logic only — the Electron-side host that actually renders the file
 * lives in the shell main process (apps/shell/src/main/headless-export.ts).
 * Keeping the grammar and the exit codes here makes them unit-testable and
 * gives the external `genoffice` CLI one place to read the contract from.
 *
 * The renderer-side waiting helpers at the bottom live here too (imported as
 * `@genoffice/electron-utils/headless-export`, never through the package root,
 * which pulls in node: builtins) so every editor module reports the same way.
 */

/** Editor module that owns a given input extension. */
export type HeadlessExportModule = 'docs' | 'sheets' | 'slides' | 'markdown' | 'html'

/** Conversion targets the headless entry accepts; see HEADLESS_TARGETS for which module renders which. */
export type HeadlessExportFormat = 'pdf' | 'docx' | 'html'

/** Targets each editor module can render without a dialog. */
export const HEADLESS_TARGETS: Record<HeadlessExportModule, readonly HeadlessExportFormat[]> = {
  docs: ['pdf', 'html'],
  sheets: ['pdf'],
  slides: ['pdf'],
  markdown: ['pdf'],
  html: ['pdf', 'docx'],
}

const ALL_TARGETS: readonly HeadlessExportFormat[] = ['pdf', 'docx', 'html']

/** What a hidden renderer window is asked to produce. */
export interface HeadlessExportTarget {
  outPath: string
  format: HeadlessExportFormat
}

export const HEADLESS_EXPORT_FLAG = '--headless-export'

/**
 * Process exit codes, matching the genoffice envelope convention.
 * 1 bad args / 2 input file error / 3 conversion failure.
 */
export const HEADLESS_EXIT = {
  success: 0,
  badArgs: 1,
  inputError: 2,
  conversionFailure: 3,
} as const

export type HeadlessExitCode = (typeof HEADLESS_EXIT)[keyof typeof HEADLESS_EXIT]

export interface HeadlessExportRequest {
  /** absolute or cwd-relative path of the document to convert */
  input: string
  targetFormat: HeadlessExportFormat
  /** where the exported file is written */
  outPath: string
  /** print the machine-readable one-line envelope instead of a human sentence */
  json: boolean
}

export type HeadlessArgvParse =
  | { kind: 'none' }
  | { kind: 'ok'; request: HeadlessExportRequest }
  | { kind: 'error'; json: boolean; message: string }

/** Result of an export attempt, rendered by `formatHeadlessEnvelope`. */
export type HeadlessExportOutcome =
  | { ok: true; input: string; outPath: string }
  | { ok: false; code: Exclude<HeadlessExitCode, 0>; message: string }

const MODULE_BY_EXTENSION: ReadonlyArray<readonly [RegExp, HeadlessExportModule]> = [
  [/\.docx$/i, 'docs'],
  [/\.(xlsx|xlsm|xls|csv|tsv)$/i, 'sheets'],
  [/\.pptx$/i, 'slides'],
  [/\.(md|markdown)$/i, 'markdown'],
  [/\.(html|htm)$/i, 'html'],
]

/** Which editor module can render this input, or null when the extension is unsupported. */
export function headlessModuleFor(inputPath: string): HeadlessExportModule | null {
  for (const [pattern, module] of MODULE_BY_EXTENSION) if (pattern.test(inputPath)) return module
  return null
}

/** Extensions the headless entry accepts, for error messages. */
export const HEADLESS_SUPPORTED_EXTENSIONS =
  '.docx, .xlsx, .xlsm, .xls, .csv, .pptx, .md, .markdown, .html, .htm'

/** `--flag value` and `--flag=value` both read the same. */
function readOption(argv: readonly string[], index: number, flag: string): string | null {
  const arg = argv[index] ?? ''
  if (arg.startsWith(`${flag}=`)) return arg.slice(flag.length + 1)
  const next = argv[index + 1]
  // a following switch is never this flag's value (`--out --json` is a missing value)
  return next === undefined || next.startsWith('--') ? null : next
}

function consumesNext(argv: readonly string[], index: number, flag: string): boolean {
  return !(argv[index] ?? '').startsWith(`${flag}=`)
}

/**
 * Reads the headless-export switches out of an Electron `process.argv`.
 * Returns `{ kind: 'none' }` when the flag is absent, so the normal GUI
 * startup path stays untouched.
 */
export function parseHeadlessExportArgv(argv: readonly string[]): HeadlessArgvParse {
  const flagAt = argv.findIndex(
    (arg) => arg === HEADLESS_EXPORT_FLAG || arg.startsWith(`${HEADLESS_EXPORT_FLAG}=`),
  )
  if (flagAt === -1) return { kind: 'none' }
  const json = argv.includes('--json')
  const fail = (message: string): HeadlessArgvParse => ({ kind: 'error', json, message })

  let input: string | null = null
  let to: string | null = null
  let outPath: string | null = null

  for (let i = 0; i < argv.length; i++) {
    const arg = argv[i] ?? ''
    if (arg === HEADLESS_EXPORT_FLAG || arg.startsWith(`${HEADLESS_EXPORT_FLAG}=`)) {
      input = readOption(argv, i, HEADLESS_EXPORT_FLAG)
      if (input === null) return fail(`${HEADLESS_EXPORT_FLAG} needs an input file`)
      if (consumesNext(argv, i, HEADLESS_EXPORT_FLAG)) i++
      continue
    }
    if (arg === '--to' || arg.startsWith('--to=')) {
      to = readOption(argv, i, '--to')
      if (to === null) return fail('--to needs a target format')
      if (consumesNext(argv, i, '--to')) i++
      continue
    }
    if (arg === '--out' || arg.startsWith('--out=')) {
      outPath = readOption(argv, i, '--out')
      if (outPath === null) return fail('--out needs an output path')
      if (consumesNext(argv, i, '--out')) i++
      continue
    }
  }

  if (!input) return fail(`${HEADLESS_EXPORT_FLAG} needs an input file`)
  if (!to) return fail(`--to is required (${ALL_TARGETS.join(', ')})`)
  const targetFormat = ALL_TARGETS.find((f) => f === to.toLowerCase())
  if (!targetFormat) {
    return fail(`unsupported target format "${to}" (${ALL_TARGETS.join(', ')})`)
  }
  if (!outPath) return fail('--out is required')

  return { kind: 'ok', request: { input, targetFormat, outPath, json } }
}

/** One-sentence description of an outcome, shared by the plain and JSON forms. */
export function headlessSummary(outcome: HeadlessExportOutcome): string {
  return outcome.ok
    ? `Exported ${outcome.input} to ${outcome.outPath}`
    : `Export failed: ${outcome.message}`
}

/**
 * The single stdout line. With `--json` it is exactly one JSON object;
 * otherwise a human-readable sentence. Newlines inside the message would
 * break the one-line contract, so they are folded to spaces.
 */
export function formatHeadlessEnvelope(outcome: HeadlessExportOutcome, json: boolean): string {
  const summary = headlessSummary(outcome).replace(/\s*[\r\n]+\s*/g, ' ')
  if (!json) return summary
  return JSON.stringify(
    outcome.ok
      ? { status: 'ok', summary, output_path: outcome.outPath }
      : { status: 'error', summary, error: outcome.message.replace(/\s*[\r\n]+\s*/g, ' ') },
  )
}

/** Exit code for an outcome. */
export function headlessExitCode(outcome: HeadlessExportOutcome): HeadlessExitCode {
  return outcome.ok ? HEADLESS_EXIT.success : outcome.code
}

// ---- renderer side ----

/** What a hidden renderer sends back when its export settles. */
export interface HeadlessExportReport {
  ok: boolean
  error?: string
}

export interface HeadlessWaitOptions {
  /** overall budget before the wait gives up */
  timeoutMs?: number
  pollMs?: number
}

export const headlessSleep = (ms: number): Promise<void> =>
  new Promise((resolve) => setTimeout(resolve, ms))

/**
 * Polls `ready` until it returns true. There is no user to notice a document
 * that never opens, so a stalled export must fail rather than hang forever;
 * `stallMessage` says what was still missing (a function is evaluated at
 * give-up time, so it can describe the state the wait actually died in).
 */
export async function pollUntilReady(
  ready: () => boolean,
  stallMessage: string | (() => string),
  { timeoutMs = 180_000, pollMs = 250 }: HeadlessWaitOptions = {},
): Promise<void> {
  const deadline = Date.now() + timeoutMs
  while (!ready()) {
    if (Date.now() > deadline) {
      throw new Error(typeof stallMessage === 'function' ? stallMessage() : stallMessage)
    }
    await headlessSleep(pollMs)
  }
}

/**
 * Resolves when the document's @font-face loads have settled (they change
 * line-break points, so measuring before that is worthless), giving up after
 * `timeoutMs` rather than stranding the export. A DOM without FontFaceSet
 * (jsdom in unit tests) resolves immediately.
 */
export function documentFontsSettled(timeoutMs = 30_000): Promise<unknown> {
  if (typeof document === 'undefined' || !('fonts' in document)) return Promise.resolve()
  return Promise.race([document.fonts.ready, headlessSleep(timeoutMs)])
}

/**
 * The renderer half of a headless export: wait for the document, run the very
 * same export the File menu runs, and turn any failure into a report. Never
 * throws — the main process owns the process exit code.
 */
export async function runHeadlessRendererExport(
  outPath: string,
  waitUntilReady: () => Promise<void>,
  exportFile: (outPath: string) => Promise<boolean>,
): Promise<HeadlessExportReport> {
  try {
    await waitUntilReady()
    const ok = await exportFile(outPath)
    return ok ? { ok: true } : { ok: false, error: 'export did not produce a file' }
  } catch (err) {
    return { ok: false, error: err instanceof Error ? err.message : String(err) }
  }
}
