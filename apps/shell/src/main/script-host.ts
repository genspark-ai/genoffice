import { renameSync, writeFileSync } from 'node:fs'
import { basename, dirname } from 'node:path'
import { createContext, runInContext } from 'node:vm'
import {
  SCRIPT_FORMATS,
  SCRIPT_LIMITS,
  type ScriptFormat,
  type ScriptRunOutcome,
} from '@genoffice/agent-core'
import type { ScriptApplyResult } from '../shared/feature-channels'
import {
  PREVIEW_CHARS,
  applyDocOps,
  closeDocument,
  describeDocument,
  openDocument,
  readInput,
  saveDocument,
} from '@genoffice/cli/headless-docs'
import {
  SandboxOutcome,
  ScriptRunError,
  SHEETS_FIND_CAP,
  SHEETS_FORMAT_CAP,
  SHEETS_READ_CELL_CAP,
  scriptFormatFor,
  type DocsSnapshot,
  type ScriptSnapshot,
} from './script-shared'
import { pdfHost } from './script-pdf'
import { deckHost } from './script-deck'
import { sheetsHost } from './script-sheets'
import { textHost } from './script-text'

/**
 * The script host: one script, one read-only snapshot, one plan.
 *
 * It lives in the shell main process because that is where a document can be
 * opened headlessly and where a registry entry can be run on demand. The
 * sandbox below is the same one the command line used (see
 * docs/mods-kind-script.md): the API is built inside the context from plain
 * JSON, so no host object crosses the boundary.
 *
 * One contract, one vocabulary per format: the document's extension decides the
 * snapshot and the namespace the script sees ($.docs, $.sheets, $.text, $.deck,
 * $.pdf), and each format module judges what the script proposed with the same
 * engine its own command line uses (script-shared.ts holds the union of the
 * snapshots and the per-format read caps).
 *
 * The API the script sees is built *inside* the sandbox from plain JSON, so no
 * host object crosses the boundary and there is nothing to reach through a host
 * function's constructor. Code generation is off, Promise is removed and no
 * timer exists, so the whole run is synchronous and the vm timeout bounds it.
 *
 * This is a capability boundary, not a hardened security boundary: the script is
 * a local file the user or their assistant wrote. Running untrusted scripts in a
 * separate process is deliberately left to a later change.
 */

export const SCRIPT_DEFAULT_TIMEOUT_MS = SCRIPT_LIMITS.defaultTimeoutMs
export const SCRIPT_MAX_TIMEOUT_MS = SCRIPT_LIMITS.maxTimeoutMs
export const SCRIPT_MAX_SOURCE_BYTES = SCRIPT_LIMITS.maxSourceBytes

/** Ceilings that keep one script from filling memory with a plan or a log. */
const MAX_OPS = 5000
const MAX_LOG_LINES = 200
const LOG_LINE_MAX_CHARS = 500

export { ScriptRunError, scriptFormatFor }
export type { ScriptFormat, ScriptSnapshot, SandboxOutcome }

export interface ScriptRunRequest {
  /** absolute path: messages and stack lines name the file */
  scriptPath: string
  source: string
  snapshot: ScriptSnapshot
  /** parameters visible as the second argument and as the input field */
  input: Record<string, unknown>
  timeoutMs?: number
}

/** Validate and clamp a --timeout value; throws the CLI's usage error shape. */
export function scriptTimeout(value: unknown): number {
  if (value === undefined) return SCRIPT_DEFAULT_TIMEOUT_MS
  const n = typeof value === 'number' ? value : Number(value)
  if (!Number.isInteger(n) || n <= 0) {
    throw new ScriptRunError(
      'timeout must be a positive whole number of milliseconds (default ' +
        SCRIPT_DEFAULT_TIMEOUT_MS +
        ', max ' +
        SCRIPT_MAX_TIMEOUT_MS +
        ')',
    )
  }
  return Math.min(n, SCRIPT_MAX_TIMEOUT_MS)
}

export function runScriptInSandbox(req: ScriptRunRequest): SandboxOutcome {
  const name = basename(req.scriptPath)
  const timeoutMs = scriptTimeout(req.timeoutMs)
  if (Buffer.byteLength(req.source, 'utf8') > SCRIPT_MAX_SOURCE_BYTES) {
    throw new ScriptRunError(
      name + ' is larger than ' + SCRIPT_MAX_SOURCE_BYTES + ' bytes',
      'shorten the script',
    )
  }
  // The file is a function body, not a module: a top-level import/export would
  // surface as a syntax error deep inside the wrapper, so say what is wrong instead.
  if (/^[ \t]*(export|import)[ \t]/m.test(req.source)) {
    throw new ScriptRunError(
      name + ': this build does not load modules',
      'remove import/export: the script is the body of a function that receives $ and input',
    )
  }
  const sandbox: Record<string, unknown> = Object.create(null)
  sandbox.Promise = undefined
  const context = createContext(sandbox, {
    codeGeneration: { strings: false, wasm: false },
    name: 'genoffice-script',
  })
  const started = Date.now()
  let encoded: unknown
  try {
    encoded = runInContext(buildSource(req, timeoutMs), context, {
      timeout: timeoutMs,
      displayErrors: true,
      filename: req.scriptPath,
    })
  } catch (err) {
    throw scriptError(err, name, timeoutMs, req.snapshot.format ?? 'docs')
  }
  const ms = Date.now() - started
  return parseOutcome(encoded, name, ms)
}

/**
 * The whole sandbox program: the shared root, then the format's namespace, then
 * the script as the body of a function, then the JSON the host reads back. The
 * snapshot and the input are spliced in as JSON, never as code.
 */
function buildSource(req: ScriptRunRequest, timeoutMs: number): string {
  const format = req.snapshot.format ?? 'docs'
  const deadline = Date.now() + timeoutMs
  return [
    'var __snapshot = ' + JSON.stringify(req.snapshot),
    'var __input = ' + JSON.stringify(req.input ?? {}),
    'var __deadline = ' + deadline,
    'var __timeoutMs = ' + timeoutMs,
    'var __plan = []',
    'var __log = []',
    'var __veto = []',
    'var __outText',
    'var __result',
    'var __pushVeto = function (reason) {',
    '  if (__veto.length >= 8) return',
    '  __veto.push(String(reason === undefined ? "no reason given" : reason).slice(0, 200))',
    '}',
    'var __pushLog = function (v) {',
    '  var s = typeof v === "string" ? v : (function () { try { return JSON.stringify(v) } catch (e) { return String(v) } })()',
    '  if (typeof s !== "string") s = String(s)',
    '  if (s.length > ' + LOG_LINE_MAX_CHARS + ') s = s.slice(0, ' + LOG_LINE_MAX_CHARS + ')',
    '  if (__log.length < ' + MAX_LOG_LINES + ') __log.push(s)',
    '}',
    'var __check = function () {',
    '  if (Date.now() > __deadline) throw new Error("script budget exceeded (" + __timeoutMs + "ms)")',
    '}',
    'globalThis.console = { log: __pushLog, info: __pushLog, warn: __pushLog, error: __pushLog }',
    'var __propose = function (ops) {',
    '  __check()',
    '  if (!Array.isArray(ops)) throw new Error("propose(ops) takes an array of ops")',
    '  for (var i = 0; i < ops.length; i++) {',
    '    var op = ops[i]',
    '    if (!op || typeof op !== "object" || Array.isArray(op)) throw new Error("op " + i + " is not an object")',
    '    if (typeof op.op !== "string" || !op.op) throw new Error("op " + i + " has no op name")',
    '    if (__plan.length >= ' + MAX_OPS + ') throw new Error("too many ops proposed")',
    '    __plan.push(op)',
    '  }',
    '  return __plan.length',
    '}',
    'var __merge = function (value) {',
    '  var list = Array.isArray(value) ? value : null',
    '  if (!list && value && typeof value === "object" && Array.isArray(value.ops)) list = value.ops',
    '  if (!list) return',
    '  for (var i = 0; i < list.length; i++) {',
    '    var op = list[i]',
    '    if (!op || typeof op !== "object" || Array.isArray(op) || typeof op.op !== "string") continue',
    '    if (__plan.length >= ' + MAX_OPS + ') break',
    '    __plan.push(op)',
    '  }',
    '}',
    'var __captureText = function (value) {',
    '  if (value && typeof value === "object" && typeof value.text === "string") value = value.text',
    '  if (typeof value === "string") __outText = value',
    '}',
    'var __captureResult = function (value) {',
    '  if (value === undefined) return',
    '  try { __result = JSON.stringify(value) } catch (e) { __result = String(value) }',
    '  if (typeof __result === "string" && __result.length > 4000) __result = __result.slice(0, 4000) + "...(clipped)"',
    '}',
    ...formatHelpers(format),
    'globalThis.$ = {',
    '  input: __input,',
    '  log: __pushLog,',
    '  event: {',
    '    name: __input.event && __input.event.name ? String(__input.event.name) : "",',
    '    path: __input.event && __input.event.path ? String(__input.event.path) : "",',
    '    veto: __pushVeto,',
    '  },',
    '  budget: {',
    '    get totalMs() { return __timeoutMs },',
    '    get remainingMs() { return Math.max(0, __deadline - Date.now()) },',
    '  },',
    ...formatNamespace(format),
    '}',
    'var __ret = (function ($, input) {',
    req.source,
    '})($, __input)',
    ...formatCapture(format),
    'JSON.stringify({ plan: __plan, log: __log, vetoes: __veto, text: __outText, result: __result })',
  ].join('\n')
}

/** The namespace one format mounts on $, as sandbox source lines (already indented two spaces). */
function formatNamespace(format: ScriptFormat): string[] {
  switch (format) {
    case 'docs':
      return docsNamespace()
    case 'sheets':
      return sheetsNamespace()
    case 'text':
      return textNamespace()
    case 'deck':
      return deckNamespace()
    case 'pdf':
      return pdfNamespace()
  }
}

/** How the script's return value is read back, per format. */
function formatCapture(format: ScriptFormat): string[] {
  switch (format) {
    case 'docs':
    case 'sheets':
    case 'deck':
      return ['__merge(__ret)']
    case 'text':
      return ['__captureText(__ret)']
    case 'pdf':
      return ['__captureResult(__ret)']
  }
}

function docsNamespace(): string[] {
  return [
    '  docs: {',
    '    blocks: function (range, maxChars) {',
    '      __check()',
    '      var from = 0',
    '      var to = __snapshot.blocks.length - 1',
    '      if (Array.isArray(range)) {',
    '        if (typeof range[0] === "number") from = range[0]',
    '        if (typeof range[1] === "number") to = range[1]',
    '      }',
    '      var cap = typeof maxChars === "number" && maxChars > 0 ? maxChars : 0',
    '      var out = []',
    '      for (var i = 0; i < __snapshot.blocks.length; i++) {',
    '        var b = __snapshot.blocks[i]',
    '        if (b.index < from || b.index > to) continue',
    '        var text = b.text',
    '        if (cap && typeof text === "string" && text.length > cap) {',
    '          text = text.slice(0, cap) + "...(+" + (text.length - cap) + " chars)"',
    '        }',
    '        out.push({ index: b.index, type: b.type, level: b.level, kind: b.kind, text: text })',
    '      }',
    '      return out',
    '    },',
    '    text: function () { __check(); return __snapshot.text },',
    '    propose: __propose,',
    '  },',
  ]
}

function sheetsNamespace(): string[] {
  return [
    '  sheets: {',
    '    sheets: function () {',
    '      __check()',
    '      return __snapshot.sheets.map(function (s) {',
    '        return { name: s.name, first: s.first, last: s.last, rows: s.rows, cols: s.cols, active: s.name === __snapshot.active }',
    '      })',
    '    },',
    '    usedRange: function (sheet) {',
    '      __check()',
    '      var s = __sheet(sheet)',
    '      return { sheet: s.name, first: s.first, last: s.last, rows: s.rows, cols: s.cols }',
    '    },',
    '    range: function (spec, maxCells) {',
    '      __check()',
    '      var r = __range(spec)',
    '      var s = __sheet(r.sheet)',
    '      var cells = (r.c2 - r.c1 + 1) * (r.r2 - r.r1 + 1)',
    '      var cap = typeof maxCells === "number" && maxCells > 0 ? Math.min(maxCells, ' +
      SHEETS_READ_CELL_CAP +
      ') : ' +
      SHEETS_READ_CELL_CAP,
    '      if (cells > cap) throw new Error("range " + spec + " covers " + cells + " cells; the read cap is " + cap + " — narrow it or raise maxCells (hard cap ' +
      SHEETS_READ_CELL_CAP +
      ')")',
    '      var values = []',
    '      for (var row = r.r1; row <= r.r2; row++) {',
    '        var line = []',
    '        for (var col = r.c1; col <= r.c2; col++) {',
    '          var v = s.cells[__cell(col, row)]',
    '          line.push(v === undefined ? null : v)',
    '        }',
    '        values.push(line)',
    '      }',
    '      return { sheet: s.name, range: __cell(r.c1, r.r1) + ":" + __cell(r.c2, r.r2), values: values }',
    '    },',
    '    formulas: function (sheet) {',
    '      __check()',
    '      var s = __sheet(sheet)',
    '      var out = {}',
    '      for (var addr in s.formulas) out[addr] = s.formulas[addr]',
    '      return out',
    '    },',
    '    formats: function (spec) {',
    '      __check()',
    '      var r = __range(spec)',
    '      var s = __sheet(r.sheet)',
    '      var cells = (r.c2 - r.c1 + 1) * (r.r2 - r.r1 + 1)',
    '      if (cells > ' +
      SHEETS_FORMAT_CAP +
      ') throw new Error("range " + spec + " covers " + cells + " cells; the formats cap is ' +
      SHEETS_FORMAT_CAP +
      ' — narrow it")',
    '      var out = {}',
    '      for (var row = r.r1; row <= r.r2; row++) {',
    '        for (var col = r.c1; col <= r.c2; col++) {',
    '          var addr = __cell(col, row)',
    '          if (s.formats[addr]) out[addr] = s.formats[addr]',
    '        }',
    '      }',
    '      return out',
    '    },',
    '    find: function (query) {',
    '      __check()',
    '      var q = String(query).toLowerCase()',
    '      if (!q) throw new Error("find(query) needs something to look for")',
    '      var out = []',
    '      var truncated = false',
    '      for (var i = 0; i < __snapshot.sheets.length; i++) {',
    '        var s = __snapshot.sheets[i]',
    '        for (var addr in s.cells) {',
    '          var value = s.cells[addr]',
    '          var formula = s.formulas[addr]',
    '          var hit = (typeof value === "string" && value.toLowerCase().indexOf(q) !== -1) || (formula && formula.toLowerCase().indexOf(q) !== -1)',
    '          if (!hit) continue',
    '          if (out.length >= ' + SHEETS_FIND_CAP + ') { truncated = true; break }',
    '          out.push({ sheet: s.name, ref: addr, value: value, formula: formula })',
    '        }',
    '        if (truncated) break',
    '      }',
    '      out.truncated = truncated',
    '      return out',
    '    },',
    '    propose: __propose,',
    '  },',
  ]
}

/** Helper declarations a format's namespace needs, placed after the $ object closes. */
function formatHelpers(format: ScriptFormat): string[] {
  if (format !== 'sheets') return []
  return [
    'var __sheet = function (name) {',
    '  var list = __snapshot.sheets',
    '  var wanted = name === undefined || name === null || name === "" ? __snapshot.active : String(name)',
    '  for (var i = 0; i < list.length; i++) if (list[i].name === wanted) return list[i]',
    '  var names = []',
    '  for (var j = 0; j < list.length; j++) names.push(list[j].name)',
    '  throw new Error("no sheet named " + wanted + " ($.sheets.sheets() lists: " + names.join(", ") + ")")',
    '}',
    'var __colIndex = function (letters) {',
    '  var n = 0',
    '  for (var i = 0; i < letters.length; i++) n = n * 26 + (letters.charCodeAt(i) - 64)',
    '  return n - 1',
    '}',
    'var __cell = function (col, row) {',
    '  var s = ""',
    '  var c = col + 1',
    '  while (c > 0) { var m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = (c - 1 - m) / 26 }',
    '  return s + (row + 1)',
    '}',
    'var __range = function (spec) {',
    '  var text = String(spec)',
    '  var bang = text.indexOf("!")',
    '  var sheet = bang === -1 ? undefined : text.slice(0, bang)',
    '  var addr = bang === -1 ? text : text.slice(bang + 1)',
    '  var m = /^([A-Za-z]+)(\\d+)(?::([A-Za-z]+)(\\d+))?$/.exec(addr)',
    '  if (!m) throw new Error("range must look like Sheet1!A2:F200 or A2:F200 (got " + spec + ")")',
    '  var c1 = __colIndex(m[1].toUpperCase()), r1 = Number(m[2]) - 1',
    '  var c2 = m[3] ? __colIndex(m[3].toUpperCase()) : c1, r2 = m[4] ? Number(m[4]) - 1 : r1',
    '  if (c2 < c1) { var t = c1; c1 = c2; c2 = t }',
    '  if (r2 < r1) { var u = r1; r1 = r2; r2 = u }',
    '  return { sheet: sheet, c1: c1, r1: r1, c2: c2, r2: r2 }',
    '}',
  ]
}

function textNamespace(): string[] {
  return [
    '  text: {',
    '    get: function () { __check(); return __snapshot.text },',
    '    lines: function () { __check(); return __snapshot.text.split("\\n") },',
    '  },',
  ]
}

function deckNamespace(): string[] {
  return [
    '  deck: {',
    '    size: __snapshot.deck.size,',
    '    slides: function () {',
    '      __check()',
    '      return __snapshot.deck.pages',
    '    },',
    '    elements: function (slide) {',
    '      __check()',
    '      var page = __snapshot.deck.pages[slide]',
    '      if (!page) throw new Error("no slide " + slide + " ($.deck.slides() lists them; 0-based)")',
    '      return page.elements',
    '    },',
    '    propose: __propose,',
    '  },',
  ]
}

function pdfNamespace(): string[] {
  return [
    '  pdf: {',
    '    info: function () {',
    '      __check()',
    '      var p = __snapshot.pdf',
    '      return { pages: p.pages, encrypted: p.encrypted, title: p.title, author: p.author, subject: p.subject, creator: p.creator, producer: p.producer }',
    '    },',
    '    pages: function (from, to, maxChars) {',
    '      __check()',
    '      from = Math.floor(Number(from))',
    '      to = to === undefined ? from : Math.floor(Number(to))',
    '      if (!(from >= 1) || !(to >= from)) throw new Error("pages(from, to) takes a 1-based inclusive span (got " + from + ", " + to + ")")',
    '      if (to > __snapshot.pdf.pageCap) throw new Error("this snapshot holds the first " + __snapshot.pdf.pagesRead + " of " + __snapshot.pdf.pages + " page(s); $.pdf.pages cannot read past page " + __snapshot.pdf.pageCap)',
    '      var cap = typeof maxChars === "number" && maxChars > 0 ? maxChars : 0',
    '      var out = []',
    '      for (var i = 0; i < __snapshot.pdf.pageTexts.length; i++) {',
    '        var page = __snapshot.pdf.pageTexts[i]',
    '        if (page.page < from || page.page > to) continue',
    '        var text = page.text',
    '        if (cap && text.length > cap) text = text.slice(0, cap) + "...(+" + (text.length - cap) + " chars)"',
    '        out.push({ page: page.page, width_pt: page.width_pt, height_pt: page.height_pt, text: text, truncated: page.truncated })',
    '      }',
    '      if (!out.length) throw new Error("no pages in " + from + "-" + to + " (the snapshot holds pages 1-" + __snapshot.pdf.pagesRead + " of " + __snapshot.pdf.pages + ")")',
    '      return out',
    '    },',
    '  },',
  ]
}

function parseOutcome(encoded: unknown, name: string, ms: number): SandboxOutcome {
  let parsed: unknown
  try {
    parsed = JSON.parse(String(encoded))
  } catch {
    throw new ScriptRunError(
      name + ': the plan could not be read back',
      'the script host is out of step; report this',
    )
  }
  const record = (parsed ?? {}) as {
    plan?: unknown
    log?: unknown
    vetoes?: unknown
    text?: unknown
    result?: unknown
  }
  const ops = Array.isArray(record.plan) ? (record.plan as Record<string, unknown>[]) : []
  const log = Array.isArray(record.log) ? record.log.map((line) => String(line)) : []
  const vetoes = Array.isArray(record.vetoes) ? record.vetoes.map((line) => String(line)) : []
  return {
    ops,
    log,
    vetoes,
    ms,
    ...(typeof record.text === 'string' ? { text: record.text } : {}),
    ...(typeof record.result === 'string' ? { result: record.result } : {}),
  }
}

function scriptError(
  err: unknown,
  name: string,
  timeoutMs: number,
  format: ScriptFormat,
): ScriptRunError {
  const anyErr = err as { message?: unknown; name?: unknown; code?: unknown } | null
  const message = typeof anyErr?.message === 'string' ? anyErr.message : String(err)
  if (anyErr?.code === 'ERR_SCRIPT_EXECUTION_TIMEOUT') {
    return new ScriptRunError(
      name + ' exceeded its ' + timeoutMs + 'ms budget',
      'raise the timeout (max ' + SCRIPT_MAX_TIMEOUT_MS + ') or make the script cheaper',
    )
  }
  if (anyErr?.name === 'SyntaxError') {
    return new ScriptRunError(
      name + ': ' + message,
      'the script is the body of a function: plain statements, no import/export/await; it receives $ and input',
    )
  }
  return new ScriptRunError(
    name + ': ' + message,
    'the script failed; $.' + format + ' is all a script can reach',
  )
}

/** The read-only view of a document a docs script is handed. */
function snapshotOf(doc: Awaited<ReturnType<typeof openDocument>>): DocsSnapshot {
  const blocks = describeDocument(doc, undefined, PREVIEW_CHARS)
  return {
    format: 'docs',
    blocks: blocks.map((block) => ({
      index: block.index,
      type: block.type,
      text: block.text,
      ...(block.level === undefined ? {} : { level: block.level }),
      ...(block.kind === undefined ? {} : { kind: block.kind }),
    })),
    text: blocks.map((block) => block.text).join('\n'),
  }
}

/** What a script proposed, before the engine has judged it. */
function planWith(
  scriptId: string,
  source: string,
  snapshot: ScriptSnapshot,
  input: Record<string, unknown>,
  timeoutMs: number | undefined,
): SandboxOutcome {
  return runScriptInSandbox({
    scriptPath: scriptId + '.js',
    source,
    snapshot,
    input,
    ...(timeoutMs === undefined ? {} : { timeoutMs }),
  })
}

/** The format a run or apply executes as, with the declared-format check. */
function requireFormat(documentPath: string, declared: ScriptFormat | undefined): ScriptFormat {
  const format = scriptFormatFor(documentPath)
  if (!format) {
    throw new ScriptRunError(
      'scripts do not support this file type (.' + (documentPath.split('.').pop() ?? '') + ')',
      'scripts run against one of: ' + formatList(),
    )
  }
  if (declared && declared !== format) {
    throw new ScriptRunError(
      'this script targets ' +
        declared +
        ' documents, but ' +
        basename(documentPath) +
        ' is a ' +
        format +
        ' file',
      'point it at a .' + declared + ' file, or update the feature with the right format',
    )
  }
  return format
}

function formatList(): string {
  return SCRIPT_FORMATS.join(', ')
}

/** What one format contributes to the host: its snapshot, and how it judges a proposal. */
interface FormatHost {
  snapshot(documentPath: string): Promise<ScriptSnapshot>
  plan(args: {
    scriptId: string
    outcome: SandboxOutcome
    documentPath: string
  }): ScriptRunOutcome | Promise<ScriptRunOutcome>
  apply?(args: {
    scriptId: string
    outcome: SandboxOutcome
    documentPath: string
  }): ScriptApplyResult | Promise<ScriptApplyResult>
}

function formatHost(format: ScriptFormat): FormatHost {
  switch (format) {
    case 'sheets':
      return sheetsHost
    case 'text':
      return textHost
    case 'deck':
      return deckHost
    case 'pdf':
      return pdfHost
    case 'docs':
      throw new ScriptRunError('docs runs its own path in this module', 'a host bug; report this')
  }
}

/**
 * Run one script feature read-only against a document: open it, hand the script
 * the format's snapshot, and judge what it proposed with the same executor the
 * apply path uses. Nothing is written — not the document on disk, not a temporary
 * file.
 *
 * For docs this validates the plan on the in-memory editor, which is the only way
 * index shifts from an earlier entry are taken into account when a later one is
 * validated; the other formats run their own engine the same way.
 */
export async function runScriptReadOnly(request: {
  /** the registry id, used in messages */
  scriptId: string
  source: string
  /** absolute path of the document the script reads */
  documentPath: string
  /** the format the feature declared, when it declared one */
  format?: ScriptFormat
  input?: Record<string, unknown>
  timeoutMs?: number
}): Promise<ScriptRunOutcome> {
  const format = requireFormat(request.documentPath, request.format)
  if (format === 'docs') return docsRun(request, false)
  const host = formatHost(format)
  const snapshot = await host.snapshot(request.documentPath)
  const sandbox = planWith(
    request.scriptId,
    request.source,
    snapshot,
    request.input ?? {},
    request.timeoutMs,
  )
  return Promise.resolve(
    host.plan({ scriptId: request.scriptId, outcome: sandbox, documentPath: request.documentPath }),
  )
}

async function docsRun(
  request: {
    scriptId: string
    source: string
    documentPath: string
    input?: Record<string, unknown>
    timeoutMs?: number
  },
  write: false,
): Promise<ScriptRunOutcome>
async function docsRun(
  request: {
    scriptId: string
    source: string
    documentPath: string
    input?: Record<string, unknown>
    timeoutMs?: number
  },
  write: true,
): Promise<ScriptApplyResult>
async function docsRun(
  request: {
    scriptId: string
    source: string
    documentPath: string
    input?: Record<string, unknown>
    timeoutMs?: number
  },
  write: boolean,
): Promise<ScriptRunOutcome | ScriptApplyResult> {
  const doc = await openDocument(readInput(request.documentPath))
  try {
    const sandbox = planWith(
      request.scriptId,
      request.source,
      snapshotOf(doc),
      request.input ?? {},
      request.timeoutMs,
    )
    if (write && !sandbox.ops.length) {
      throw new ScriptRunError(
        'the script proposed nothing, so there is nothing to apply',
        'check what the script reads: $.docs.blocks() is the whole document',
      )
    }
    if (!sandbox.ops.length) {
      return { applied: 0, plan: [], failures: [], log: sandbox.log, ms: sandbox.ms }
    }
    const outcome = await applyDocOps(doc, sandbox.ops, {
      ctx: { cwd: dirname(request.documentPath), env: process.env },
      mode: write ? 'atomic' : 'best_effort',
    })
    if (write && outcome.failures.length) {
      throw new ScriptRunError(
        'the engine refused ' + outcome.failures.length + ' operation(s), so nothing was written',
        outcome.failures[0]?.error,
      )
    }
    const plan = outcome.results.map(
      (result) => result.index + ': ' + result.op + ': ' + (result.output.split('\n')[0] ?? ''),
    )
    const failures = outcome.failures.map(
      (failure) => failure.index + ' ' + failure.op + ': ' + failure.error,
    )
    if (!write) {
      return { applied: outcome.results.length, plan, failures, log: sandbox.log, ms: sandbox.ms }
    }
    const bytes = await saveDocument(doc)
    // replace in one step: a reader never sees a half-written file
    const temp = request.documentPath + '.genoffice-apply-' + process.pid
    writeFileSync(temp, bytes)
    renameSync(temp, request.documentPath)
    return {
      applied: outcome.results.length,
      plan,
      failures: [],
      log: sandbox.log,
      ms: sandbox.ms,
      written: request.documentPath,
      bytes: bytes.length,
    }
  } finally {
    closeDocument(doc)
  }
}

/**
 * Apply a script feature to one document — the write half of the script kind.
 *
 * The script still proposes and nothing else: the host opens the document, runs
 * the sandbox, and only then, if the format's engine accepted **every**
 * operation, writes with the engine's own writer and replaces the file in one
 * step. A refusal, a thrown script or a timeout leaves the file exactly as it
 * was. One deliberate format difference: pdf scripts are read-only today, so
 * there is nothing to apply.
 *
 * Two gates beyond that: an explicit confirm, and the caller's open-document
 * check (a file an editor has open belongs to that editor; writing it here would
 * be lost on the editor's next save).
 */
export async function applyScriptToDocument(request: {
  scriptId: string
  source: string
  documentPath: string
  /** a write needs a deliberate yes: there is no default */
  confirm: boolean
  format?: ScriptFormat
  input?: Record<string, unknown>
  timeoutMs?: number
}): Promise<ScriptApplyResult> {
  if (request.confirm !== true) {
    throw new ScriptRunError(
      'writing a document needs an explicit confirmation',
      'run it read-only first, then ask again with confirm: true',
    )
  }
  const format = requireFormat(request.documentPath, request.format)
  if (format === 'pdf') {
    throw new ScriptRunError(
      'pdf scripts are read-only today',
      'the dry run returns the script\u2019s report as its result; nothing on a pdf is written',
    )
  }
  if (format === 'docs') return docsRun(request, true)
  const host = formatHost(format)
  if (!host.apply) {
    throw new ScriptRunError('this format has no apply path', 'a host bug; report this')
  }
  const snapshot = await host.snapshot(request.documentPath)
  const sandbox = planWith(
    request.scriptId,
    request.source,
    snapshot,
    request.input ?? {},
    request.timeoutMs,
  )
  return Promise.resolve(
    host.apply({
      scriptId: request.scriptId,
      outcome: sandbox,
      documentPath: request.documentPath,
    }),
  )
}
