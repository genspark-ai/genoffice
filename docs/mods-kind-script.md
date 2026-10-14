# The script kind (a later PR, not PR 1) — kept code

PR 1 became "expose the interface to the mod assistant" (see docs/mods-pr1-mod-interface.md).
The script host and command below were written for the earlier framing and are parked here so the
work is not lost. They are not wired into the CLI. When the script kind becomes its due PR, restore:
packages/cli/src/script/host.ts, packages/cli/src/commands/script.ts and the matching test, register
the command in cli.ts, and re-add the README/skill lines the sync test requires.

## packages/cli/src/script/host.ts

    import { basename } from 'node:path'
    import { createContext, runInContext } from 'node:vm'
    import { CliError, EXIT } from '../result'

    /**
     * The script host: one script, one read-only snapshot, one plan.
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

    export const SCRIPT_DEFAULT_TIMEOUT_MS = 5000
    export const SCRIPT_MAX_TIMEOUT_MS = 30_000
    export const SCRIPT_MAX_SOURCE_BYTES = 256 * 1024

    /** Ceilings that keep one script from filling memory with a plan or a log. */
    const MAX_OPS = 5000
    const MAX_LOG_LINES = 200
    const LOG_LINE_MAX_CHARS = 500

    /** One block of the read-only snapshot a script sees. */
    export interface ScriptBlock {
      index: number
      type: string
      level?: number
      kind?: string
      text: string
    }

    export interface ScriptSnapshot {
      blocks: ScriptBlock[]
      text: string
    }

    export interface ScriptRunRequest {
      /** absolute path: messages and stack lines name the file */
      scriptPath: string
      source: string
      snapshot: ScriptSnapshot
      /** parameters visible as the second argument and as the input field */
      input: Record<string, unknown>
      timeoutMs?: number
    }

    export interface ScriptRunOutcome {
      /** the operations the script proposed, in order; the caller validates them */
      ops: Record<string, unknown>[]
      log: string[]
      ms: number
    }

    /** Validate and clamp a --timeout value; throws the CLI's usage error shape. */
    export function scriptTimeout(value: unknown): number {
      if (value === undefined) return SCRIPT_DEFAULT_TIMEOUT_MS
      const n = typeof value === 'number' ? value : Number(value)
      if (!Number.isInteger(n) || n <= 0) {
        throw new CliError(
          EXIT.usage,
          '--timeout must be a positive whole number of milliseconds',
          { default: SCRIPT_DEFAULT_TIMEOUT_MS, max: SCRIPT_MAX_TIMEOUT_MS },
          { reason: 'invalid_argument' },
        )
      }
      return Math.min(n, SCRIPT_MAX_TIMEOUT_MS)
    }

    export function runScriptInSandbox(req: ScriptRunRequest): ScriptRunOutcome {
      const name = basename(req.scriptPath)
      const timeoutMs = scriptTimeout(req.timeoutMs)
      if (Buffer.byteLength(req.source, 'utf8') > SCRIPT_MAX_SOURCE_BYTES) {
        throw new CliError(
          EXIT.usage,
          name + ' is larger than ' + SCRIPT_MAX_SOURCE_BYTES + ' bytes',
          undefined,
          { reason: 'resource_limit' },
        )
      }
      // The file is a function body, not a module: a top-level import/export would
      // surface as a syntax error deep inside the wrapper, so say what is wrong instead.
      if (/^[ \t]*(export|import)[ \t]/m.test(req.source)) {
        throw new CliError(EXIT.usage, name + ': this build does not load modules', undefined, {
          reason: 'invalid_argument',
          suggestion:
            'remove import/export: the file is the body of a function that receives $ and input',
        })
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
        throw scriptError(err, name, timeoutMs)
      }
      const ms = Date.now() - started
      const parsed = parseOutcome(encoded, name)
      return { ops: parsed.ops, log: parsed.log, ms }
    }

    /**
     * The whole sandbox program: the API, then the script as the body of a function,
     * then the JSON the host reads back. The snapshot and the input are spliced in as
     * JSON, never as code.
     */
    function buildSource(req: ScriptRunRequest, timeoutMs: number): string {
      const deadline = Date.now() + timeoutMs
      return [
        'var __snapshot = ' + JSON.stringify(req.snapshot),
        'var __input = ' + JSON.stringify(req.input ?? {}),
        'var __deadline = ' + deadline,
        'var __timeoutMs = ' + timeoutMs,
        'var __plan = []',
        'var __log = []',
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
        'globalThis.$ = {',
        '  input: __input,',
        '  log: __pushLog,',
        '  budget: {',
        '    get totalMs() { return __timeoutMs },',
        '    get remainingMs() { return Math.max(0, __deadline - Date.now()) },',
        '  },',
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
        '    propose: function (ops) {',
        '      __check()',
        '      if (!Array.isArray(ops)) throw new Error("propose(ops) takes an array of ops")',
        '      for (var i = 0; i < ops.length; i++) {',
        '        var op = ops[i]',
        '        if (!op || typeof op !== "object" || Array.isArray(op)) throw new Error("op " + i + " is not an object")',
        '        if (typeof op.op !== "string" || !op.op) throw new Error("op " + i + " has no op name")',
        '        if (__plan.length >= ' + MAX_OPS + ') throw new Error("too many ops proposed")',
        '        __plan.push(op)',
        '      }',
        '      return __plan.length',
        '    },',
        '  },',
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
        'var __ret = (function ($, input) {',
        req.source,
        '})($, __input)',
        '__merge(__ret)',
        'JSON.stringify({ plan: __plan, log: __log })',
      ].join('\n')
    }

    function parseOutcome(
      encoded: unknown,
      name: string,
    ): { ops: Record<string, unknown>[]; log: string[] } {
      let parsed: unknown
      try {
        parsed = JSON.parse(String(encoded))
      } catch {
        throw new CliError(EXIT.conversion, name + ': the plan could not be read back', undefined, {
          reason: 'conversion_failed',
        })
      }
      const record = (parsed ?? {}) as { plan?: unknown; log?: unknown }
      const ops = Array.isArray(record.plan) ? (record.plan as Record<string, unknown>[]) : []
      const log = Array.isArray(record.log) ? record.log.map((line) => String(line)) : []
      return { ops, log }
    }

    function scriptError(err: unknown, name: string, timeoutMs: number): CliError {
      const anyErr = err as { message?: unknown; name?: unknown; code?: unknown } | null
      const message = typeof anyErr?.message === 'string' ? anyErr.message : String(err)
      if (anyErr?.code === 'ERR_SCRIPT_EXECUTION_TIMEOUT') {
        return new CliError(
          EXIT.usage,
          name + ' exceeded its ' + timeoutMs + 'ms budget',
          { timeout_ms: timeoutMs },
          {
            reason: 'resource_limit',
            suggestion:
              'raise --timeout (max ' + SCRIPT_MAX_TIMEOUT_MS + ') or make the script cheaper',
          },
        )
      }
      if (anyErr?.name === 'SyntaxError') {
        return new CliError(EXIT.usage, name + ': ' + message, undefined, {
          reason: 'invalid_argument',
          suggestion:
            'the file is the body of a function: plain statements, no import/export/await; it receives $ and input',
        })
      }
      return new CliError(EXIT.usage, name + ': ' + message, undefined, {
        reason: 'invalid_argument',
        suggestion: 'the script failed; the docs reader lists what it can read',
      })
    }

## packages/cli/src/commands/script.ts

    import { readFileSync } from 'node:fs'
    import { basename } from 'node:path'
    import { flagBool, flagString } from '../args'
    import { batchCounts, batchResult } from '../batch'
    import {
      PREVIEW_CHARS,
      applyDocOps,
      closeDocument,
      describeDocument,
      openDocument,
      type BlockSummary,
    } from '../formats/docx'
    import { readInput, resolveInput } from '../fs'
    import { opSuggestion } from '../op-errors'
    import type { CommandContext, CommandDef } from '../registry'
    import { CliError, EXIT } from '../result'
    import { runScriptInSandbox, type ScriptBlock } from '../script/host'

    /**
     * A script reads the open document and proposes operations; this command runs it
     * against a read-only snapshot, validates the proposals with the same executor
     * the apply paths use, and writes nothing. Applying is the next change, which is
     * why --dry-run is required here rather than optional.
     */
    export const scriptCommand: CommandDef = {
      name: 'script',
      summary: 'Run a JavaScript plan against a read-only view of a document.',
      usage: 'script run <file.js> --doc <file.docx> --dry-run [options]',
      options: [
        {
          name: 'doc',
          value: 'file',
          description: 'run: the .docx the script reads; the script cannot write it',
        },
        {
          name: 'input',
          value: 'json|file',
          description:
            'run: parameters for the script, as inline JSON or a .json file; exposed as .input',
        },
        {
          name: 'timeout',
          value: 'ms',
          description: 'run: wall-clock budget for the script (default 5000, max 30000)',
        },
        {
          name: 'dry-run',
          description:
            'run: validate the proposed ops in memory and print the plan; the only mode in this build',
        },
      ],
      async run(args, ctx) {
        const [verb, scriptFile] = args.positionals
        if (verb !== 'run') {
          throw new CliError(
            EXIT.usage,
            'expected "script run <file.js> --doc <file.docx> --dry-run"',
            undefined,
            {
              reason: verb === undefined ? 'missing_argument' : 'invalid_argument',
              suggestion: 'run genoffice help script',
            },
          )
        }
        if (!flagBool(args, 'dry-run')) {
          throw new CliError(
            EXIT.usage,
            'a script cannot be applied yet: this build only plans',
            undefined,
            { reason: 'unsupported', suggestion: 'add --dry-run to print the validated plan' },
          )
        }
        const scriptPath = resolveInput(scriptFile, ctx)
        const docPath = resolveInput(flagString(args, 'doc'), ctx)
        const timeout = flagString(args, 'timeout')
        const input = readParams(flagString(args, 'input'), ctx)
        const source = readFileSync(scriptPath, 'utf8')
        const doc = await openDocument(readInput(docPath))
        try {
          const blocks = describeDocument(doc, undefined, PREVIEW_CHARS)
          const outcome = runScriptInSandbox({
            scriptPath,
            source,
            snapshot: {
              blocks: blocks.map(toScriptBlock),
              text: blocks.map((b) => b.text).join('\n'),
            },
            input,
            ...(timeout === undefined ? {} : { timeoutMs: Number(timeout) }),
          })
          const common = {
            script: basename(scriptPath),
            document: basename(docPath),
            ops: outcome.ops.length,
            ms: outcome.ms,
            ...(outcome.log.length ? { log: outcome.log } : {}),
          }
          if (!outcome.ops.length) {
            return {
              summary: basename(scriptPath) + ': the script proposed no changes',
              detail: common,
            }
          }
          // The dry run is the engine's own: the batch runs on the in-memory editor
          // and is never saved, which is the only way index shifts from an earlier
          // entry are taken into account when a later one is validated.
          const { results, failures } = await applyDocOps(doc, outcome.ops, {
            ctx,
            mode: 'best_effort',
          })
          const counts = batchCounts(outcome.ops.length, results.length, failures.length)
          if (!results.length) {
            const first = failures[0]!
            throw new CliError(
              EXIT.usage,
              'the plan was rejected: op ' + first.index + ' (' + first.op + '): ' + first.error,
              { failures, batch: counts, script: basename(scriptPath) },
              { reason: first.reason, suggestion: opSuggestion(first, 'docs') },
            )
          }
          return batchResult(
            {
              summary:
                basename(scriptPath) +
                ': ' +
                results.length +
                ' of ' +
                outcome.ops.length +
                ' ops validated, nothing written',
              detail: {
                ...common,
                plan: results.map((r) => r.index + ': ' + r.op + ': ' + r.output.split('\n')[0]),
              },
            },
            counts,
            failures,
          )
        } finally {
          closeDocument(doc)
        }
      },
    }

    /** --input: inline JSON, or a path to a .json file. */
    function readParams(spec: string | undefined, ctx: CommandContext): Record<string, unknown> {
      if (spec === undefined) return {}
      let text = spec
      let source = '--input'
      if (!spec.trimStart().startsWith('{')) {
        const path = resolveInput(spec, ctx)
        text = readFileSync(path, 'utf8')
        source = basename(path)
      }
      let parsed: unknown
      try {
        parsed = JSON.parse(text)
      } catch (err) {
        throw new CliError(
          EXIT.usage,
          source + ': not valid JSON (' + (err as Error).message + ')',
          undefined,
          { reason: 'invalid_json' },
        )
      }
      if (!parsed || typeof parsed !== 'object' || Array.isArray(parsed)) {
        throw new CliError(EXIT.usage, source + ': expected a JSON object of parameters', undefined, {
          reason: 'invalid_argument',
        })
      }
      return parsed as Record<string, unknown>
    }

    /** Only the fields a script may need: the table model stays out until a reader wants it. */
    function toScriptBlock(block: BlockSummary): ScriptBlock {
      return {
        index: block.index,
        type: block.type,
        text: block.text,
        ...(block.level === undefined ? {} : { level: block.level }),
        ...(block.kind === undefined ? {} : { kind: block.kind }),
      }
    }

## packages/cli/tests/script.test.ts

    import { readFileSync, writeFileSync } from 'node:fs'
    import { join, resolve } from 'node:path'
    import { describe, expect, it } from 'vitest'
    import { run, tempDir } from './helpers'

    const REPO = resolve(__dirname, '../../..')
    const DOCX = join(REPO, 'apps/docs/tests/pagination-corpus/docx/01-simple-english.docx')

    function script(dir: string, body: string): string {
      const path = join(dir, 'plan.js')
      writeFileSync(path, body)
      return path
    }

    function planArgs(path: string, extra: string[] = []): string[] {
      return ['script', 'run', path, '--doc', DOCX, '--dry-run', '--json', ...extra]
    }

    describe('script run', () => {
      it('plans from the read API, validates the ops, and writes nothing', async () => {
        const dir = tempDir()
        const before = readFileSync(DOCX)
        const r = await run(
          planArgs(
            script(
              dir,
              [
                'var blocks = $.docs.blocks()',
                '$.log("blocks=" + blocks.length)',
                '$.docs.propose([{ op: "insert_content", afterBlockIndex: 0, html: "<p>from script</p>" }])',
              ].join('\n'),
            ),
          ),
        )
        expect(r.code).toBe(0)
        const json = r.json()
        expect(json).toMatchObject({
          status: 'ok',
          command: 'script',
          detail: { ops: 1, batch: { applied: 1 } },
        })
        expect(json.detail.log[0]).toMatch(/^blocks=[0-9]+$/)
        expect(json.detail.plan).toHaveLength(1)
        expect(readFileSync(DOCX).equals(before)).toBe(true)
      })

      it('exposes nothing beyond the API: process, require and timers are undefined', async () => {
        const dir = tempDir()
        const r = await run(
          planArgs(
            script(dir, '$.log(typeof process + "|" + typeof require + "|" + typeof setTimeout)'),
          ),
        )
        expect(r.code).toBe(0)
        expect(r.json().detail.log[0]).toBe('undefined|undefined|undefined')
      })

      it('refuses the constructor escapes', async () => {
        const dir = tempDir()
        const r = await run(
          planArgs(
            script(
              dir,
              [
                'try {',
                '  $.log("ESCAPED:" + typeof globalThis.constructor.constructor("return process")())',
                '} catch (e) {',
                '  $.log("BLOCKED")',
                '}',
              ].join('\n'),
            ),
          ),
        )
        expect(r.code).toBe(0)
        expect(r.json().detail.log[0]).toBe('BLOCKED')
      })

      it('stops a script that overruns its budget', async () => {
        const dir = tempDir()
        const r = await run([
          'script',
          'run',
          script(dir, 'while (true) {}'),
          '--doc',
          DOCX,
          '--dry-run',
          '--timeout',
          '300',
          '--json',
        ])
        expect(r.code).toBe(1)
        expect(r.json()).toMatchObject({ status: 'error', error: 'resource_limit' })
      })

      it('rejects an op the engine does not know', async () => {
        const dir = tempDir()
        const r = await run(planArgs(script(dir, '$.docs.propose([{ op: "nope" }])')))
        expect(r.code).toBe(1)
        expect(r.json()).toMatchObject({ status: 'error' })
      })

      it('requires --dry-run while applying is not implemented', async () => {
        const dir = tempDir()
        const r = await run(['script', 'run', script(dir, '$.log("x")'), '--doc', DOCX, '--json'])
        expect(r.code).toBe(1)
        expect(r.json()).toMatchObject({ status: 'error', error: 'unsupported' })
      })

      it('reports a missing --doc and a bad --timeout', async () => {
        const dir = tempDir()
        const path = script(dir, '$.log("x")')
        expect((await run(['script', 'run', path, '--dry-run', '--json'])).json()).toMatchObject({
          error: 'missing_argument',
        })
        expect(
          (
            await run(['script', 'run', path, '--doc', DOCX, '--dry-run', '--timeout', '0', '--json'])
          ).json(),
        ).toMatchObject({ error: 'invalid_argument' })
      })

      it('refuses module syntax with a readable message', async () => {
        const dir = tempDir()
        const r = await run(planArgs(script(dir, 'export default function () {}')))
        expect(r.code).toBe(1)
        const json = r.json()
        expect(json).toMatchObject({ status: 'error', error: 'invalid_argument' })
        expect(json.message).toContain('does not load modules')
      })

      it('hands --input to the script', async () => {
        const dir = tempDir()
        const r = await run(planArgs(script(dir, '$.log("n=" + $.input.n)'), ['--input', '{"n":3}']))
        expect(r.code).toBe(0)
        expect(r.json().detail.log[0]).toBe('n=3')
      })

      it('treats an empty plan as a successful no-op', async () => {
        const dir = tempDir()
        const r = await run(planArgs(script(dir, '$.log("nothing to do")')))
        expect(r.code).toBe(0)
        expect(r.json().summary).toContain('proposed no changes')
      })
    })
