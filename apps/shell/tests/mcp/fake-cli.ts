import { readFileSync, writeFileSync } from 'node:fs'
import type { CliRunner, CliRunOutcome } from '../../src/main/mcp/cli-runner'

/**
 * Test double for the bundled genoffice CLI.
 *
 * The MCP headless tools are a *client* of the CLI: what they own is the argv
 * they build (and the stdin they pipe), not document fidelity — that belongs to
 * the CLI's own tests. So this fake records every call and writes a stub output
 * file, letting the tool tests assert delegation without spawning a process.
 * The stdin a tool pipes and any `--from <file>` content are snapshotted at call
 * time (the tool's scratch dir is cleaned up the moment the call returns).
 */

export interface CliCall {
  args: string[]
  input: string | undefined
  /** contents of the `--from` input file at call time, when one was passed */
  fromContent?: string
  /** contents of the `--ops` input file at call time, when one was passed */
  opsContent?: string
  /** contents of the `--data` input file at call time, when one was passed */
  dataContent?: string
}

export interface FakeCli {
  runner: CliRunner
  calls: CliCall[]
  /** the last call's args */
  last(): string[]
  /** the last call's stdin */
  lastInput(): string | undefined
  /** the last call's `--from` file contents */
  lastFrom(): string | undefined
  /** the last call's `--ops` file contents */
  lastOps(): string | undefined
  /** the last call's `--data` file contents */
  lastData(): string | undefined
  /** reply for `docs read` style commands: the blocks detail items */
  setReadItems(items: Array<Record<string, unknown>>): void
  /** reply for `docs read --properties`: the document properties object */
  setDocProperties(properties: Record<string, unknown>): void
  /** reply for `merge` commands: the result's `detail` object */
  setMergeDetail(detail: Record<string, unknown>): void
  /** reply for `docs check` commands: the result's `detail` object */
  setCheckDetail(detail: Record<string, unknown>): void
  /** reply for `sheet links` commands: the result's `detail` object */
  setSheetLinksDetail(detail: Record<string, unknown>): void
  /** make the next run fail with this CLI error message */
  failNext(message: string): void
}

export function fakeCli(): FakeCli {
  const calls: CliCall[] = []
  let readItems: Array<Record<string, unknown>> = []
  let docProperties: Record<string, unknown> | undefined
  let mergeDetail: Record<string, unknown> | undefined
  let checkDetail: Record<string, unknown> | undefined
  let sheetLinksDetail: Record<string, unknown> | undefined
  let nextError: string | undefined

  const readStaged = (args: string[], flag: string): string | undefined => {
    const index = args.indexOf(flag)
    if (index < 0 || !args[index + 1]) return undefined
    try {
      return readFileSync(args[index + 1]!, 'utf8')
    } catch {
      return undefined
    }
  }

  const runner: CliRunner = {
    async run(args, options): Promise<CliRunOutcome> {
      const call: CliCall = { args, input: options?.input }
      call.fromContent = readStaged(args, '--from')
      call.opsContent = readStaged(args, '--ops')
      call.dataContent = readStaged(args, '--data')
      calls.push(call)
      if (nextError !== undefined) {
        const message = nextError
        nextError = undefined
        return {
          ok: false,
          code: 1,
          json: { status: 'error', command: args[0] ?? null, code: 1, message },
          stdout: JSON.stringify({ status: 'error', command: args[0] ?? null, code: 1, message }),
          stderr: '',
        }
      }
      // `docs check`: the scripted quality report; `docs read`: the block items
      if (args[0] === 'docs') {
        if (args[1] === 'check') {
          const json = {
            status: 'ok' as const,
            command: 'docs',
            summary: 'check',
            detail: checkDetail ?? {},
          }
          return { ok: true, code: 0, json, stdout: JSON.stringify(json), stderr: '' }
        }
        const json = {
          status: 'ok' as const,
          command: 'docs',
          summary: 'read',
          detail: {
            items: readItems,
            range: `0-${Math.max(0, readItems.length - 1)}`,
            ...(docProperties !== undefined ? { properties: docProperties } : {}),
          },
        }
        return { ok: true, code: 0, json, stdout: JSON.stringify(json), stderr: '' }
      }
      // `merge`: write a stub file at --out and hand back the scripted detail
      if (args[0] === 'merge') {
        const outIndex = args.indexOf('--out')
        const out = outIndex >= 0 ? args[outIndex + 1] : undefined
        if (out) writeFileSync(out, 'stub')
        const json = {
          status: 'ok' as const,
          command: 'merge',
          summary: 'merged',
          ...(out ? { output_path: out } : {}),
          detail: mergeDetail ?? {},
        }
        return { ok: true, code: 0, json, stdout: JSON.stringify(json), stderr: '' }
      }
      // `sheet links`: hand back the scripted detail (links report or a mutation result)
      if (args[0] === 'sheet' && args[1] === 'links') {
        const json = {
          status: 'ok' as const,
          command: 'sheet',
          summary: 'sheet links',
          detail: sheetLinksDetail ?? {},
        }
        return { ok: true, code: 0, json, stdout: JSON.stringify(json), stderr: '' }
      }
      // `create ...`: write a stub file at --out so existsSync assertions hold
      const outIndex = args.indexOf('--out')
      const out = outIndex >= 0 ? args[outIndex + 1] : undefined
      if (out) writeFileSync(out, 'stub')
      const json = {
        status: 'ok' as const,
        command: args[0] ?? 'create',
        summary: `created ${out ?? ''}`,
        ...(out ? { output_path: out } : {}),
        detail: {},
      }
      return { ok: true, code: 0, json, stdout: JSON.stringify(json), stderr: '' }
    },
  }

  return {
    runner,
    calls,
    last: () => calls.at(-1)?.args ?? [],
    lastInput: () => calls.at(-1)?.input,
    lastFrom: () => calls.at(-1)?.fromContent,
    lastOps: () => calls.at(-1)?.opsContent,
    lastData: () => calls.at(-1)?.dataContent,
    setReadItems: (items) => {
      readItems = items
    },
    setDocProperties: (properties) => {
      docProperties = properties
    },
    setMergeDetail: (detail) => {
      mergeDetail = detail
    },
    setCheckDetail: (detail) => {
      checkDetail = detail
    },
    setSheetLinksDetail: (detail) => {
      sheetLinksDetail = detail
    },
    failNext: (message) => {
      nextError = message
    },
  }
}
