import { mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { cliErrorMessage, type CliRunOutcome, type CliRunner } from './cli-runner'

/**
 * Thin helpers that turn an MCP tool's arguments into a `genoffice` CLI
 * invocation. The headless MCP tools are a client of the CLI, not a second
 * implementation of the engines: input that arrives inline (markdown, a row
 * matrix, an ops array) is staged to a temp file, the CLI writes the output,
 * and its `--json` payload is mapped back to a tool result.
 */

/** run `fn` with a scratch directory that is always removed afterwards */
async function withTempDir<T>(fn: (dir: string) => Promise<T>): Promise<T> {
  const dir = await mkdtemp(join(tmpdir(), 'genoffice-mcp-'))
  try {
    return await fn(dir)
  } finally {
    await rm(dir, { recursive: true, force: true })
  }
}

/** stage inline content in the scratch dir and return its path */
async function stage(dir: string, name: string, content: string): Promise<string> {
  const path = join(dir, name)
  await writeFile(path, content, 'utf8')
  return path
}

/**
 * `create --type <type> --from <staged input> --out <target>`.
 *
 * The caller has already resolved/validated `out`; `overwrite` maps to
 * `--force`, so the CLI's own clobber guard only fires when it should.
 */
export async function createFileViaCli(
  runner: CliRunner,
  opts: {
    type: 'docx' | 'pptx' | 'xlsx'
    /** staged input file name + contents (e.g. `report.md`, a table JSON) */
    input: { name: string; content: string }
    out: string
    overwrite?: boolean
  },
): Promise<{ outcome: CliRunOutcome; summary: string; outputPath: string }> {
  return withTempDir(async (dir) => {
    const inputPath = await stage(dir, opts.input.name, opts.input.content)
    const args = ['create', '--type', opts.type, '--from', inputPath]
    return runCreate(runner, args, opts.out, opts.overwrite)
  })
}

async function runCreate(
  runner: CliRunner,
  baseArgs: string[],
  out: string,
  overwrite: boolean | undefined,
  runOptions?: { input: string },
): Promise<{ outcome: CliRunOutcome; summary: string; outputPath: string }> {
  const args = [...baseArgs, '--out', out, ...(overwrite ? ['--force'] : [])]
  const outcome = await runner.run(args, runOptions)
  if (!outcome.ok || !outcome.json || outcome.json.status !== 'ok') {
    throw new Error(cliErrorMessage(outcome))
  }
  return { outcome, summary: outcome.json.summary, outputPath: outcome.json.output_path ?? out }
}

/**
 * Read a .docx's visible text through `docs read`. Returns one paragraph per
 * line, matching the tool's documented shape.
 *
 * `--full` is required: without it the CLI clips each block to a 200-character
 * preview, so long paragraphs would come back silently truncated.
 */
export async function readDocxTextViaCli(runner: CliRunner, path: string): Promise<string> {
  const outcome = await runner.run(['docs', 'read', path, '--full'])
  if (!outcome.ok || !outcome.json || outcome.json.status !== 'ok') {
    throw new Error(cliErrorMessage(outcome))
  }
  const items = (outcome.json.detail?.items ?? []) as Array<{ text?: string }>
  return items
    .map((item) => String(item.text ?? ''))
    .join('\n')
    .replace(/\n+$/, '')
}
