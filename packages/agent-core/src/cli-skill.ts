import type { AgentToolDef } from './types'
import type { AgentSkill, ExecutedToolCall } from './skill'
import { kbCiteHref } from './knowledge-base-skill'

/**
 * The bundled `genoffice` command line as an agent skill: the same engine the
 * CLI, `genoffice mcp` and the app's own MCP server expose, driven here by the
 * in-app loop. The host owns the one dangerous operation (spawning the CLI):
 * it fixes the executable, confines reads and writes to the roots it hands
 * out, and stages the files the model writes into a per-conversation work
 * directory. This file only shapes the tools and the workflow prompt.
 */

export const RUN_GENOFFICE_TOOL = 'run_genoffice'
export const LOAD_GUIDE_TOOL = 'load_guide'

export const GUIDE_DOMAINS = ['docs', 'sheets', 'slides'] as const
export type GuideDomain = (typeof GUIDE_DOMAINS)[number]

/** one file the model writes before a command runs (relative to the work dir) */
export interface CliStagedFile {
  name: string
  content: string
}

export interface CliRunRequest {
  args: string[]
  files?: CliStagedFile[]
}

export interface CliRunResult {
  ok: boolean
  /** process exit code; -1 when the CLI never ran */
  code: number
  /** the CLI's `--json` payload when it printed one */
  json?: unknown
  stdout: string
  stderr: string
}

export interface CliContext {
  /** scratch directory for ops, specs and staged inputs; relative paths resolve here */
  workDir: string
  /** where finished documents go, the folder the home screen lists */
  outDir: string
}

export interface CliSkillDeps {
  /** null until the host has resolved the directories for this conversation */
  context(): CliContext | null
  run(request: CliRunRequest): Promise<CliRunResult>
}

export interface CliSkill extends AgentSkill {
  /** absolute paths the CLI reported as output_path during the session */
  createdPaths(): readonly string[]
}

const MAX_OUTPUT_CHARS = 24_000

const CLI_SYSTEM_PROMPT = `## Building documents with genoffice
run_genoffice runs the bundled genoffice command line (always with --json) and
returns its JSON result. Use it whenever the user wants a real file built from
their material: a Word report with headings and tables, a workbook with live
formulas and charts, a slide deck. Plain prose a user only wants saved can go
through create_document; anything with structure, formulas, several files as
sources, or a .xlsx / .pptx result goes through run_genoffice.
- Paths: the context names a work directory and an output directory. Write
  ops, specs, markdown and other inputs with the files argument (names relative
  to the work directory), reference them by that relative name in args, and put
  every finished document under the output directory with --out, using a short
  descriptive file name in the user's language. Source files are the knowledge
  base paths from the context; read them with "docs read", "sheet read" or
  "pdf read" (add --full or --range when the preview is not enough).
- A file inside the output directory may be edited in place ("apply" /
  "replace" without --out; the original is backed up first). A knowledge-base
  file anywhere else is read-only: edit it with --out to a new file under the
  output directory. --force is refused; pick a name that does not exist yet.
- "<command> --help" prints a command's flags when the guide does not.
- Before writing ops or a deck spec, call load_guide for that domain once
  (docs, sheets or slides; a group or op name narrows it). Follow the staged
  deck pipeline the slides design guide describes: style.md, outline.json,
  one page file at a time checked with "slides check", then "create --type
  pptx --spec pages --outline outline.json --out ... --audit". A deck page
  that needs a native table or chart is built with ops instead ("create
  --type pptx --ops" with addTable / addChart): a page spec has no table
  element, and a grid drawn from shapes is not a table.
- Word: "create --type docx --from report.md" builds a document from Markdown
  (headings, lists, tables, bold); then "docs apply --ops ... --out" for table
  styles, fields or page setup. Excel: "create --type xlsx --from table.json"
  then "sheet apply --ops" for formulas, formats, conditional rules and charts;
  write formulas as strings starting with "=".
- Read the JSON: status "ok" with output_path means the file exists; an error
  carries error, message and suggestion, so fix the input and resend the whole
  batch. After a build, "render <file> --out shots" gives one PNG per page when
  you need to check the layout.
- Report the result briefly and link each finished file as
  [file name](cite link), copying the cite link the tool returned exactly.
  Only claim a file was created when run_genoffice returned status "ok".`

// save-only phrasing, as in create-document-skill: "created"/"generated" also describe corpus facts
const CLAIMS_BUILT =
  /\bI(?:'ve| have)? (?:saved|exported|built|written) (?:it|this|that|the|a|an|your)\b|\b(?:has been|have been|is now) (?:saved|exported|built)\b|\bsaved (?:it |the (?:file|document|workbook|deck) )?(?:as|to|under)\b/i

function clip(text: string): string {
  return text.length > MAX_OUTPUT_CHARS
    ? `${text.slice(0, MAX_OUTPUT_CHARS)}\n…(+${text.length - MAX_OUTPUT_CHARS} chars)`
    : text
}

function stringList(value: unknown): string[] | null {
  if (!Array.isArray(value) || value.length === 0) return null
  const out: string[] = []
  for (const item of value) {
    if (typeof item !== 'string') return null
    out.push(item)
  }
  return out
}

function stagedFiles(value: unknown): CliStagedFile[] | null {
  if (value === undefined || value === null) return []
  if (!Array.isArray(value)) return null
  const out: CliStagedFile[] = []
  for (const item of value) {
    if (!item || typeof item !== 'object') return null
    const { name, content } = item as Record<string, unknown>
    if (typeof name !== 'string' || !name.trim() || typeof content !== 'string') return null
    out.push({ name, content })
  }
  return out
}

/** the CLI's own one-line summary, when the payload has one */
function summaryOf(json: unknown): string | undefined {
  if (!json || typeof json !== 'object') return undefined
  const { summary, message } = json as { summary?: unknown; message?: unknown }
  if (typeof summary === 'string') return summary
  if (typeof message === 'string') return message
  return undefined
}

function outputPathOf(json: unknown): string | undefined {
  if (!json || typeof json !== 'object') return undefined
  const { status, output_path: out } = json as { status?: unknown; output_path?: unknown }
  return status === 'ok' && typeof out === 'string' && out ? out : undefined
}

export function createCliSkill(deps: CliSkillDeps): CliSkill {
  const created: string[] = []

  const runTool: AgentToolDef = {
    name: RUN_GENOFFICE_TOOL,
    description:
      'Run the bundled genoffice command line and return its JSON result. args is the command ' +
      'and its arguments (--json is added), e.g. ["create","--type","docx","--from","report.md","--out","<output dir>/Report.docx"]. ' +
      'files stages inputs into the work directory first (relative names), so ops, specs and markdown ' +
      'never have to be quoted on the command line. Finished documents go under the output directory with --out.',
    inputSchema: {
      type: 'object',
      properties: {
        args: {
          type: 'array',
          items: { type: 'string' },
          description: 'the genoffice command and its arguments, one string each',
        },
        files: {
          type: 'array',
          description: 'files to write into the work directory before the command runs',
          items: {
            type: 'object',
            properties: {
              name: { type: 'string', description: 'relative path inside the work directory' },
              content: { type: 'string', description: 'the whole file as text' },
            },
            required: ['name', 'content'],
          },
        },
      },
      required: ['args'],
    },
  }

  const guideTool: AgentToolDef = {
    name: LOAD_GUIDE_TOOL,
    description:
      'The op reference and design guides for one domain (what "genoffice guide" prints). ' +
      'Call it before writing ops or a deck spec; topic narrows it to a group, one op, "design" or "spec" (slides).',
    inputSchema: {
      type: 'object',
      properties: {
        domain: { type: 'string', enum: [...GUIDE_DOMAINS] },
        topic: {
          type: 'string',
          description: 'optional: a group name, an op name, or for slides "design" / "spec"',
        },
      },
      required: ['domain'],
    },
  }

  const execute = async (request: CliRunRequest, summaryLabel: string, plain = false) => {
    const result = await deps.run(request)
    const out = outputPathOf(result.json)
    if (out) created.push(out)
    // a guide is prose: hand it over as text, not as one JSON string with escaped newlines
    const body =
      plain && result.ok
        ? (summaryOf(result.json) ?? result.stdout)
        : result.json
          ? JSON.stringify(result.json, null, 1)
          : result.stdout
    const lines = [`exit ${result.code}`]
    const summary = summaryOf(result.json)
    if (summary && !plain) lines.push(`summary: ${summary}`)
    if (out) lines.push(`output_path: ${out}`, `cite: ${kbCiteHref(out)}`)
    lines.push(clip(body))
    if (result.stderr.trim()) lines.push(`[stderr]\n${clip(result.stderr)}`)
    return {
      output: lines.join('\n'),
      isError: !result.ok,
      mutated: false,
      summary: summaryLabel,
    }
  }

  return {
    id: 'genoffice-cli',
    systemPrompt: CLI_SYSTEM_PROMPT,
    tools: [runTool, guideTool],
    createdPaths: () => created,
    buildContext: () => {
      const ctx = deps.context()
      if (!ctx) return ''
      return `genoffice directories:\n  work directory: ${ctx.workDir}\n  output directory: ${ctx.outDir}`
    },
    executeTool: async (call) => {
      if (call.name === LOAD_GUIDE_TOOL) {
        const domain = call.input.domain
        if (typeof domain !== 'string' || !(GUIDE_DOMAINS as readonly string[]).includes(domain)) {
          return {
            output: `domain must be one of ${GUIDE_DOMAINS.join(', ')}`,
            isError: true,
            summary: 'bad domain',
          }
        }
        const topic = typeof call.input.topic === 'string' ? call.input.topic.trim() : ''
        const args = ['guide', domain, ...(topic ? [topic] : [])]
        return execute({ args }, `guide ${domain}${topic ? ` ${topic}` : ''}`, true)
      }
      if (call.name !== RUN_GENOFFICE_TOOL) {
        return { output: `unknown tool: ${call.name}`, isError: true, summary: call.name }
      }
      const args = stringList(call.input.args)
      if (!args) {
        return {
          output: 'args must be a non-empty array of strings',
          isError: true,
          summary: 'bad args',
        }
      }
      const files = stagedFiles(call.input.files)
      if (!files) {
        return {
          output: 'files must be an array of { name, content } with string fields',
          isError: true,
          summary: 'bad files',
        }
      }
      try {
        return await execute({ args, files }, `genoffice ${args.slice(0, 3).join(' ')}`)
      } catch (error) {
        const message = error instanceof Error ? error.message : String(error)
        return {
          output: `genoffice could not run: ${message}`,
          isError: true,
          summary: 'cli failed',
        }
      }
    },
    verifyResponse: (finalText, executed: readonly ExecutedToolCall[]) => {
      const built = executed.some((c) => c.name === RUN_GENOFFICE_TOOL && c.ok)
      if (built || !CLAIMS_BUILT.test(finalText)) return null
      return (
        'The reply says a file was built or saved, but no run_genoffice call succeeded in this run. ' +
        'Say plainly that nothing was written, or build the file first.'
      )
    },
  }
}
