import { CREATE_DOCUMENT_TOOL, RUN_GENOFFICE_TOOL, type AgentToolCall } from '@genoffice/agent-core'

/**
 * The task card of one Ask AI answer: every file the turn set out to build,
 * tracked from the tool call that names it to the result that confirms it.
 * Derived from the loop's tool events alone, so the card needs nothing from
 * the model's prose; a path the CLI reports as output_path is the one a click
 * opens.
 */

export type ArtifactStatus = 'building' | 'done' | 'failed'

export interface Artifact {
  /** matches the tool call that started it */
  callId: string
  /** file name as shown; the absolute path once the result reports it */
  name: string
  path?: string
  status: ArtifactStatus
  /** the CLI's one-line reason when status is 'failed' */
  error?: string
}

const DOC_EXT = /\.(docx|xlsx|pptx|pdf|md|html|csv|png|jpe?g)$/i

function baseName(path: string): string {
  return path.split(/[\\/]/).pop() ?? path
}

/** the --out value of a run_genoffice call, or the title of a create_document call */
export function artifactTarget(call: AgentToolCall): string | null {
  if (call.name === CREATE_DOCUMENT_TOOL) {
    const title = typeof call.input.title === 'string' ? call.input.title.trim() : ''
    const format = call.input.format === 'md' ? 'md' : 'docx'
    return `${title || 'Untitled'}.${format}`
  }
  if (call.name !== RUN_GENOFFICE_TOOL || !Array.isArray(call.input.args)) return null
  const args = call.input.args as unknown[]
  const at = args.indexOf('--out')
  const out = at >= 0 ? args[at + 1] : undefined
  if (typeof out !== 'string' || !out) return null
  // render --out names a folder of PNGs, convert/create a file: only files are artifacts
  return DOC_EXT.test(out) ? out : null
}

/**
 * A new 'building' artifact for a call that names a target, else null. With
 * an output directory known, only files going there count: a relative or
 * work-directory --out is an intermediate (a base workbook, a staged spec),
 * not something the user asked for.
 */
export function artifactFromCall(call: AgentToolCall, outDir?: string | null): Artifact | null {
  const target = artifactTarget(call)
  if (!target) return null
  if (outDir && call.name === RUN_GENOFFICE_TOOL && !isInside(outDir, target)) return null
  return { callId: call.id, name: baseName(target), status: 'building' }
}

function isInside(dir: string, path: string): boolean {
  const root = dir.replace(/[\\/]+$/, '')
  return path.startsWith(`${root}/`) || path.startsWith(`${root}\\`)
}

/**
 * Settle the artifact of `call` from the tool's result text. The CLI skill
 * prints `output_path: <abs>` on success and `summary: <reason>` on failure;
 * create_document prints `path: <abs>`.
 */
export function settleArtifact(
  artifacts: readonly Artifact[],
  call: AgentToolCall,
  execution: { output: string; isError?: boolean },
): Artifact[] {
  const index = artifacts.findIndex((a) => a.callId === call.id)
  if (index < 0) return [...artifacts]
  const current = artifacts[index]!
  // create_document indents its `path:` line under "saved <name>"
  const path = /^\s*(?:output_path|path): (.+)$/m.exec(execution.output)?.[1]?.trim() ?? undefined
  const next: Artifact = execution.isError
    ? {
        ...current,
        status: 'failed',
        error:
          /^summary: (.+)$/m.exec(execution.output)?.[1]?.trim() ??
          execution.output.split('\n')[0]?.slice(0, 160),
      }
    : { ...current, status: 'done', ...(path ? { path, name: baseName(path) } : {}) }
  const out = [...artifacts]
  out[index] = next
  return out
}

/** the prompt a Retry button sends for one failed artifact */
export function retryInstruction(artifact: Artifact): string {
  return artifact.error
    ? `Build ${artifact.name} again; the previous attempt failed with: ${artifact.error}`
    : `Build ${artifact.name} again.`
}
