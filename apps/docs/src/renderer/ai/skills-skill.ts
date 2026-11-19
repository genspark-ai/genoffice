import type { AgentSkill } from '@genoffice/agent-core'
import type { UserCliRun } from '../../shared/ipc'
import { t } from '../i18n/locale'

/**
 * The command line the user's skills teach, as an AgentSkill.
 *
 * Skills themselves are not offered to the model: no listing, no lookup. The
 * only way a skill's text reaches a turn is the user attaching it from the
 * composer "+" menu, and then it rides in front of the request (see
 * withAttachedSkills). What stays here is `run_cli`, so an attached skill
 * that says "run genoffice convert …" can be followed. The main process owns
 * the boundary — it spawns the bundled genoffice CLI and nothing else
 * (user-cli.ts) — so this file is a caller, not a guard.
 */

const SKILLS_SYSTEM_PROMPT = `## Attached skills and the command line
The user may attach a skill to a message: its full text then appears at the
top of that message inside <skill> tags. Follow its steps with the document
tools you already have; quote its constraints rather than paraphrasing them,
and prefer its steps over guessing. Never assume a skill the user did not
attach.
- run_cli executes the bundled genoffice command line (the same "genoffice …
  --json" contract skills describe). Use it when an attached skill directs you
  to, or for batch work no single document call covers. It runs only this
  app's CLI — it is not a general shell — and long output is truncated.`

export function createSkillsSkill(): AgentSkill {
  return {
    id: 'skills',
    systemPrompt: SKILLS_SYSTEM_PROMPT,
    tools: [
      {
        name: 'run_cli',
        description:
          "Run the bundled genoffice command line (it always answers --json). Only this app's own CLI, not a general shell: document commands such as convert, create, info, docs, sheet, slides, pdf; no --force, and apply/replace must take --out. Relative paths resolve in the open document's folder; reads and writes stay inside the allowed folders the result names.",
        inputSchema: {
          type: 'object',
          properties: {
            args: {
              type: 'array',
              items: { type: 'string' },
              description: 'the command and its arguments, e.g. ["convert","in.docx","--to","pdf"]',
            },
          },
          required: ['args'],
        },
      },
    ],
    executeTool: async (call) => {
      try {
        if (call.name === 'run_cli') {
          const args = call.input.args
          if (
            !Array.isArray(args) ||
            args.length === 0 ||
            args.some((a) => typeof a !== 'string')
          ) {
            return {
              output: 'args must be a non-empty array of strings',
              isError: true,
              summary: t('aiSumRunCli', { args: '?' }),
            }
          }
          const out: UserCliRun = await window.desktop.runCli(args as string[])
          const json = (out.json ?? {}) as { summary?: string; output_path?: string }
          const summary = typeof json.summary === 'string' ? `\nsummary: ${json.summary}` : ''
          const where = [
            out.cwd ? `cwd: ${out.cwd}` : '',
            out.allowedRoots?.length ? `allowed folders: ${out.allowedRoots.join(', ')}` : '',
          ]
            .filter(Boolean)
            .join('\n')
          return {
            output: `exit ${out.code}${summary}\n${out.stdout}${
              out.stderr ? `\n[stderr]\n${out.stderr}` : ''
            }${where ? `\n${where}` : ''}`,
            isError: !out.ok,
            // a CLI run that wrote a file changed the world, not this editor's
            // document: count it as a mutation so the loop does not treat the
            // repeat as a redundant read, and so the final answer reports it
            mutated: out.ok && typeof json.output_path === 'string',
            summary: t('aiSumRunCli', { args: (args as string[]).join(' ') }),
          }
        }
        return { output: `unknown tool: ${call.name}`, isError: true, summary: call.name }
      } catch (err) {
        // in a standalone docs window (not the full app) the main process has
        // no handler for these channels — say that instead of leaking a raw
        // IPC error into the transcript
        return {
          output: err instanceof Error ? err.message : String(err),
          isError: true,
          summary: call.name,
        }
      }
    },
  }
}
