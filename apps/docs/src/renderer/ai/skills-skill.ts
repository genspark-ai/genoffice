import type { AgentSkill } from '@genoffice/agent-core'
import type { UserCliRun, UserSkillEntry } from '../../shared/ipc'
import { t } from '../i18n/locale'

/**
 * The skills a user installed as an AgentSkill, plus the CLI those skills teach.
 *
 * #1895 put the skills on disk (userData/skills); this is the half that reaches
 * the model. The main process owns every boundary — only installed skills are
 * listed, `read_skill` takes a name and the main process resolves the path, and
 * `run_cli` spawns the bundled genoffice CLI and nothing else (user-cli.ts) —
 * so this file is a caller, not a guard.
 *
 * buildContext is synchronous by contract, and the list lives across an IPC
 * call: the first turn of a session relies on the system prompt pointing at
 * list_skills, and every turn after one sees the cached list in context.
 */

const SKILLS_SYSTEM_PROMPT = `## User skills
The user may have installed skills — short manuals for document workflows this
app's command line can carry out. They are listed in the per-turn context when
known; if that section is absent, call list_skills before assuming there are
none.
- read_skill returns one skill's full text, by its exact name. Follow its steps
  with the document tools you already have; quote its constraints rather than
  paraphrasing them, and prefer its steps over guessing.
- run_cli executes the bundled genoffice command line (the same "genoffice …
  --json" contract the skills describe). Use it when a skill directs you to, or
  for batch work no single document call covers. It runs only this app's CLI —
  it is not a general shell — and long output is truncated.`

export function createSkillsSkill(): AgentSkill {
  let cache: UserSkillEntry[] | null = null

  const contextLine = (s: UserSkillEntry): string =>
    `- ${s.name} — ${s.description}${s.relevant ? '' : ' (not about document formats)'}`

  return {
    id: 'skills',
    systemPrompt: SKILLS_SYSTEM_PROMPT,
    tools: [
      {
        name: 'list_skills',
        description:
          'List the skills the user installed in this app. Cheap; call it before assuming none exist.',
        inputSchema: { type: 'object', properties: {} },
      },
      {
        name: 'read_skill',
        description:
          'Read the full text of one installed skill, by its exact name from list_skills.',
        inputSchema: {
          type: 'object',
          properties: {
            name: { type: 'string', description: 'the skill name, exactly as listed' },
          },
          required: ['name'],
        },
      },
      {
        name: 'run_cli',
        description:
          "Run the bundled genoffice command line (it always answers --json). Only this app's own CLI, not a general shell; the first argument must be a command such as convert, create, info.",
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
    buildContext: () => {
      if (!cache || cache.length === 0) return ''
      return `Installed skills (read one with read_skill; use run_cli when a skill directs you to):\n${cache
        .map(contextLine)
        .join('\n')}`
    },
    executeTool: async (call) => {
      try {
        if (call.name === 'list_skills') {
          cache = await window.desktop.skillsList()
          if (cache.length === 0) {
            return {
              output:
                'No skills are installed. The user can add some under Settings → Integrations.',
              mutated: false,
              summary: t('aiSumSkillsList'),
            }
          }
          return {
            output: `${cache.length} installed skill(s):\n${cache.map(contextLine).join('\n')}`,
            mutated: false,
            summary: t('aiSumSkillsList'),
          }
        }
        if (call.name === 'read_skill') {
          const name = call.input.name
          if (typeof name !== 'string' || name.trim() === '') {
            return {
              output: 'a skill name is required',
              isError: true,
              summary: t('aiSumSkillRead', { name: '?' }),
            }
          }
          const text = await window.desktop.skillText(name)
          return {
            output: text,
            mutated: false,
            summary: t('aiSumSkillRead', { name }),
          }
        }
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
          const summary =
            out.json &&
            typeof out.json === 'object' &&
            'summary' in (out.json as Record<string, unknown>)
              ? `\nsummary: ${(out.json as { summary?: string }).summary}`
              : ''
          return {
            output: `exit ${out.code}${summary}\n${out.stdout}${
              out.stderr ? `\n[stderr]\n${out.stderr}` : ''
            }`,
            isError: !out.ok,
            mutated: false,
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
