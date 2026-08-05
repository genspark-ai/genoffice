/**
 * Generic project-memory agent tools (remember / forget) plus the context
 * section that surfaces persisted entries to the model.
 *
 * The adapter is app-provided: each app wraps its resolved project-store
 * ProjectApi (projectId + entries ref + IPC calls) in a MemoryStoreAdapter.
 * This module stays dependency-free so agent-core can be used standalone.
 */
import type { AgentToolCall, AgentToolDef, ToolExecution } from './types.js'

export interface MemoryEntryLike {
  id: string
  text: string
  ts: string
}

/** Project-memory adapter implemented by each app over its project-store API. */
export interface MemoryStoreAdapter {
  /** Whether memory can be read/written right now (project resolved + API present). */
  available(): boolean
  /** Current memory entries (newest first). */
  entries(): MemoryEntryLike[]
  /** Persist a new entry; returns the stored entry (with id/ts). */
  add(text: string): Promise<MemoryEntryLike>
  /** Delete an entry by id (no-op when unknown). */
  remove(id: string): Promise<void>
}

export interface MemoryTools {
  tools: AgentToolDef[]
  execute(call: AgentToolCall): ToolExecution | Promise<ToolExecution>
  /** Markdown-ish context section listing persisted entries ('' when unavailable). */
  contextSection(): string
}

const MAX_ENTRY_CHARS = 2_000

export function createMemoryTools(adapter: MemoryStoreAdapter): MemoryTools {
  const tools: AgentToolDef[] = [
    {
      name: 'remember_memory',
      description:
        'Persist a fact, preference or decision about this project so it is remembered across sessions and files (e.g. "the client prefers US date formats", "quarterly reports use 12pt font"). Save only durable, reusable facts — not one-off task details. Overwrite outdated facts by forgetting the old entry first.',
      inputSchema: {
        type: 'object',
        properties: {
          text: { type: 'string', description: 'the fact/preference to remember (one sentence)' },
        },
        required: ['text'],
      },
    },
    {
      name: 'forget_memory',
      description:
        'Delete a previously persisted project memory entry by its id (ids are listed in the project memory context section).',
      inputSchema: {
        type: 'object',
        properties: {
          id: { type: 'string', description: 'memory entry id to delete' },
        },
        required: ['id'],
      },
    },
  ]

  const execute = (call: AgentToolCall): ToolExecution | Promise<ToolExecution> => {
    if (!adapter.available()) {
      return {
        output: 'Project memory is unavailable (no resolved project).',
        isError: true,
        summary: 'project memory',
      }
    }
    if (call.name === 'remember_memory') {
      const text = typeof call.input.text === 'string' ? call.input.text.trim() : ''
      if (!text) {
        return { output: 'No text provided.', isError: true, summary: 'project memory' }
      }
      return adapter.add(text.slice(0, MAX_ENTRY_CHARS)).then((entry) => ({
        output: `Saved to project memory (id ${entry.id}).`,
        mutated: true,
        summary: 'project memory: remember',
      }))
    }
    if (call.name === 'forget_memory') {
      const id = typeof call.input.id === 'string' ? call.input.id : ''
      if (!id) {
        return { output: 'No id provided.', isError: true, summary: 'project memory' }
      }
      return adapter.remove(id).then(() => ({
        output: `Deleted project memory entry ${id} (or it did not exist).`,
        mutated: true,
        summary: 'project memory: forget',
      }))
    }
    return {
      output: `Unknown memory tool: ${call.name}`,
      isError: true,
      summary: 'project memory',
    }
  }

  const contextSection = (): string => {
    if (!adapter.available()) return ''
    const entries = adapter.entries()
    const list =
      entries.length === 0
        ? '(empty)'
        : entries.map((e) => `- ${e.text} (id: ${e.id})`).join('\n')
    return [
      'Project memory (persisted facts to honor across this project):',
      list,
      'Use remember_memory to persist a new fact and forget_memory to delete one by id.',
    ].join('\n')
  }

  return { tools, execute, contextSection }
}
