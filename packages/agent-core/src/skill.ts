import type { AgentToolCall, AgentToolDef, ToolExecution } from './types'

/** One tool call actually executed during a run, as seen by verifyResponse. */
export interface ExecutedToolCall {
  name: string
  /** false when the execution returned an error result */
  ok: boolean
}

/**
 * A skill packages one capability domain for the agent loop: its system
 * prompt section, its tools, per-turn context, and the tool executor.
 * AI Docs ships a docx skill; Excel / PPT skills plug in the same way.
 */
export interface AgentSkill {
  id: string
  /** system prompt section describing this skill's rules and tools */
  systemPrompt: string
  tools: AgentToolDef[]
  /**
   * Fresh context sections attached to every user turn (e.g. document
   * skeleton + selection). Return '' when there is nothing to attach.
   */
  buildContext?(): string
  /**
   * signal: aborted when the user hits stop. Long-running tools (e.g.
   * generate_deck with internal LLM calls) should check signal.aborted in
   * their loops and stop promptly.
   */
  executeTool(call: AgentToolCall, signal?: AbortSignal): ToolExecution | Promise<ToolExecution>
  /**
   * Claimed-action guard: inspect the run's final assistant text against the
   * tools that actually executed during the run. Return a corrective
   * instruction to force one more model turn (e.g. the text claims "I
   * selected/located ..." but no matching tool call succeeded), or null to
   * accept the reply. The loop applies the correction at most once per run,
   * so a detector false-positive costs one extra turn and cannot loop.
   */
  verifyResponse?(finalText: string, executed: readonly ExecutedToolCall[]): string | null
}

/**
 * Merge several skills into one (tool names must be globally unique).
 * `intro` becomes the shared preamble of the combined system prompt.
 */
export function composeSkills(id: string, intro: string, skills: AgentSkill[]): AgentSkill {
  // Recomputed per access: a sub-skill may expose `tools` through a getter
  // keyed on runtime capability (e.g. gsk login/toggle), and the loop reads
  // the composed skill's tools before every model request.
  const ownerOf = (name: string): AgentSkill | undefined =>
    skills.find((skill) => skill.tools.some((tool) => tool.name === name))
  return {
    id,
    // live like tools: a sub-skill's prompt may vary with the same capability its tools key on
    get systemPrompt() {
      return [intro, ...skills.map((s) => s.systemPrompt)].filter(Boolean).join('\n\n')
    },
    get tools() {
      const all = skills.flatMap((s) => s.tools)
      const seen = new Set<string>()
      for (const tool of all) {
        if (seen.has(tool.name)) throw new Error(`duplicate tool name: ${tool.name}`)
        seen.add(tool.name)
      }
      return all
    },
    buildContext: () =>
      skills
        .map((s) => s.buildContext?.() ?? '')
        .filter(Boolean)
        .join('\n\n'),
    executeTool: (call, signal) => {
      const skill = ownerOf(call.name)
      if (!skill) {
        return { output: `Unknown tool: ${call.name}`, isError: true, summary: call.name }
      }
      return skill.executeTool(call, signal)
    },
    verifyResponse: (finalText, executed) => {
      for (const skill of skills) {
        const correction = skill.verifyResponse?.(finalText, executed)
        if (correction) return correction
      }
      return null
    },
  }
}
