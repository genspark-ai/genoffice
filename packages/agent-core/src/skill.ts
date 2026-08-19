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
  const owner = new Map<string, AgentSkill>()
  for (const skill of skills) {
    for (const tool of skill.tools) {
      if (owner.has(tool.name)) throw new Error(`duplicate tool name: ${tool.name}`)
      owner.set(tool.name, skill)
    }
  }
  return {
    id,
    systemPrompt: [intro, ...skills.map((s) => s.systemPrompt)].filter(Boolean).join('\n\n'),
    tools: skills.flatMap((s) => s.tools),
    buildContext: () =>
      skills
        .map((s) => s.buildContext?.() ?? '')
        .filter(Boolean)
        .join('\n\n'),
    executeTool: (call, signal) => {
      const skill = owner.get(call.name)
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
