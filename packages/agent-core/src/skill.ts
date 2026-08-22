import type { AgentToolCall, AgentToolDef, ToolExecution } from './types'

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
  /** Optional fast-path variants for simple generation (blank/small docs). */
  fastSystemPrompt?: string
  fastTools?: AgentToolDef[]
  buildFastContext?(): string
  /** Return true when this run's instruction qualifies for the fast path. */
  isFastPath?(instruction: string): boolean
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
  // Fast owner also includes fastTools so executeTool can find insert_content in fast mode
  const fastOwner = new Map<string, AgentSkill>()
  for (const skill of skills) {
    const fTools = skill.fastTools ?? skill.tools
    for (const tool of fTools) {
      if (!fastOwner.has(tool.name)) fastOwner.set(tool.name, skill)
    }
  }
  const hasFast = skills.some((s) => s.fastSystemPrompt || s.fastTools || s.buildFastContext || s.isFastPath)
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
      const skill = owner.get(call.name) ?? fastOwner.get(call.name)
      if (!skill) {
        return { output: `Unknown tool: ${call.name}`, isError: true, summary: call.name }
      }
      return skill.executeTool(call, signal)
    },
    ...(hasFast
      ? {
          fastSystemPrompt: [intro, ...skills.map((s) => s.fastSystemPrompt ?? s.systemPrompt)]
            .filter(Boolean)
            .join('\n\n'),
          fastTools: skills.flatMap((s) => s.fastTools ?? s.tools),
          buildFastContext: () =>
            skills
              .map((s) => (s.buildFastContext ? s.buildFastContext() : s.buildContext?.() ?? ''))
              .filter(Boolean)
              .join('\n\n'),
          isFastPath: (instruction: string) => skills.some((s) => s.isFastPath?.(instruction) ?? false),
        }
      : {}),
  }
}
