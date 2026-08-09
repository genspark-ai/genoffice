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
}

/**
 * Merge several skills into one (tool names must be globally unique).
 * `intro` becomes the shared preamble of the combined system prompt.
 */
export function composeSkills(id: string, intro: string, skills: AgentSkill[]): AgentSkill {
  /**
   * Resolved on read, not at compose time: a skill may expose a tool list that
   * changes during the session (the web skill only offers `load_skill` once the
   * user has one). The loop re-reads `skill.tools` per request, so keeping this
   * lazy is what lets a newly added skill be callable without remounting the
   * panel — and it keeps the routing map from going stale against it.
   */
  const currentOwners = (): Map<string, AgentSkill> => {
    const owner = new Map<string, AgentSkill>()
    for (const skill of skills) {
      for (const tool of skill.tools) {
        if (owner.has(tool.name)) throw new Error(`duplicate tool name: ${tool.name}`)
        owner.set(tool.name, skill)
      }
    }
    return owner
  }
  // fail fast on a duplicate that is present from the start
  currentOwners()
  return {
    id,
    systemPrompt: [intro, ...skills.map((s) => s.systemPrompt)].filter(Boolean).join('\n\n'),
    get tools() {
      return skills.flatMap((s) => s.tools)
    },
    buildContext: () =>
      skills
        .map((s) => s.buildContext?.() ?? '')
        .filter(Boolean)
        .join('\n\n'),
    executeTool: (call, signal) => {
      const skill = currentOwners().get(call.name)
      if (!skill) {
        return { output: `Unknown tool: ${call.name}`, isError: true, summary: call.name }
      }
      return skill.executeTool(call, signal)
    },
  }
}
