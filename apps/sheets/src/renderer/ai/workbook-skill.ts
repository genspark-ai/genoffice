import type { AgentSkill, MemoryStoreAdapter } from '@genoffice/agent-core'
import { createMemoryTools } from '@genoffice/agent-core'
import basePrompt from './prompts/base.md?raw'
import {
  WORKBOOK_TOOLS,
  buildWorkbookContext,
  executeWorkbookTool,
  type SheetsSkillDeps,
} from './tools'

/**
 * The workbook DSL as an AgentSkill: mirrors createDocsSkill's shape
 * (systemPrompt + tools + buildContext + executeTool) so it plugs into the
 * same packages/agent-core AgentLoop docx uses.
 *
 * Prompt layout: the always-loaded base prompt (prompts/base.md) stays small
 * — workflow, op catalog, cross-cutting discipline — while per-domain field
 * definitions and conventions live in prompts/guides/*.md, loaded on demand
 * via load_guide.
 */
export function createWorkbookSkill(deps: SheetsSkillDeps, memory?: MemoryStoreAdapter): AgentSkill {
  const memoryTools = memory ? createMemoryTools(memory) : null
  return {
    id: 'sheets',
    systemPrompt: basePrompt,
    tools: [...WORKBOOK_TOOLS, ...(memoryTools?.tools ?? [])],
    buildContext: () => {
      const base = buildWorkbookContext(deps)
      const memorySection = memoryTools?.contextSection()
      return memorySection ? `${base}\n\n${memorySection}` : base
    },
    executeTool: (call) => {
      if (memoryTools && (call.name === 'remember_memory' || call.name === 'forget_memory')) {
        return memoryTools.execute(call)
      }
      return executeWorkbookTool(call, deps)
    },
  }
}
