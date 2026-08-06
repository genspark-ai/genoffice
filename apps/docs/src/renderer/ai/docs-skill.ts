import type { Editor } from '@tiptap/core'
import type { AgentSkill, MemoryStoreAdapter } from '@genoffice/agent-core'
import { createMemoryTools } from '@genoffice/agent-core'
import { AGENT_SYSTEM_PROMPT, buildDocContext, type AiTrack, type NumIds } from './protocol'
import { AGENT_TOOLS, executeTool, markDocSeen } from './tools'

/**
 * The docx capability as an AgentSkill: document skeleton context, the five
 * document tools, and the local executor. Future apps register their own
 * skills (Excel / PPT) against the same agent loop.
 */
export function createDocsSkill(
  getEditor: () => Editor,
  getNumIds: () => NumIds,
  getTrack?: () => AiTrack | undefined,
  memory?: MemoryStoreAdapter,
): AgentSkill {
  const memoryTools = memory ? createMemoryTools(memory) : null
  return {
    id: 'docx',
    systemPrompt: AGENT_SYSTEM_PROMPT,
    tools: [...AGENT_TOOLS, ...(memoryTools?.tools ?? [])],
    buildContext: () => {
      const editor = getEditor()
      markDocSeen(editor) // the context the model receives is the freshness baseline for index-addressed writes
      const base = buildDocContext(editor)
      const memorySection = memoryTools?.contextSection()
      return memorySection ? `${base}\n\n${memorySection}` : base
    },
    executeTool: (call, signal) => {
      if (memoryTools && (call.name === 'remember_memory' || call.name === 'forget_memory')) {
        return memoryTools.execute(call)
      }
      return executeTool(getEditor(), call, getNumIds(), getTrack?.(), signal)
    },
  }
}
