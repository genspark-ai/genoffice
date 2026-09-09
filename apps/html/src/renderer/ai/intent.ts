import type { AgentSkill } from '@genoffice/agent-core'

/** what the user picked on the blank-document card: a designed page or written content */
export type PageIntent = 'design' | 'write'

const WRITE_DIRECTIVE = [
  'User intent for this blank document: WRITE CONTENT (content first, not a designed page).',
  '- Do not call ask_clarification or plan_page unless the user explicitly asks for design choices.',
  '- Call write_document with the title, a plan (outline, key points, length, voice) and the reference material as context; the system writes the whole document as a clean single-column page.',
  '- Put the substance in the plan and context; the page is the document, not a showcase for it.',
].join('\n')

/** the intent only steers the first generation: once the page has content the default rules apply again */
export function createIntentSkill(getIntent: () => PageIntent, isBlank: () => boolean): AgentSkill {
  return {
    id: 'intent',
    systemPrompt: '',
    tools: [],
    buildContext: () => (getIntent() === 'write' && isBlank() ? WRITE_DIRECTIVE : ''),
    executeTool: async (call) => ({
      output: `unknown tool: ${call.name}`,
      isError: true,
      summary: call.name,
    }),
  }
}
