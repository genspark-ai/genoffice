import type { AgentSkill, AgentToolDef } from '@genoffice/agent-core'
import { clipPlainText, SELECTION_PREVIEW_CHARS } from '../studio-text'

const SYSTEM_PROMPT = `You are GenOffice's Hangul (HWP) assistant. You read the currently open Hangul document and answer questions about it.

This release has no editing tools. Do not claim you changed the document, and do not invent document content. You may use web_search for facts outside the document.

# Tools
- get_document_text: full document as plain text (may be truncated for long files)
- get_selection: currently selected text, if any

Read with tools before answering. If the user asks you to edit, rewrite, replace, or insert text, explain that editing is not available yet and answer from the current text instead.`

const TOOLS: AgentToolDef[] = [
  {
    name: 'get_document_text',
    description: 'Return the Hangul document as plain text. Long documents may be truncated.',
    inputSchema: { type: 'object', properties: {} },
  },
  {
    name: 'get_selection',
    description: 'Return the current selection as plain text, or report that nothing is selected.',
    inputSchema: { type: 'object', properties: {} },
  },
]

export interface HangulSkillDeps {
  fileName(): string
  pageCount(): number
  currentPage(): number | null
  hasSelection(): boolean
  selectionPreview(): string | null
  getDocumentText(): Promise<string>
  getSelection(): Promise<string | null>
}

export function createHangulSkill(getDeps: () => HangulSkillDeps): AgentSkill {
  return {
    id: 'hangul',
    systemPrompt: SYSTEM_PROMPT,
    tools: TOOLS,
    buildContext: () => {
      const deps = getDeps()
      const parts = [`Hangul document: "${deps.fileName()}", ${deps.pageCount()} page(s).`]
      const page = deps.currentPage()
      if (page) parts.push(`Current page: ${page}.`)
      if (deps.hasSelection()) {
        const preview = deps.selectionPreview()?.trim()
        if (preview) {
          const clipped =
            preview.length > SELECTION_PREVIEW_CHARS
              ? `${preview.slice(0, SELECTION_PREVIEW_CHARS)}…`
              : preview
          parts.push(`The user has selected the following text:\n"""\n${clipped}\n"""`)
        } else {
          parts.push('The user has a selection, but the raw selected text is not available.')
        }
      } else {
        parts.push('No text is selected.')
      }
      parts.push('Editing tools are not available in this release.')
      return parts.join('\n')
    },
    executeTool: async (call) => {
      const deps = getDeps()
      if (call.name === 'get_document_text') {
        try {
          const text = await deps.getDocumentText()
          return {
            output: text || '(empty document)',
            summary: 'Read document text',
          }
        } catch (err) {
          return {
            output: err instanceof Error ? err.message : String(err),
            isError: true,
            summary: 'Read document text',
          }
        }
      }
      if (call.name === 'get_selection') {
        try {
          const text = await deps.getSelection()
          if (!text?.trim()) {
            return { output: '(no selection)', summary: 'Read selection' }
          }
          return { output: clipPlainText(text), summary: 'Read selection' }
        } catch (err) {
          return {
            output: err instanceof Error ? err.message : String(err),
            isError: true,
            summary: 'Read selection',
          }
        }
      }
      return { output: `Unknown tool: ${call.name}`, isError: true, summary: call.name }
    },
  }
}
