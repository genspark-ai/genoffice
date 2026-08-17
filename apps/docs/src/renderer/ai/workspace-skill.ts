import type { AgentSkill } from '@genoffice/agent-core'

/**
 * Workspace Q&A as an AgentSkill: the agent calls workspace_search to retrieve
 * relevant passages from the user's saved documents (semantic search over a
 * local Ollama embedding index built in the main process — files never leave
 * the machine).
 */

const WORKSPACE_SYSTEM_PROMPT = `## Workspace Q&A
The user's saved documents (the shared GenOffice folder: docs, sheets, slides, PDFs and Markdown) are indexed locally for semantic search.
- When the user asks about their own documents ("what did I write about X?", "find the budget spreadsheet", "summarize my notes on Y"), search the workspace with workspace_search first and answer from the retrieved passages.
- If the search returns nothing useful, say so plainly instead of inventing file contents.
- Cite the file name alongside any factual claims drawn from search results.`

export function createWorkspaceSkill(): AgentSkill {
  return {
    id: 'workspace',
    systemPrompt: WORKSPACE_SYSTEM_PROMPT,
    tools: [
      {
        name: 'workspace_search',
        description:
          'Semantic search over the user\'s saved GenOffice documents (docx/pdf/pptx/xlsx/markdown/txt, indexed locally with Ollama embeddings). Returns the most relevant passages with file names. Use when the request concerns the user\'s own documents.',
        inputSchema: {
          type: 'object',
          properties: {
            query: {
              type: 'string',
              description: 'natural-language search query describing the content to find',
            },
            k: {
              type: 'integer',
              description: 'maximum number of passages to return (default 5)',
            },
          },
          required: ['query'],
        },
      },
    ],
    buildContext: () => '',
    executeTool: async (call) => {
      if (call.name !== 'workspace_search') {
        return { output: `unknown tool: ${call.name}`, isError: true, summary: call.name }
      }
      const query = String(call.input.query ?? '').trim()
      if (!query) {
        return { output: 'provide a search query', isError: true, summary: 'workspace_search' }
      }
      const k = Math.min(Math.max(Number(call.input.k) || 5, 1), 10)
      const res = await window.desktop.workspaceSearch(query, k)
      if (!res.ok) {
        return {
          output:
            res.error ??
            'workspace search failed (is Ollama running with an embedding model installed?)',
          isError: true,
          summary: 'workspace_search',
        }
      }
      const results = res.results ?? []
      if (results.length === 0) {
        return {
          output: 'No matching passages found in the workspace index.',
          summary: 'workspace_search: 0 matches',
        }
      }
      const lines = results.map(
        (r, i) => `[${i + 1}] ${r.file}\n${r.snippet}`,
      )
      return {
        output: `Top ${results.length} matches:\n\n${lines.join('\n\n')}`,
        summary: `workspace_search: ${results.length} matches`,
      }
    },
  }
}
