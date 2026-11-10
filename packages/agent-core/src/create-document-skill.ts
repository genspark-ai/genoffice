import type { AgentToolDef } from './types'
import type { AgentSkill, ExecutedToolCall } from './skill'
import { kbCiteHref } from './knowledge-base-skill'

/**
 * Lets a chat that has no open document write its result to disk: the model
 * drafts the content as markdown and create_document lands it as a .md or
 * .docx in the user's save folder. The skill is storage-agnostic — the host
 * injects the one write operation — so it pairs with the knowledge-base skill
 * on the home screen today and can back a "save as new file" in the editor
 * panels later.
 *
 * The created path is handed back as a filenav cite link (the same primitive
 * the knowledge-base citations use) so the answer can offer a click-to-open
 * link; createdPaths() lets a sibling skill's citation guard accept it.
 */

export type CreateDocumentFormat = 'md' | 'docx'

export interface CreateDocumentRequest {
  title: string
  format: CreateDocumentFormat
  /** the whole document as markdown */
  content: string
}

export interface CreateDocumentDeps {
  create(request: CreateDocumentRequest): Promise<{ path: string }>
  /** sibling tools whose success also makes "I saved it" a true statement */
  alsoSaves?: readonly string[]
}

export interface CreateDocumentSkill extends AgentSkill {
  /** absolute paths this skill wrote during the session */
  createdPaths(): readonly string[]
}

export const CREATE_DOCUMENT_TOOL = 'create_document'

/** keep in step with the main-process gate, which refuses larger bodies */
export const CREATE_DOCUMENT_MAX_CHARS = 2_000_000

const CREATE_SYSTEM_PROMPT = `## Creating documents
When the user asks for a file — a report, summary, memo, notes, "save this",
"export this", "make a document / markdown / Word file" — write the complete
content and call create_document; do not paste the document into the chat and
do not describe what the file would contain.
- Write the content in Markdown (headings, lists, tables, bold). format "docx"
  gives a Word document, "md" a Markdown file; pick what the user asked for,
  default to "docx" when they only said "document" and "md" when they said
  Markdown.
- Keep the knowledge-base citations inside the document as plain text (file
  name and folder), not as cite links.
- The tool returns the saved path with a cite link: reply briefly and link the
  file as [file name](cite link), copying the link exactly.
- Only claim a file was saved when create_document returned ok.`

// save-only phrasing: "created" / "generated" also describe corpus facts
// ("the file was created in 2019"), so they are not claims here
const CLAIMS_SAVED =
  /\bI(?:'ve| have)? (?:saved|exported) (?:it|this|that|the|a|an|your)\b|\b(?:has been|have been|is now) saved\b|\bsaved (?:it |the (?:file|document) )?(?:as|to|in) \S+\.(?:md|docx)\b/i

export function createDocumentSkill(deps: CreateDocumentDeps): CreateDocumentSkill {
  const created: string[] = []
  const tool: AgentToolDef = {
    name: CREATE_DOCUMENT_TOOL,
    description:
      'Save a new document in the user’s save folder from Markdown content. ' +
      'format "docx" writes a Word document, "md" a Markdown file. ' +
      'Returns the saved path and its cite link; link it in the reply.',
    inputSchema: {
      type: 'object',
      properties: {
        title: { type: 'string', description: 'file name without extension' },
        format: { type: 'string', enum: ['md', 'docx'], description: 'output format' },
        content: { type: 'string', description: 'the complete document as Markdown' },
      },
      required: ['title', 'format', 'content'],
    },
  }
  return {
    id: 'create-document',
    systemPrompt: CREATE_SYSTEM_PROMPT,
    tools: [tool],
    createdPaths: () => created,
    executeTool: async (call) => {
      if (call.name !== CREATE_DOCUMENT_TOOL) {
        return { output: `unknown tool: ${call.name}`, isError: true, summary: call.name }
      }
      const title = typeof call.input.title === 'string' ? call.input.title.trim() : ''
      const format = call.input.format
      const content = typeof call.input.content === 'string' ? call.input.content : ''
      if (format !== 'md' && format !== 'docx') {
        return { output: 'format must be "md" or "docx"', isError: true, summary: 'bad format' }
      }
      if (!content.trim()) {
        return { output: 'content must not be empty', isError: true, summary: 'empty content' }
      }
      if (content.length > CREATE_DOCUMENT_MAX_CHARS) {
        return {
          output: `content is too long (${content.length} characters, cap ${CREATE_DOCUMENT_MAX_CHARS}); split it into several documents`,
          isError: true,
          summary: 'content too long',
        }
      }
      try {
        const { path } = await deps.create({ title: title || 'Untitled', format, content })
        created.push(path)
        const name = path.split(/[\\/]/).pop() ?? path
        return {
          output: `saved ${name}\n  path: ${path}\n  cite: ${kbCiteHref(path)}`,
          summary: name,
        }
      } catch (error) {
        const message = error instanceof Error ? error.message : String(error)
        return {
          output: `could not save the document: ${message}`,
          isError: true,
          summary: 'save failed',
        }
      }
    },
    verifyResponse: (finalText, executed: readonly ExecutedToolCall[]) => {
      const savers = [CREATE_DOCUMENT_TOOL, ...(deps.alsoSaves ?? [])]
      const saved = executed.some((c) => savers.includes(c.name) && c.ok)
      if (saved || !CLAIMS_SAVED.test(finalText)) return null
      return (
        'The reply says a file was saved, but create_document did not succeed in this run. ' +
        'Say plainly that nothing was saved; call create_document only if the user asked for a file.'
      )
    },
  }
}
