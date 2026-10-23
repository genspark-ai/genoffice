import type { AgentSkill } from './skill'
import { formatKbSize, kbCiteHref } from './knowledge-base-skill'

/**
 * Chat attachments as an AgentSkill: the per-turn context lists the files the
 * user attached, and read_attachment pages through their extracted text. The
 * parsing lives in the main process; deps inject it, so the skill is shared by
 * the editors' panels and the home screen's Ask AI without an electron import.
 */

export const ATTACHMENT_IMAGE_EXTS: ReadonlySet<string> = new Set([
  'png',
  'jpg',
  'jpeg',
  'gif',
  'webp',
])

export const READ_ATTACHMENT_TOOL = 'read_attachment'
export const ATTACHMENT_READ_CHUNK_CHARS = 24_000

export interface AttachmentInfo {
  /** absolute local path; the file never leaves the machine */
  path: string
  name: string
  /** lowercased extension without the dot */
  ext: string
  sizeBytes: number
}

export interface AttachmentPage {
  text: string
  totalChars: number
  offset: number
}

export interface AttachmentsSkillDeps {
  list(): readonly AttachmentInfo[]
  /** one slice of the parsed text; rejects with the user-facing reason */
  read(path: string, offset: number, maxChars: number): Promise<AttachmentPage>
  /** list absolute paths and cite links too (home: the CLI tools take paths, answers cite files) */
  pathsInContext?: boolean
  /** tool-row labels shown in the panel */
  labels?: {
    read?: (name: string) => string
    image?: (name: string) => string
    invalid?: string
  }
}

const ATTACHMENTS_SYSTEM_PROMPT = `## Attachments
The user may attach local files to the conversation (see the "attachment list" in each turn's context).
- When the user's request involves attachment content, read it with read_attachment first, then answer or write; do not guess content from file names.
- Long files are read in pages: the result reports the total character count and the current range; to continue, set offset to the end position of the previous slice.
- Image attachments (png/jpg/gif/webp) are already sent as images with the user message — just look at them; read_attachment is only for text-like attachments.
- Do not call read_attachment when there are no attachments or they are unrelated to the request.`

const PATHS_PROMPT = `
- Attachments are listed with their absolute paths: tools that take a file path (converting, building from a template) may use them directly, and a statement drawn from an attachment is cited with its cite link like a knowledge-base file.`

export function createAttachmentsSkill(deps: AttachmentsSkillDeps): AgentSkill {
  const labels = deps.labels ?? {}
  return {
    id: 'files',
    systemPrompt: deps.pathsInContext
      ? ATTACHMENTS_SYSTEM_PROMPT + PATHS_PROMPT
      : ATTACHMENTS_SYSTEM_PROMPT,
    tools: [
      {
        name: READ_ATTACHMENT_TOOL,
        description:
          'Read the text content of an attachment (parsed locally). Long files are paged: read offset=0 first, then decide whether to continue based on the returned total character count.',
        inputSchema: {
          type: 'object',
          properties: {
            index: {
              type: 'integer',
              description: 'attachment index (0-based, see the attachment list)',
            },
            offset: { type: 'integer', description: 'start character position, default 0' },
          },
          required: ['index'],
        },
      },
    ],
    buildContext: () => {
      const list = deps.list()
      if (list.length === 0) return ''
      const lines = list.map(
        (a, i) => `${i} | ${a.name} | .${a.ext} | ${formatKbSize(a.sizeBytes)}`,
      )
      const head = `Attachment list (index | file name | type | size):\n${lines.join('\n')}`
      if (!deps.pathsInContext) return head
      const paths = list.map((a, i) => `${i} = ${a.path}\n    cite: ${kbCiteHref(a.path)}`)
      return `${head}\n\nAbsolute paths and cite links:\n${paths.join('\n')}`
    },
    executeTool: async (call) => {
      if (call.name !== READ_ATTACHMENT_TOOL) {
        return { output: `unknown tool: ${call.name}`, isError: true, summary: call.name }
      }
      const list = deps.list()
      const index = Number(call.input.index)
      const att = Number.isInteger(index) ? list[index] : undefined
      if (!att) {
        return {
          output: 'invalid attachment index (see the attachment list)',
          isError: true,
          summary: labels.invalid ?? 'read attachment',
        }
      }
      if (ATTACHMENT_IMAGE_EXTS.has(att.ext)) {
        return {
          output: `${att.name} is an image attachment already sent as an image with the user message; just look at the image in the message, no text to read.`,
          mutated: false,
          summary: labels.image?.(att.name) ?? att.name,
        }
      }
      const offset = Math.max(0, Number(call.input.offset) || 0)
      let page: AttachmentPage
      try {
        page = await deps.read(att.path, offset, ATTACHMENT_READ_CHUNK_CHARS)
      } catch (error) {
        return {
          output: error instanceof Error ? error.message : String(error),
          isError: true,
          summary: labels.read?.(att.name) ?? att.name,
        }
      }
      const end = page.offset + page.text.length
      const header = `File ${att.name}, total characters ${page.totalChars}, this slice ${page.offset}-${end}${
        end < page.totalChars
          ? ' (not finished, continue with offset=' + end + ')'
          : ' (end of file)'
      }`
      return {
        output: `${header}\n---\n${page.text}`,
        mutated: false,
        summary: labels.read?.(att.name) ?? att.name,
      }
    },
  }
}
