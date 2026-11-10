import type { AgentSkill } from '@genoffice/agent-core'
import type { AttachmentMeta } from '../../shared/ipc'
import { ATTACHMENT_IMAGE_EXTS } from '../../shared/ipc'
import { t } from '../i18n/locale'

/**
 * Chat-attachment capability as an AgentSkill (isomorphic to apps/docs files-skill):
 * each turn's context lists the attached local files; read_attachment reads their
 * extracted text in pages (parsing happens in the main process; files never leave
 * this machine).
 */

const READ_CHUNK_CHARS = 24_000

const FILES_SYSTEM_PROMPT = `## Attachments
The user may attach local files to the conversation (see the "attachment list" in each turn's context).
- When the user's request involves attachment content, read it with read_attachment first, then answer or generate; don't guess content from file names.
- Long files are read in pages: the result gives the total character count and the current range; to continue, set offset to the previous chunk's end position.
- Image attachments (png/jpg/gif/webp) were already sent as images with the user message — just look at them; read_attachment is only for text attachments. To PLACE an attached image on a slide, call insert_web_image (or replace_image) with url=attachment://<file name> so the original file is embedded as-is — never recreate it with generate_image.
- When there are no attachments or they are irrelevant to the request, don't call read_attachment.`

function formatSize(bytes: number): string {
  if (bytes < 1024) return `${bytes}B`
  if (bytes < 1024 * 1024) return `${Math.round(bytes / 1024)}KB`
  return `${(bytes / 1024 / 1024).toFixed(1)}MB`
}

export function createFilesSkill(
  getAttachments: () => AttachmentMeta[],
  /** Called on each successful text read — lets the deck generator gate on unread attachments */
  onTextRead?: (path: string) => void,
): AgentSkill {
  return {
    id: 'files',
    systemPrompt: FILES_SYSTEM_PROMPT,
    tools: [
      {
        name: 'read_attachment',
        description:
          "Read an attachment's text content (parsed locally). Long files are paginated: read offset=0 first, then decide whether to continue based on the returned total character count.",
        inputSchema: {
          type: 'object',
          properties: {
            index: {
              type: 'integer',
              description: 'Attachment index (0-based, see the attachment list)',
            },
            offset: { type: 'integer', description: 'Starting character position, default 0' },
          },
          required: ['index'],
        },
      },
    ],
    buildContext: () => {
      const list = getAttachments()
      if (list.length === 0) return ''
      const lines = list.map((a, i) => `${i} | ${a.name} | .${a.ext} | ${formatSize(a.sizeBytes)}`)
      return `Attachment list (index | file name | type | size):\n${lines.join('\n')}`
    },
    executeTool: async (call) => {
      if (call.name !== 'read_attachment') {
        return { output: `Unknown tool: ${call.name}`, isError: true, summary: call.name }
      }
      const list = getAttachments()
      const index = Number(call.input.index)
      const att = Number.isInteger(index) ? list[index] : undefined
      if (!att) {
        return {
          output: 'Invalid attachment index (see the attachment list)',
          isError: true,
          summary: t('aiSumReadAttachment'),
        }
      }
      // No text extraction for images: they are already provided as multimodal images with the user message
      if (ATTACHMENT_IMAGE_EXTS.has(att.ext)) {
        return {
          output: `${att.name} is an image attachment already sent as an image with the user message; just view the image in the message, no text reading needed. To place it on a slide, use insert_web_image (or replace_image) with url=attachment://${att.name} — the original file is embedded as-is.`,
          mutated: false,
          summary: t('aiSumImageAttachment', { name: att.name }),
        }
      }
      const offset = Math.max(0, Number(call.input.offset) || 0)
      const result = await window.desktop.readAttachment(att.path, offset, READ_CHUNK_CHARS)
      if (!result.ok) {
        return {
          output: result.error ?? 'Read failed',
          isError: true,
          summary: t('aiSumReadAttachmentName', { name: att.name }),
        }
      }
      onTextRead?.(att.path)
      const end = (result.offset ?? 0) + (result.text?.length ?? 0)
      const header = `File ${att.name}, total characters ${result.totalChars}, this chunk covers ${result.offset}-${end}${
        end < (result.totalChars ?? 0)
          ? ' (not finished; continue with offset=' + end + ')'
          : ' (end of file)'
      }`
      return {
        output: `${header}\n---\n${result.text ?? ''}`,
        mutated: false,
        summary: t('aiSumReadAttachmentName', { name: att.name }),
      }
    },
  }
}
