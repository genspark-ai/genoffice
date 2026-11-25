// Modified for ForaOffice from GenOffice v0.12.852.
import { placeholderInstruction } from './redact'
import { redactLabelsOf } from './redact-view'
import type { Editor } from '@tiptap/core'
import type { AgentSkill } from '@genoffice/agent-core'
import { AGENT_SYSTEM_PROMPT, type AiTrack, type NumIds } from './protocol'
import type { AiDocWriter } from './doc-writer'
import type { AiPageSetupAccess } from './page-setup'
import {
  DOCS_EDITOR_TOOLS,
  type AiCommentsAccess,
  type AiHeaderFooterAccess,
  type AiDocExtras,
} from './tools'
import type { AiNotesAccess } from './note-ops'
import { createDocsRuntime, type DocsRuntime } from '../runtime/docs-runtime'
import type { DocsResourceTools } from '../../shared/resource-tools'
import { DOCS_FILE_COMMANDS } from '../../shared/resource-tools'
import { NATIVE_DOCS_WORKFLOW_TOOLS } from './native-tool-catalog'
import { createNativeDocsWorkflows } from './native-workflows'
import { z } from 'zod'

const IMAGE_GEN_OFF_NOTE =
  '\n\nNote: generate_image is currently unavailable (no image provider: signed out of Genspark or cloud tools off, and no media API key in Settings). Do not call or promise it; use image_search for imagery.'

const MEDIA_ANALYSIS_OFF_NOTE =
  '\n\nNote: analyze_media is currently unavailable (no media provider: signed out of Genspark or cloud tools off, and no media API key in Settings). You cannot see what a picture inside the document shows — tell the user that instead of guessing at its content.'

/**
 * The docx capability as an AgentSkill: document skeleton context, the five
 * document tools, and the local executor. Future apps register their own
 * skills (Excel / PPT) against the same agent loop.
 */
export function createDocsSkill(
  getEditor: () => Editor,
  getNumIds: () => NumIds,
  getTrack?: () => AiTrack | undefined,
  getComments?: () => AiCommentsAccess | undefined,
  getHf?: () => AiHeaderFooterAccess | undefined,
  /** live predicate (gsk login && cloud-tools toggle, or a BYOK media key); false hides generate_image */
  imageGenAvailable?: () => boolean,
  /** streaming long-form writer behind write_document (panel-owned: progress chip, partial keep/discard) */
  getWriter?: () => AiDocWriter | undefined,
  getPageSetup?: () => AiPageSetupAccess | undefined,
  /** styles.xml catalog + page watermark stores (define_style / applyStyle / set_watermark) */
  getExtras?: () => AiDocExtras | undefined,
  getNotes?: () => AiNotesAccess | undefined,
  /** same as imageGenAvailable, for analyze_media (seeing pictures that are in the document) */
  mediaAnalysisAvailable?: () => boolean,
  getRuntime?: () => DocsRuntime,
  getResources?: () => DocsResourceTools | undefined,
): AgentSkill {
  const runtime =
    getRuntime?.() ??
    createDocsRuntime({
      current: () => ({
        editor: getEditor(),
        identity: getEditor(),
        numIds: getNumIds(),
        track: getTrack?.(),
        comments: getComments?.(),
        headerFooter: getHf?.(),
        pageSetup: getPageSetup?.(),
        extras: getExtras?.(),
        notes: getNotes?.(),
      }),
    })
  /** tools that need a media provider: Genspark login + cloud tools, or a BYOK media key in Settings */
  const mediaToolsOff = (): Set<string> => {
    const hidden = new Set<string>()
    if (imageGenAvailable?.() === false) hidden.add('generate_image')
    if (mediaAnalysisAvailable?.() === false) hidden.add('analyze_media')
    return hidden
  }
  const mediaOffNote = (): string =>
    [...mediaToolsOff()]
      .map((name) => (name === 'generate_image' ? IMAGE_GEN_OFF_NOTE : MEDIA_ANALYSIS_OFF_NOTE))
      .join('')
  const client = runtime.connect({ id: crypto.randomUUID(), resources: getResources })
  const workflows = createNativeDocsWorkflows(client, getWriter)
  const catalog = [...DOCS_EDITOR_TOOLS, ...NATIVE_DOCS_WORKFLOW_TOOLS]
  const schemas = new Map(
    NATIVE_DOCS_WORKFLOW_TOOLS.map((tool) => [tool.name, z.fromJSONSchema(tool.inputSchema)]),
  )
  return {
    id: 'docx',
    // live: the predicate is re-read before every model request
    get systemPrompt() {
      // only when the open document actually withholds something, so a
      // document without spans pays nothing and the model is not told to
      // look for markers that are not there
      const editor = getEditor()
      const labels = editor
        ? redactLabelsOf(editor.getJSON() as unknown as Parameters<typeof redactLabelsOf>[0])
        : []
      const redact = labels.length > 0 ? `\n\n${placeholderInstruction(labels)}` : ''
      return AGENT_SYSTEM_PROMPT + mediaOffNote() + redact
    },
    get tools() {
      const hidden = mediaToolsOff()
      return catalog.filter(
        (tool) => !hidden.has(tool.name) && (tool.name !== 'write_document' || !!getWriter?.()),
      )
    },
    buildContext: () => client.context().text,
    executeTool: async (call, signal) => {
      const schema = schemas.get(call.name)
      if (schema) {
        if (mediaToolsOff().has(call.name))
          return {
            output: `${call.name} is unavailable in the native provider configuration`,
            summary: call.name,
            isError: true,
            errorCode: 'unavailable',
            committed: false,
          }
        const input = schema.safeParse(call.input)
        if (!input.success)
          return {
            output: input.error.message,
            summary: call.name,
            isError: true,
            errorCode: 'invalid_input',
            committed: false,
          }
        return workflows.execute(call, signal)
      }
      if (
        call.name === 'run_cli' &&
        !DOCS_FILE_COMMANDS.has(String((call.input.args as unknown[] | undefined)?.[0] ?? ''))
      )
        return client.executeNativeCli(call, signal)
      return client.execute(call, signal)
    },
  }
}
