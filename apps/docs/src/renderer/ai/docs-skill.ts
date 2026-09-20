import type { Editor } from '@tiptap/core'
import type { AgentSkill } from '@genoffice/agent-core'
import {
  AGENT_SYSTEM_PROMPT,
  buildDocContext,
  getSelectionScope,
  type AiTrack,
  type NumIds,
} from './protocol'
import type { AiDocWriter } from './doc-writer'
import type { AiPageSetupAccess } from './page-setup'
import {
  AGENT_TOOLS,
  executeTool,
  markDocSeen,
  type AiCommentsAccess,
  type AiHeaderFooterAccess,
  type FrozenSelection,
  type AiDocExtras,
} from './tools'
import type { AiNotesAccess } from './note-ops'

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
): AgentSkill {
  // Selection frozen per run: tools act on the range the prompt described,
  // not on wherever the user's live selection has wandered mid-run. The doc
  // snapshot bounds the freeze's validity (see FrozenSelection).
  let frozen: FrozenSelection | null = null
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
  return {
    id: 'docx',
    // live: the predicate is re-read before every model request
    get systemPrompt() {
      return AGENT_SYSTEM_PROMPT + mediaOffNote()
    },
    get tools() {
      const hidden = mediaToolsOff()
      return hidden.size === 0 ? AGENT_TOOLS : AGENT_TOOLS.filter((t) => !hidden.has(t.name))
    },
    buildContext: () => {
      const editor = getEditor()
      markDocSeen(editor) // the context the model receives is the freshness baseline for index-addressed writes
      frozen = { scope: getSelectionScope(editor), doc: editor.state.doc }
      return buildDocContext(
        editor,
        frozen.scope,
        getComments?.()?.list(),
        getHf?.()?.read(),
        getPageSetup?.()?.list(),
      )
    },
    executeTool: (call, signal) =>
      executeTool(
        getEditor(),
        call,
        getNumIds(),
        getTrack?.(),
        signal,
        frozen,
        getComments?.(),
        getHf?.(),
        getWriter?.(),
        getPageSetup?.(),
        getExtras?.(),
        getNotes?.(),
      ),
  }
}
