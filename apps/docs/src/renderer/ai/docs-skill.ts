import type { Editor } from '@tiptap/core'
import type { AgentSkill } from '@genoffice/agent-core'
import {
  AGENT_SYSTEM_PROMPT,
  buildDocContext,
  getSelectionScope,
  type AiTrack,
  type NumIds,
} from './protocol'
import {
  AGENT_TOOLS,
  executeTool,
  markDocSeen,
  type AiCommentsAccess,
  type AiHeaderFooterAccess,
  type FrozenSelection,
} from './tools'

const IMAGE_GEN_OFF_NOTE =
  '\n\nNote: generate_image is currently unavailable (no image provider: signed out of Genspark or cloud tools off, and no media API key in Settings). Do not call or promise it; use image_search for imagery.'

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
): AgentSkill {
  // Selection frozen per run: tools act on the range the prompt described,
  // not on wherever the user's live selection has wandered mid-run. The doc
  // snapshot bounds the freeze's validity (see FrozenSelection).
  let frozen: FrozenSelection | null = null
  return {
    id: 'docx',
    // live: the predicate is re-read before every model request
    get systemPrompt() {
      return imageGenAvailable?.() === false
        ? AGENT_SYSTEM_PROMPT + IMAGE_GEN_OFF_NOTE
        : AGENT_SYSTEM_PROMPT
    },
    get tools() {
      return imageGenAvailable?.() === false
        ? AGENT_TOOLS.filter((t) => t.name !== 'generate_image')
        : AGENT_TOOLS
    },
    buildContext: () => {
      const editor = getEditor()
      markDocSeen(editor) // the context the model receives is the freshness baseline for index-addressed writes
      frozen = { scope: getSelectionScope(editor), doc: editor.state.doc }
      return buildDocContext(editor, frozen.scope, getComments?.()?.list(), getHf?.()?.read())
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
      ),
  }
}
