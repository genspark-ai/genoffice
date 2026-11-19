/**
 * The app-side half of model file naming.
 *
 * ## What lives here and what does not
 *
 * The main process owns the model call and the timeout
 * (`apps/shell/src/main/ai-naming-service.ts`). What it deliberately does *not*
 * own is the text: only an editor knows which of its spans are withheld from
 * the model, so the caller has to pass the **projected** text. That is the whole
 * safety property, and it is why this helper takes content rather than a
 * document — there is no path by which raw text reaches the model from here.
 *
 * ## The empty-string contract
 *
 * `suggestFileName` never throws and never returns a name it is not sure of.
 * An empty string means "no name to offer" — no provider, an unreachable
 * provider, a refusal, or a reply that sanitised away to nothing. The caller
 * keeps the name it already had, which is the only acceptable outcome: a
 * rename that costs the reader their file name is worse than no rename at all.
 */

export interface SuggestFileNameOptions {
  /**
   * The text the **model** is allowed to see. Pass the editor's projected view,
   * never the raw document — withheld spans must stay withheld from this call
   * too, and this is the only place the text is handed over.
   */
  content: string
}

interface NamingApi {
  suggestFileName?: (input: { content: string }) => Promise<{
    ok: boolean
    name?: string
    reason?: string
  }>
}

/**
 * The naming bridge.
 *
 * One global, one name: every window that can name a file exposes the channel
 * as `aiNamingApi` — the shell for its own button, each editor for its own.
 *
 * The shell's `aiOffice` namespace is deliberately **not** used here. It is
 * the shell's home API, and a second, partial `aiOffice` in an editor window
 * would shadow the real one in that window: `window.aiOffice.setTheme` would
 * disappear, and anything that identifies the shell window by that global —
 * the e2e suite included — would find the editor instead.
 */
function api(): NamingApi {
  return (window as unknown as { aiNamingApi?: NamingApi }).aiNamingApi ?? {}
}

/**
 * Ask the model for a file name.
 *
 * Returns the name, or '' when the model could not answer one. The caller puts
 * the result in front of the reader before anything is written.
 */
export async function suggestFileName(options: SuggestFileNameOptions): Promise<string> {
  const call = api().suggestFileName
  if (!call) return ''
  // nothing worth sending: an empty document has no name to derive
  if (!options.content.trim()) return ''
  try {
    const result = await call({ content: options.content })
    if (!result?.ok) return ''
    // an empty name is a failed naming, not a request to name the file ""
    return result.name?.trim() ? result.name.trim() : ''
  } catch {
    return ''
  }
}
