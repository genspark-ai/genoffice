/**
 * The app-side half of model file naming.
 *
 * ## What lives here and what does not
 *
 * The main process owns the model call, the preference and the timeout
 * (`apps/shell/src/main/ai-naming-service.ts`). What it deliberately does *not*
 * own is the text: only an editor knows which of its spans are withheld from
 * the model, so the caller has to pass the **projected** text. That is the whole
 * safety property, and it is why this helper takes content rather than a
 * document — there is no path by which raw text reaches the model from here.
 *
 * The trigger logic is shared so six editors cannot disagree about when a
 * document gets named, and so the "one attempt per document" bookkeeping is
 * written once rather than six times.
 *
 * ## The fallback contract
 *
 * `nameForSave` never throws and never returns an empty string. A declined,
 * failed or empty answer means the caller keeps the name it already had, which
 * is the only acceptable outcome: a rename that costs the reader their file
 * name is worse than no rename at all.
 */

export type NamingTrigger = 'first-save' | 'manual'

export interface NameForSaveOptions {
  /**
   * The text the **model** is allowed to see. Pass the editor's projected view,
   * never the raw document — withheld spans must stay withheld from this call
   * too, and this is the only place the text is handed over.
   */
  content: string
  trigger: NamingTrigger
  /** where the document is being saved, when known */
  filePath?: string | null
  /** the name to keep when naming does not happen */
  fallback: string
  /**
   * Force the attempt even with the preference off — the manual trigger is the
   * reader asking for it, so it goes through whether or not first-save naming
   * is switched on.
   */
  ignorePreference?: boolean
}

interface NamingApi {
  suggestFileName?: (input: {
    content: string
    trigger: NamingTrigger
    filePath?: string | null
  }) => Promise<{ ok: boolean; name?: string; reason?: string }>
  getFileNamingEnabled?: () => Promise<boolean>
  setFileNamingEnabled?: (on: boolean) => Promise<boolean>
}

function api(): NamingApi {
  return (window as unknown as { aiOffice?: NamingApi }).aiOffice ?? {}
}

/**
 * Ask for a name, and fall back to what the caller already had.
 *
 * Returns the name to use. Never rejects: a model that is switched off,
 * unconfigured, unreachable, or unhelpful leaves the document named exactly as
 * it was.
 */
export async function nameForSave(options: NameForSaveOptions): Promise<string> {
  const { content, trigger, filePath, fallback } = options
  const call = api().suggestFileName
  if (!call) return fallback
  // nothing worth sending: an empty document has no name to derive
  if (!content.trim()) return fallback
  // the manual trigger is the reader asking, so it does not consult the setting
  if (trigger === 'first-save' && options.ignorePreference !== true) {
    const enabled = await api()
      .getFileNamingEnabled?.()
      .catch(() => false)
    if (!enabled) return fallback
  }
  try {
    const result = await call({ content, trigger, filePath: filePath ?? null })
    if (!result?.ok) return fallback
    // an empty name is a failed naming, not a request to name the file ""
    return result.name?.trim() ? result.name.trim() : fallback
  } catch {
    return fallback
  }
}

/** Whether first-save naming is switched on, for a settings row. */
export function fileNamingEnabled(): Promise<boolean> {
  const call = api().getFileNamingEnabled
  if (!call) return Promise.resolve(false)
  return call()
    .then((v) => v === true)
    .catch(() => false)
}

/** The settings-row toggle; the returned value is what the main process kept. */
export function setFileNamingEnabled(on: boolean): Promise<boolean> {
  const call = api().setFileNamingEnabled
  if (!call) return Promise.resolve(false)
  return call(on)
    .then((v) => v === true)
    .catch(() => false)
}
