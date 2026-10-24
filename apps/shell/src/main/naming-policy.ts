/**
 * When a document may be named by the model, and what happens when it is not.
 *
 * The rule that matters: an already-named file is never renamed behind the
 * user's back. A document keeps whatever name it was saved under until the
 * reader explicitly asks for a rename, because silently renaming a file that
 * is open in three places, or referenced from a note, is not a convenience.
 */

export type NamingTrigger =
  /** the first save of a document that has never been on disk */
  | 'first-save'
  /** the reader pressed the rename button */
  | 'manual'

export type NamingDecision =
  /** call the model */
  | 'name'
  /** do nothing, and say why */
  | 'skip-disabled'
  | 'skip-already-named'
  | 'skip-empty'
  | 'skip-no-provider'

export interface NamingContext {
  /** the model-naming preference from Settings → General */
  enabled: boolean
  /** how the naming was triggered */
  trigger: NamingTrigger
  /** the document's current path; null until its first save */
  filePath: string | null
  /** the document source handed to the model */
  content: string
  /** whether an API key is configured for the model provider */
  hasProvider: boolean
}

/**
 * Decide whether to name now. Kept separate from the transport so the rule is
 * testable without a model, a file, or an Electron main process.
 */
export function decideNaming(ctx: NamingContext): NamingDecision {
  if (!ctx.enabled && ctx.trigger === 'first-save') return 'skip-disabled'
  if (!ctx.hasProvider) return 'skip-no-provider'
  // a manual rename is the reader asking, so it is allowed to rename a named
  // file; a first save has nothing to rename yet
  if (ctx.trigger === 'first-save' && ctx.filePath) return 'skip-already-named'
  if (!hasNameable(ctx.content)) return 'skip-empty'
  return 'name'
}

/** An untitled document with no real content keeps its default name. */
function hasNameable(content: string): boolean {
  return content.replace(/[\s#*_\-`~>|[\](){}]/g, '').length > 0
}
