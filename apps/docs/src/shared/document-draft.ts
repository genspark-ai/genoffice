import type { DocsErrorCode } from './operation-result'

export type DocsDraftState =
  'streaming' | 'awaiting_decision' | 'committed' | 'retained' | 'discarded' | 'failed'

export type DocsDraftSettlement =
  | {
      ok: true
      blocks: number
      partial: boolean
      validatedSequence?: number
      validatedCharacters?: number
      omittedTail?: boolean
    }
  | { ok: false; error: string; errorCode?: DocsErrorCode }

/** JSON carried in draft operation output; source authority is owned by the runtime. */
export interface DocsDraftInspection {
  state: DocsDraftState
  sequence: number
  validatedSequence: number
  characters: number
  validatedCharacters: number
  buffering: boolean
  conflicted: boolean
  referenceSnapshot?: { documentId: string; revision: number }
  result?: DocsDraftSettlement
}
