import type { ExecutedToolCall } from '@genoffice/agent-core'

/**
 * Claimed-selection guard for the sheets skill (AgentSkill.verifyResponse):
 * a reply that tells the user "I selected / located / highlighted ... for
 * you" is only acceptable when a select_range call actually succeeded during
 * the run. Prompt rules alone are soft — models still occasionally narrate an
 * action instead of performing it (claimed-action hallucination); this check
 * is the mechanical backstop that forces one corrective turn.
 *
 * Detection is intentionally narrow (Chinese + English claim phrasings, the
 * dominant reply languages). A false positive is cheap: the correction tells
 * the model to either really call select_range or reword, both fine outcomes,
 * and the loop caps it at one extra turn per run.
 */

// The perfective marker (U+5DF2, "already") plus a selection verb, allowing
// a short gap for filler words ("already [for you] selected ..."). The gap
// must not cross punctuation — otherwise "done. please select ..." (an
// instruction to the user) would match — and the attributive form
// "the already-selected range" (referring to the user's own selection) is
// excluded by the lookahead guard rejecting U+7684 after the verb.
const ZH_SELECTION_CLAIM =
  /已[^。．.!！?？，,;；:：\n]{0,8}?(?:定位|选中|選中|选定|選定|圈选|圈選|高亮|跳转|跳轉)(?!的)/
const EN_SELECTION_CLAIM =
  /\bI(?:'ve| have)?(?: now| already| just)? (?:selected|highlighted|located|jumped to|moved the selection)\b|\bselection (?:has been|is now|was) (?:moved|set|placed)\b/i

const CORRECTION =
  '[System check] Your reply claims that something was located / selected / highlighted on screen for the user, ' +
  'but no successful select_range call happened during this run — the visible selection has NOT moved. ' +
  'Either call select_range now to actually perform the action, or rewrite your reply so it no longer claims an action that did not happen.'

export function claimsSelection(text: string): boolean {
  return ZH_SELECTION_CLAIM.test(text) || EN_SELECTION_CLAIM.test(text)
}

/** verifyResponse impl for the sheets skill; returns a corrective instruction or null. */
export function verifySheetsResponse(
  finalText: string,
  executed: readonly ExecutedToolCall[],
): string | null {
  if (!claimsSelection(finalText)) return null
  if (executed.some((call) => call.name === 'select_range' && call.ok)) return null
  return CORRECTION
}
