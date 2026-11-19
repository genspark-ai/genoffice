/** the rule fields Mailings > Rules writes: the merge reads them back by their keyword */

/** the rules that open a dialog; the first three ask for a condition on one of the list's fields */
export const MERGE_RULES = ['skipIf', 'if', 'nextIf', 'set', 'ask', 'fillIn'] as const
export type MergeRule = (typeof MERGE_RULES)[number]

/** the rules whose condition names a field of the recipient list */
export const ruleNeedsList = (rule: MergeRule): boolean =>
  rule === 'skipIf' || rule === 'if' || rule === 'nextIf'

/** the rules that go in at once: their field takes no argument */
export const PLAIN_RULES = { next: 'NEXT', mergeRec: 'MERGEREC', mergeSeq: 'MERGESEQ' } as const

export const RULE_OPS = ['=', '<>', '<', '>', '<=', '>='] as const
export type RuleOp = (typeof RULE_OPS)[number]

/** a field literal ends at its closing quote, so a quote inside it becomes a typographic one */
const literal = (text: string) => `"${text.replace(/"/g, '”')}"`

/** the condition both rules open with: a data column against a literal */
const condition = (field: string, op: RuleOp, value: string) => `«${field}» ${op} ${literal(value)}`

/** leaves the record out of the merge when the condition holds */
export const skipIfInstruction = (field: string, op: RuleOp, value: string) =>
  `SKIPIF ${condition(field, op, value)}`

/** writes one text when the condition holds and the other when it does not */
export const ifInstruction = (
  field: string,
  op: RuleOp,
  value: string,
  then: string,
  otherwise: string,
) => `IF ${condition(field, op, value)} ${literal(then)} ${literal(otherwise)}`

/** moves on to the next record inside the same document when the condition holds */
export const nextIfInstruction = (field: string, op: RuleOp, value: string) =>
  `NEXTIF ${condition(field, op, value)}`

/** gives a bookmark a text for the merge; a REF field to the bookmark shows it */
export const setInstruction = (bookmark: string, value: string) =>
  `SET ${bookmark} ${literal(value)}`

/** the default switch holds the answer when none is typed; the once switch asks one time for the whole merge */
const promptTail = (prompt: string, fallback: string, once: boolean) =>
  `${literal(prompt)} \\d ${literal(fallback)}${once ? ' \\o' : ''}`

/** asks during the merge and keeps the answer in a bookmark */
export const askInstruction = (bookmark: string, prompt: string, fallback: string, once: boolean) =>
  `ASK ${bookmark} ${promptTail(prompt, fallback, once)}`

/** asks during the merge and writes the answer where the field stands */
export const fillInInstruction = (prompt: string, fallback: string, once: boolean) =>
  `FILLIN ${promptTail(prompt, fallback, once)}`

/**
 * The rule fields that act each time they stand in the text. Each one carries
 * a field id of its own, so two of them side by side stay two fields.
 */
export const RULE_FIELD_INSTR =
  /^\s*(?:NEXT|NEXTIF|SKIPIF|MERGEREC|MERGESEQ|ASK|FILLIN|SET)(?=\s|$)/i
