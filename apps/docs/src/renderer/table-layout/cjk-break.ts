/**
 * Word's min-content break opportunities inside East Asian text (probe
 * 2026-11-09, ja-JP Yu Mincho, word-truths.md "CJK min-content"): a line may
 * break between any two East Asian characters, so the widest unbreakable unit
 * is one character - except that a closer never starts a line and an opener
 * never ends one, which glues it to its neighbour (two-character floor).
 * Small kana, the sokuon and the prolonged sound mark may start a line; a
 * Latin run beside East Asian text is a token of its own.
 */
const NO_LINE_START = new Set('、。，．」）』】｝〕》〉〗〙・')
const NO_LINE_END = new Set('「（『【｛〔《〈〖〘')

export function cjkForbidsLineStart(ch: string): boolean {
  return NO_LINE_START.has(ch)
}
export function cjkForbidsLineEnd(ch: string): boolean {
  return NO_LINE_END.has(ch)
}
