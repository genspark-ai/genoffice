/**
 * Excel resolves a cell's reading order from its content ("context" reading
 * order): the first strong directional character decides the paragraph
 * direction, so "1446<AR>" renders with the Arabic suffix on the LEFT and the
 * digits on the RIGHT. Univer's fast text path draws every line with the
 * canvas default (ltr) paragraph direction, which mirrors such mixed-direction
 * strings. Detect the paragraph direction per UAX#9 P2/P3 (first strong
 * character) and hand it to the canvas bidi algorithm via ctx.direction.
 *
 * textAlign is pinned to 'left' while drawing: the fast path computes line x
 * offsets for a left anchor, and the canvas default 'start' would flip the
 * anchor to the right edge under direction:'rtl'.
 */
import { Text } from '@univerjs/engine-render'

const STRONG_RTL_MARK = /[\u200F\u061C]/ // RLM, ALM
const LTR_MARK = '\u200E' // LRM
const RTL_SCRIPT_LETTER =
  /[\p{Script=Arabic}\p{Script=Hebrew}\p{Script=Syriac}\p{Script=Thaana}\p{Script=Nko}\p{Script=Samaritan}\p{Script=Mandaic}\p{Script=Adlam}]/u
const LETTER = /\p{L}/u

/// First-strong-character scan. Digits, punctuation and spaces are weak or
/// neutral and skipped; RTL-script digits (e.g. Arabic-Indic ٠-٩) are not
/// letters, so they stay weak here too, matching the bidi classes.
export function resolveBidiDirection(text: string): 'ltr' | 'rtl' {
  for (const ch of text) {
    if (STRONG_RTL_MARK.test(ch)) return 'rtl'
    if (ch === LTR_MARK) return 'ltr'
    if (LETTER.test(ch)) return RTL_SCRIPT_LETTER.test(ch) ? 'rtl' : 'ltr'
  }
  return 'ltr'
}

interface DirectionalContext {
  save(): void
  restore(): void
  direction: CanvasDirection
  textAlign: CanvasTextAlign
}

let installed = false

export function installRtlTextDirectionFix(): void {
  if (installed) return
  installed = true

  const textClass = Text as unknown as {
    drawWith(ctx: DirectionalContext, props: { text?: unknown }, skeleton?: unknown): void
  }
  const previousDrawWith = textClass.drawWith
  if (typeof previousDrawWith !== 'function') return
  textClass.drawWith = function (
    this: unknown,
    ctx: DirectionalContext,
    props: { text?: unknown },
    skeleton?: unknown,
  ): void {
    if (typeof props?.text !== 'string' || resolveBidiDirection(props.text) !== 'rtl') {
      return previousDrawWith.call(this, ctx, props, skeleton)
    }
    ctx.save()
    ctx.direction = 'rtl'
    ctx.textAlign = 'left'
    try {
      return previousDrawWith.call(this, ctx, props, skeleton)
    } finally {
      ctx.restore()
    }
  }
}
