import { useRef } from 'react'

/**
 * Tells the Enter that confirms an IME candidate apart from the Enter that
 * means "submit". `nativeEvent.isComposing` alone is not enough: WebKit fires
 * `compositionend` first and only then the confirming Enter keydown, which by
 * that point reports `isComposing === false` with an ordinary keyCode — so a
 * JP/CN/KR user confirming a candidate would submit a half-typed instruction.
 * The guard therefore combines three signals: our own composition flag from
 * compositionstart/end, the event's `isComposing`, the legacy IME keyCode 229
 * (Chromium keeps it on composition keydowns), and additionally swallows an
 * Enter arriving in the same tick as `compositionend` (the WebKit ordering).
 * The flag resets on a 0ms timer, so the user's next, deliberate Enter
 * submits as usual.
 */
export type ImeEnterGuard = {
  readonly handleCompositionStart: () => void
  readonly handleCompositionEnd: () => void
  /** pass the native keydown event (React: `e.nativeEvent`) */
  readonly isImeConfirm: (event: { isComposing?: boolean; keyCode?: number }) => boolean
}

export function createImeEnterGuard(
  scheduleReset: (reset: () => void) => void = (reset) => setTimeout(reset, 0),
): ImeEnterGuard {
  let composing = false
  let justEnded = false
  return {
    handleCompositionStart() {
      composing = true
    },
    handleCompositionEnd() {
      composing = false
      justEnded = true
      scheduleReset(() => {
        justEnded = false
      })
    },
    isImeConfirm(event) {
      return composing || justEnded || event.isComposing === true || event.keyCode === 229
    },
  }
}

/** per-component guard instance; the returned handlers are referentially stable */
export function useImeEnterGuard(): ImeEnterGuard {
  const ref = useRef<ImeEnterGuard | null>(null)
  if (ref.current === null) ref.current = createImeEnterGuard()
  return ref.current
}
