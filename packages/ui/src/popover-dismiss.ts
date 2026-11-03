/**
 * Unified dropdown/popover dismissal (the slides PR #505 model, generalized):
 *
 *  1. a press anywhere outside the popover closes it
 *  2. window blur closes it (app/window switch)
 *  3. a press on the shell tab strip — a sibling WebContentsView whose input
 *     never reaches this document — closes it via the app:chrome-pressed IPC
 *     relay (the shell main process also broadcasts it on window drag)
 *
 * Apps whose ribbon tab row doubles as the frameless-window drag region must
 * additionally suspend that region while a popover is open (drag regions
 * swallow mouse events, so no listener can see presses on the blank band).
 * While any popover installed here is open, `<html>` carries the
 * `genoffice-popover-open` class — suspend the drag region with
 * `html.genoffice-popover-open .your-drag-row { -webkit-app-region: no-drag; }`.
 */
import { useEffect, useRef } from 'react'

type ChromePressedApi = { onChromePressed?: (handler: () => void) => () => void }

/** Each app's preload exposes the app:chrome-pressed subscription under its
 * own namespace; probe the known ones so callers never need to care. */
export function subscribeChromePressed(handler: () => void): (() => void) | undefined {
  const w = window as unknown as Record<string, ChromePressedApi | undefined>
  for (const name of [
    'slidesApi',
    'desktopApi',
    'desktop',
    'pdfApi',
    'markdownApi',
    'aiOfficeTabs',
  ]) {
    const sub = w[name]?.onChromePressed
    if (typeof sub === 'function') return sub.call(w[name], handler)
  }
  return undefined
}

/** open-popover refcount driving the html-level drag-region suspension class */
let openPopovers = 0
function bumpOpenPopovers(delta: 1 | -1): void {
  openPopovers = Math.max(0, openPopovers + delta)
  document.documentElement.classList.toggle('genoffice-popover-open', openPopovers > 0)
}

export interface PopoverDismissOptions {
  /**
   * Roots the press may land in without dismissing (the popover panel and the
   * trigger that toggles it). When provided, the outside-press listener runs
   * on capture-phase pointerdown with this containment guard. When omitted,
   * it runs on bubble-phase mousedown and the popover must protect itself
   * with `onMouseDown={(e) => e.stopPropagation()}` (the slides convention).
   */
  inside?: () => ReadonlyArray<Element | null | undefined>
  /**
   * Also close on Escape. Every dropdown grew its own 7-line keydown effect
   * for this; opting in folds it into the same install/teardown.
   */
  escape?: boolean
}

/** Install the three dismissal listeners; returns a teardown function. */
export function installPopoverDismiss(
  close: () => void,
  options?: PopoverDismissOptions,
): () => void {
  const inside = options?.inside
  const onPress = (e: Event) => {
    if (inside) {
      const target = e.target as Node | null
      if (target) {
        for (const root of inside()) if (root && root.contains(target)) return
      }
    }
    close()
  }
  const onBlur = () => close()
  const onKey = (e: KeyboardEvent) => {
    if (e.key === 'Escape') close()
  }
  if (inside) window.addEventListener('pointerdown', onPress, true)
  else window.addEventListener('mousedown', onPress)
  window.addEventListener('blur', onBlur)
  if (options?.escape) window.addEventListener('keydown', onKey)
  const offChrome = subscribeChromePressed(close)
  bumpOpenPopovers(1)
  let tornDown = false
  return () => {
    if (tornDown) return
    tornDown = true
    if (inside) window.removeEventListener('pointerdown', onPress, true)
    else window.removeEventListener('mousedown', onPress)
    window.removeEventListener('blur', onBlur)
    if (options?.escape) window.removeEventListener('keydown', onKey)
    offChrome?.()
    bumpOpenPopovers(-1)
  }
}

/**
 * Escape closes an open thing and that is all: the exact behavior of the
 * window-keydown effect that every dropdown, dialog and armed mode used to
 * grow by hand. For states that are true popovers prefer `useDismissablePopover`
 * with `escape: true` (it also handles outside presses, blur and the shell tab
 * strip); this one is for the rest — armed modes, inline editors, modal overlays.
 * No preventDefault: Escape-close sites never swallowed the key's default.
 */
export function useEscapeToClose(active: boolean, close: () => void): void {
  const closeRef = useRef(close)
  closeRef.current = close
  useEffect(() => {
    if (!active) return
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') closeRef.current()
    }
    window.addEventListener('keydown', onKey)
    return () => window.removeEventListener('keydown', onKey)
  }, [active])
}

/** React binding: listeners live only while `open` is true. */
export function useDismissablePopover(
  open: boolean,
  close: () => void,
  options?: PopoverDismissOptions,
): void {
  const closeRef = useRef(close)
  closeRef.current = close
  const insideRef = useRef(options?.inside)
  insideRef.current = options?.inside
  const guarded = options?.inside != null
  const escape = options?.escape === true
  useEffect(() => {
    if (!open) return
    return installPopoverDismiss(
      () => closeRef.current(),
      guarded ? { inside: () => insideRef.current?.() ?? [], escape } : { escape },
    )
  }, [open, guarded, escape])
}
