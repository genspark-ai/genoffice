/* Office-style ScreenTips for ribbon/toolbar buttons: one delegated listener
 * drives a single fixed-position tip element, replacing the native `title`
 * attribute (whose ~1s delay is not adjustable and whose styling cannot show
 * shortcuts or descriptions).
 *
 * Markup contract, on any element:
 *   data-tip        — command name (required; the tip line)
 *   data-tip-kbd    — optional shortcut, rendered dimmed after the name
 *   data-tip-detail — optional one-line description under the name
 *   data-tip-anchor — optional descendant selector; the tip is positioned
 *                     relative to that element instead of the host (e.g. a
 *                     wide row anchoring its tip to a trailing icon)
 *   data-tip-place  — optional 'right' places the tip beside the anchor,
 *                     vertically centered (default: below, centered)
 *
 * Timing follows the Windows/Office ScreenTip standard: 500ms initial delay,
 * fast (100ms) reshow while "warm" — a tip was visible less than 500ms ago —
 * so sweeping the pointer along a toolbar walks the tip across the buttons,
 * and auto-hide after 5s of resting on the same control. Clicking hides the
 * tip and keeps it hidden until the pointer leaves that control. */

const INITIAL_DELAY_MS = 500
const RESHOW_DELAY_MS = 100
const WARM_WINDOW_MS = 500
const AUTO_HIDE_MS = 5000
const OFFSET_PX = 6
const EDGE_PAD_PX = 4

interface TipState {
  el: HTMLDivElement
  name: HTMLSpanElement
  kbd: HTMLSpanElement
  detail: HTMLDivElement
}

let tip: TipState | null = null
let showTimer: number | null = null
let autoHideTimer: number | null = null
let anchor: Element | null = null
// clicked anchor: the tip stays hidden until the pointer leaves it (Office behavior)
let suppressed: Element | null = null
let warmUntil = 0

function ensureTip(doc: Document): TipState {
  if (tip && tip.el.isConnected) return tip
  const el = doc.createElement('div')
  el.className = 'ui-screentip'
  el.setAttribute('role', 'tooltip')
  const name = doc.createElement('span')
  name.className = 'ui-screentip-name'
  const kbd = doc.createElement('span')
  kbd.className = 'ui-screentip-kbd'
  const detail = doc.createElement('div')
  detail.className = 'ui-screentip-detail'
  el.append(name, kbd, detail)
  doc.body.appendChild(el)
  tip = { el, name, kbd, detail }
  return tip
}

function cancelShow(): void {
  if (showTimer !== null) {
    window.clearTimeout(showTimer)
    showTimer = null
  }
}

function hide(): void {
  cancelShow()
  if (autoHideTimer !== null) {
    window.clearTimeout(autoHideTimer)
    autoHideTimer = null
  }
  if (tip && tip.el.style.visibility === 'visible') {
    tip.el.style.visibility = 'hidden'
    warmUntil = Date.now() + WARM_WINDOW_MS
  }
  anchor = null
}

function show(el: Element, doc: Document): void {
  const text = el.getAttribute('data-tip')
  if (!text) return
  const t = ensureTip(doc)
  t.name.textContent = text
  const kbd = el.getAttribute('data-tip-kbd')
  t.kbd.textContent = kbd ?? ''
  t.kbd.style.display = kbd ? '' : 'none'
  const detail = el.getAttribute('data-tip-detail')
  t.detail.textContent = detail ?? ''
  t.detail.style.display = detail ? '' : 'none'

  // measure hidden, then place below the control (above when it would clip);
  // data-tip-anchor re-anchors to a descendant, data-tip-place='right' puts
  // the tip beside the anchor instead of under it
  t.el.style.visibility = 'hidden'
  t.el.style.left = '0px'
  t.el.style.top = '0px'
  const anchorSel = el.getAttribute('data-tip-anchor')
  const rect = ((anchorSel && el.querySelector(anchorSel)) || el).getBoundingClientRect()
  const tipRect = t.el.getBoundingClientRect()
  const maxLeft = window.innerWidth - tipRect.width - EDGE_PAD_PX
  let left: number
  let top: number
  if (el.getAttribute('data-tip-place') === 'right') {
    left = rect.right + OFFSET_PX
    if (left > maxLeft) left = Math.max(EDGE_PAD_PX, rect.left - tipRect.width - OFFSET_PX)
    top = Math.max(
      EDGE_PAD_PX,
      Math.min(
        rect.top + rect.height / 2 - tipRect.height / 2,
        window.innerHeight - tipRect.height - EDGE_PAD_PX,
      ),
    )
  } else {
    left = Math.max(EDGE_PAD_PX, Math.min(rect.left + rect.width / 2 - tipRect.width / 2, maxLeft))
    top = rect.bottom + OFFSET_PX
    if (top + tipRect.height > window.innerHeight - EDGE_PAD_PX)
      top = rect.top - tipRect.height - OFFSET_PX
  }
  t.el.style.left = `${Math.round(left)}px`
  t.el.style.top = `${Math.round(top)}px`
  t.el.style.visibility = 'visible'

  if (autoHideTimer !== null) window.clearTimeout(autoHideTimer)
  autoHideTimer = window.setTimeout(hide, AUTO_HIDE_MS)
}

/** Install the global ScreenTip listeners. Call once from the app entry; returns an uninstaller. */
export function installScreenTips(doc: Document = document): () => void {
  const onPointerOver = (e: PointerEvent): void => {
    const target = e.target instanceof Element ? e.target : null
    const el = target?.closest('[data-tip]') ?? null
    if (el === anchor) return
    if (suppressed && suppressed !== el) suppressed = null
    hide()
    if (!el || el === suppressed) return
    anchor = el
    const delay = Date.now() < warmUntil ? RESHOW_DELAY_MS : INITIAL_DELAY_MS
    showTimer = window.setTimeout(() => {
      showTimer = null
      // the anchor may have been re-rendered or removed while the timer ran
      if (anchor === el && el.isConnected) show(el, doc)
    }, delay)
  }
  // pointer left the window entirely (no pointerover follows)
  const onPointerOut = (e: PointerEvent): void => {
    if (!e.relatedTarget) {
      suppressed = null
      hide()
    }
  }
  const onPointerDown = (): void => {
    suppressed = anchor
    hide()
    warmUntil = 0
  }
  const onHide = (): void => hide()

  doc.addEventListener('pointerover', onPointerOver, true)
  doc.addEventListener('pointerout', onPointerOut, true)
  doc.addEventListener('pointerdown', onPointerDown, true)
  doc.addEventListener('scroll', onHide, true)
  window.addEventListener('blur', onHide)
  window.addEventListener('resize', onHide)
  return () => {
    doc.removeEventListener('pointerover', onPointerOver, true)
    doc.removeEventListener('pointerout', onPointerOut, true)
    doc.removeEventListener('pointerdown', onPointerDown, true)
    doc.removeEventListener('scroll', onHide, true)
    window.removeEventListener('blur', onHide)
    window.removeEventListener('resize', onHide)
    hide()
    tip?.el.remove()
    tip = null
  }
}
