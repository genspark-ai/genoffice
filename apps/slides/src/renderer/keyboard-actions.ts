/**
 * Global keyboard shortcuts extracted from App.tsx: ⌘Z undo, ⌘F
 * find, zoom, element clipboard, format painter, selection cycling, nudging…
 * The handler reads the latest App state through ActionCtx, so App attaches it
 * once with an empty dependency list.
 */
import type { ActionCtx } from './action-context'
import * as clipboardActions from './clipboard-actions'
import * as arrangeActions from './arrange-actions'
import * as slideActions from './slide-actions'
import * as showActions from './show-actions'
import { shouldRouteUndoToDeck } from './undo-routing'

/** Whether focus is in a text input (input/textarea/contentEditable) — these cases use native undo/delete */
function inTextField(): boolean {
  const el = document.activeElement as HTMLElement | null
  return !!el && (el.tagName === 'INPUT' || el.tagName === 'TEXTAREA' || el.isContentEditable)
}

// Shortcuts: ⌘/Ctrl+Z undo, ⇧⌘Z / Ctrl+Y redo, ⌘/Ctrl+=/-/0 zoom,
// Delete/Backspace delete selection, arrow keys nudge selection
export function handleGlobalKeydown(ctx: ActionCtx, e: KeyboardEvent): void {
  if (ctx.slideShow || ctx.presenter) return // In show/presenter view: navigation keys are handled by those views
  const mod = e.metaKey || e.ctrlKey
  const inField = inTextField()
  const { editing, selectedIds, slide } = ctx
  // F5 show from start / ⇧F5 from current slide
  if (e.key === 'F5' && !mod) {
    e.preventDefault()
    showActions.startSlideShow(ctx, !e.shiftKey)
    return
  }
  // Undo/redo (menu accelerators normally intercept; fallback for shell/menuless scenarios)
  if (mod && !e.altKey && (e.key === 'z' || e.key === 'Z')) {
    if (editing || (inField && !shouldRouteUndoToDeck(e.target as HTMLElement))) return
    e.preventDefault()
    void (e.shiftKey ? ctx.redo() : ctx.undo())
    return
  }
  if (mod && !e.altKey && (e.key === 'y' || e.key === 'Y')) {
    if (editing || inField) return
    e.preventDefault()
    void ctx.redo()
    return
  }
  // ⌘F find/replace
  if (mod && !e.altKey && (e.key === 'f' || e.key === 'F')) {
    if (editing) return
    e.preventDefault()
    ctx.setFindOpen(true)
    return
  }
  // ⌘P print
  if (mod && !e.altKey && !e.shiftKey && (e.key === 'p' || e.key === 'P')) {
    if (editing || inField) return
    e.preventDefault()
    ctx.setPrintDlgOpen(true)
    return
  }
  // Zoom shortcut fallback
  if (mod && !inField && (e.key === '=' || e.key === '+')) {
    e.preventDefault()
    ctx.setZoom((z) => Math.min(z * 1.15, 3))
    return
  }
  if (mod && !inField && e.key === '-') {
    e.preventDefault()
    ctx.setZoom((z) => Math.max(z / 1.15, 0.25))
    return
  }
  if (mod && !inField && e.key === '0') {
    e.preventDefault()
    ctx.setZoom(1)
    return
  }
  if (editing || inField) return
  // ⌘C/⌘X with text dragged in plain DOM (e.g. AI panel, focus on body): let the
  // native copy run instead of hijacking it for the slide/element clipboard
  if (mod && !e.altKey && !e.shiftKey && ['c', 'C', 'x', 'X'].includes(e.key)) {
    const sel = window.getSelection()
    if (sel && !sel.isCollapsed) return
  }
  // Esc: exit format painter continuous mode
  if (e.key === 'Escape' && ctx.brushMode) {
    e.preventDefault()
    ctx.setBrushMode(null)
    ctx.setStatus('')
    return
  }
  // Element clipboard (text editing uses the native clipboard, not intercepted)
  if (mod && !e.altKey && (e.key === 'v' || e.key === 'V')) {
    if (e.shiftKey) {
      // ⌘⇧V: paste format
      e.preventDefault()
      void clipboardActions.pasteFormat(ctx)
      return
    }
    e.preventDefault()
    void clipboardActions.pasteClipboard(ctx)
    return
  }
  // Esc exits in-group editing, back to whole-group selection
  if (e.key === 'Escape' && ctx.enteredGroupId) {
    e.preventDefault()
    ctx.setSelectedIds([ctx.enteredGroupId])
    ctx.setEnteredGroupId(null)
    return
  }
  // Typing directly on a selected text shape = select-all replace and enter editing
  // (IME start key is 'Process' with length>1 and naturally excluded; typing Chinese directly on selection unsupported for now)
  if (
    !mod &&
    !e.altKey &&
    e.key.length === 1 &&
    selectedIds.length === 1 &&
    document.activeElement === document.body
  ) {
    const ctx0 = ctx.findNodeCtx(selectedIds[0]!)
    const n0 = ctx0?.node
    if (n0 && (n0.type === 'text' || n0.type === 'shape')) {
      e.preventDefault()
      ctx.setEditing({
        sourceId: n0.sourceId,
        replaceWith: e.key,
        ...(ctx0!.groupId ? { groupId: ctx0!.groupId } : {}),
      })
      return
    }
  }
  // Tab / Shift+Tab: cycle shape selection in z-order. Only take over when
  // focus is on the canvas/body, not hijacking focus navigation in ribbon controls etc.
  if (e.key === 'Tab' && !mod && !e.altKey && document.activeElement === document.body) {
    const ids = (slide?.nodes ?? [])
      .filter((n) => !n.decoration && n.type !== 'placeholder-chip')
      .map((n) => n.sourceId)
    if (ids.length) {
      e.preventDefault()
      const cur = selectedIds.length ? ids.indexOf(selectedIds[selectedIds.length - 1]!) : -1
      const next =
        cur < 0
          ? e.shiftKey
            ? ids.length - 1
            : 0
          : e.shiftKey
            ? (cur - 1 + ids.length) % ids.length
            : (cur + 1) % ids.length
      ctx.setSelectedIds([ids[next]!])
    }
    return
  }
  // ⌘A: select all elements on the page (excluding decoration layer/placeholder chips/full-page backgrounds)
  if (mod && !e.altKey && (e.key === 'a' || e.key === 'A')) {
    e.preventDefault()
    const ids = (slide?.nodes ?? [])
      .filter((n) => !n.decoration && !n.background && n.type !== 'placeholder-chip')
      .map((n) => n.sourceId)
    if (ids.length) ctx.setSelectedIds(ids)
    return
  }
  // Nothing selected: ↑/↓ and PageUp/PageDown switch slides (PowerPoint
  // thumbnail-pane behavior; mac keyboards have no physical PageUp/PageDown).
  // defaultPrevented skips events already consumed by the reading view's
  // capture-phase handler; body-focus check keeps ribbon widgets untouched.
  if (selectedIds.length === 0) {
    if (
      !mod &&
      !e.altKey &&
      !e.defaultPrevented &&
      !ctx.masterItems &&
      ctx.slides.length > 0 &&
      document.activeElement === document.body &&
      (e.key === 'ArrowUp' || e.key === 'ArrowDown' || e.key === 'PageUp' || e.key === 'PageDown')
    ) {
      e.preventDefault()
      const delta = e.key === 'ArrowUp' || e.key === 'PageUp' ? -1 : 1
      ctx.setCurrent((c) => Math.max(0, Math.min(c + delta, ctx.slides.length - 1)))
      return
    }
    // ⌘C/⌘X act on the current slide (thumbnail-pane behavior)
    if (mod && !e.altKey && !e.shiftKey && (e.key === 'c' || e.key === 'C')) {
      e.preventDefault()
      void clipboardActions.copySlideAt(ctx, ctx.current)
    } else if (mod && !e.altKey && !e.shiftKey && (e.key === 'x' || e.key === 'X')) {
      e.preventDefault()
      void slideActions.cutSlideAt(ctx, ctx.current)
    }
    return
  }
  if (mod && !e.altKey && (e.key === 'c' || e.key === 'C')) {
    if (e.shiftKey) {
      // ⌘⇧C: copy format
      e.preventDefault()
      clipboardActions.copyFormat(ctx)
      return
    }
    e.preventDefault()
    void clipboardActions.copySelected(ctx)
    return
  }
  if (mod && !e.altKey && (e.key === 'x' || e.key === 'X')) {
    e.preventDefault()
    void clipboardActions.cutSelected(ctx)
    return
  }
  // ⌘D duplicate in place
  if (mod && !e.altKey && (e.key === 'd' || e.key === 'D')) {
    e.preventDefault()
    void clipboardActions.duplicateSelected(ctx, selectedIds)
    return
  }
  // ⌘G group / ⌘⇧G ungroup
  if (mod && !e.altKey && (e.key === 'g' || e.key === 'G')) {
    e.preventDefault()
    if (e.shiftKey) void arrangeActions.ungroupSelected(ctx)
    else void arrangeActions.groupSelected(ctx)
    return
  }
  if (e.key === 'Delete' || e.key === 'Backspace') {
    e.preventDefault()
    void clipboardActions.deleteSelected(ctx)
    return
  }
  const step = e.shiftKey ? 10 : 1
  let dx = 0
  let dy = 0
  if (e.key === 'ArrowLeft') dx = -step
  else if (e.key === 'ArrowRight') dx = step
  else if (e.key === 'ArrowUp') dy = -step
  else if (e.key === 'ArrowDown') dy = step
  else return
  e.preventDefault()
  // Nudge the whole multi-selection; undo granularity is one step per element for now
  for (const id of selectedIds) {
    const nodeCtx = ctx.findNodeCtx(id)
    if (nodeCtx) {
      const b = nodeCtx.node.box
      void ctx.onTransform(
        id,
        { x: b.x + dx, y: b.y + dy, w: b.w, h: b.h, rotationDeg: b.rotationDeg },
        undefined,
        nodeCtx.groupId,
      )
    }
  }
}
