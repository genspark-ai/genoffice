/**
 * Global keyboard shortcuts extracted from App.tsx: ⌘Z undo, ⌘F
 * find, zoom, element clipboard, format painter, selection cycling, nudging…
 * The handler reads the latest App state through ActionCtx, so App attaches it
 * once with an empty dependency list.
 */
import { deleteSelectedVertex } from './edit-points-actions'
import type { FontSizeStep } from '@genoffice/pptx-ops/font-size'
import type { ActionCtx, EditingState } from './action-context'
import type { EditTransformMultiOp } from '../shared/ipc'
import { FIT_WIDTH, NUDGE_STEP_PX } from './app-constants'
import { isConnectorNode, isEditableText } from './konva-adapter'
import * as clipboardActions from './clipboard-actions'
import * as arrangeActions from './arrange-actions'
import * as slideActions from './slide-actions'
import * as showActions from './show-actions'
import * as styleActions from './style-actions'
import { flushActiveEdit } from './file-actions'
import { shouldRouteUndoToDeck } from './undo-routing'
import { rangeSelection } from '../shared/slide-selection'
import { nextPreset, prevPreset } from './zoom-steps'

/** Whether focus is in a text input (input/textarea/contentEditable) — these cases use native undo/delete */
function inTextField(): boolean {
  const el = document.activeElement as HTMLElement | null
  return !!el && (el.tagName === 'INPUT' || el.tagName === 'TEXTAREA' || el.isContentEditable)
}

/** The text edit overlay (shape text or table cell) is the one contentEditable the format shortcuts act in */
function inEditorOverlay(): boolean {
  return !!document.activeElement?.closest('.slide-text-editor')
}

const ALIGN_KEYS: Record<string, 'left' | 'center' | 'right' | 'justify'> = {
  l: 'left',
  e: 'center',
  r: 'right',
  j: 'justify',
}

const TOGGLE_KEYS: Record<string, 'bold' | 'italic' | 'underline'> = {
  b: 'bold',
  i: 'italic',
  u: 'underline',
}

/** ⌘]/⌘[ = ±1 pt; ⇧⌘> / ⇧⌘< = next/previous ladder rung (key is layout-resolved, so '.'/',' cover keyboards where Shift yields no '>'/'<') */
function fontSizeStepFor(e: KeyboardEvent): FontSizeStep | null {
  if (!e.shiftKey) {
    if (e.key === ']') return { dir: 1, mode: 'point' }
    if (e.key === '[') return { dir: -1, mode: 'point' }
    return null
  }
  if (e.key === '>' || e.key === '.') return { dir: 1, mode: 'ladder' }
  if (e.key === '<' || e.key === ',') return { dir: -1, mode: 'ladder' }
  return null
}

/** PowerPoint deletes/cuts/copies whole slides only while the thumbnail pane, outline pane or sorter has keyboard focus */
export function slideRailHasFocus(): boolean {
  return !!document.activeElement?.closest('.slide-list, .outline-pane, .sorter-view')
}

/**
 * How a key pressed on a lone selected text shape opens the edit session (PowerPoint):
 * a printable character replaces the whole text, Enter leaves the caret after it, F2
 * selects all of it. IME start key is 'Process' with length>1 and naturally excluded.
 */
function editEntry(e: KeyboardEvent, ctx: ActionCtx): Partial<EditingState> | null {
  if (e.metaKey || e.ctrlKey || e.altKey) return null
  if (e.key.length === 1) return { replaceWith: e.key }
  if (e.key !== 'Enter' && e.key !== 'F2') return null
  if (e.shiftKey || e.defaultPrevented || ctx.masterItems || ctx.viewMode === 'reading') return null
  return e.key === 'F2' ? { selectAll: true } : {}
}

/** One Shift+arrow step on a side; shrinking never drops below one step (or the side's current size if already thinner) */
function resizeBy(size: number, dir: number): number {
  return dir > 0
    ? size + NUDGE_STEP_PX
    : Math.max(Math.min(size, NUDGE_STEP_PX), size - NUDGE_STEP_PX)
}

// Shortcuts: ⌘/Ctrl+Z undo, ⇧⌘Z / Ctrl+Y redo, ⌘/Ctrl+=/-/0 zoom,
// Delete/Backspace delete selection, arrow keys nudge selection
export function handleGlobalKeydown(
  ctx: ActionCtx,
  e: KeyboardEvent,
  platform = navigator.platform,
): void {
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
  // PowerPoint for macOS: ⌘+Enter starts from the current slide. Keep
  // Ctrl+Enter untouched on Windows, where PowerPoint uses it to move between
  // placeholders (and Shift+F5 already starts from the current slide).
  if (
    /mac/i.test(platform) &&
    e.metaKey &&
    !e.ctrlKey &&
    !e.altKey &&
    !e.shiftKey &&
    e.key === 'Enter'
  ) {
    if (e.defaultPrevented || editing || inField || ctx.cropTarget) return
    e.preventDefault()
    showActions.startSlideShow(ctx, false)
    return
  }
  // PowerPoint's new-slide chord: Ctrl+M on Windows, ⇧⌘N on mac. The renderer sees ⌘M
  // before the menu accelerator (measured), so it must stay untouched there for Minimize.
  // Commits an open text edit and the notes draft first, in save()'s order.
  const mac = /mac/i.test(platform)
  if (
    mod &&
    !e.altKey &&
    ((!mac && !e.shiftKey && (e.key === 'm' || e.key === 'M')) ||
      (mac && e.metaKey && e.shiftKey && (e.key === 'n' || e.key === 'N')))
  ) {
    if (e.defaultPrevented || ctx.viewMode === 'reading' || ctx.masterItems || ctx.cropTarget)
      return
    e.preventDefault()
    void flushActiveEdit(ctx)
      .then(() => ctx.flushNotes())
      .then(() => slideActions.addSlide(ctx))
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
  // ⌘K: annotate the selection with an AI edit
  if (mod && !e.altKey && !e.shiftKey && (e.key === 'k' || e.key === 'K')) {
    if (editing || inField || selectedIds.length === 0) return
    e.preventDefault()
    ctx.openAskPopover()
    return
  }
  // ⌘P print
  if (mod && !e.altKey && !e.shiftKey && (e.key === 'p' || e.key === 'P')) {
    if (editing || inField) return
    e.preventDefault()
    ctx.setPrintDlgOpen(true)
    return
  }
  // PowerPoint text formatting: ⌘L/E/R/J align, ⌘]/⌘[ ±1 pt, ⇧⌘>/⇧⌘< grow/shrink
  // font, ⌘B/I/U. Inside the edit overlay they hit the selection (B/I/U stay
  // native there); with shapes selected they change the whole shapes; any other
  // field keeps its native keys.
  if (
    mod &&
    !e.altKey &&
    !e.defaultPrevented &&
    !ctx.masterItems &&
    ctx.viewMode !== 'reading' &&
    !ctx.cropTarget
  ) {
    const inEditor = !!(editing || ctx.editingCell) && inEditorOverlay()
    if (inEditor || (!inField && selectedIds.length > 0)) {
      const lower = e.key.toLowerCase()
      const align = e.shiftKey ? undefined : ALIGN_KEYS[lower]
      if (align) {
        e.preventDefault()
        styleActions.onAlign(ctx, align)
        return
      }
      const step = fontSizeStepFor(e)
      if (step) {
        e.preventDefault()
        styleActions.onFontSizeStep(ctx, step)
        return
      }
      const toggle = e.shiftKey ? undefined : TOGGLE_KEYS[lower]
      if (toggle && !inEditor) {
        e.preventDefault()
        styleActions.onTextToggle(ctx, toggle)
        return
      }
    }
  }
  // Zoom shortcut fallback
  if (mod && !inField && (e.key === '=' || e.key === '+')) {
    e.preventDefault()
    ctx.setZoom(nextPreset)
    return
  }
  if (mod && !inField && e.key === '-') {
    e.preventDefault()
    ctx.setZoom(prevPreset)
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
  // An open overlay (menu, dialog, popover, crop, media, draw mode) owns Escape outright
  if (e.key === 'Escape' && ctx.overlayOpen) return
  // Esc: drop the ink pen/highlighter/eraser back to the select tool (PowerPoint behavior)
  if (e.key === 'Escape' && ctx.inkTool !== 'select') {
    e.preventDefault()
    ctx.setInkTool('select')
    return
  }
  // Esc: exit format painter continuous mode
  if (e.key === 'Escape' && ctx.brushMode) {
    e.preventDefault()
    ctx.setBrushMode(null)
    ctx.setStatus('')
    return
  }
  if (e.key === 'Escape' && ctx.editPointsTarget) {
    e.preventDefault()
    ctx.setEditPointsTarget(null)
    return
  }
  // Edit Points: Delete removes the selected vertex, not the shape
  if ((e.key === 'Delete' || e.key === 'Backspace') && ctx.editPointsTarget?.vertex != null) {
    e.preventDefault()
    deleteSelectedVertex(ctx)
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
  // Esc deselects (PowerPoint)
  if (e.key === 'Escape' && selectedIds.length > 0) {
    e.preventDefault()
    ctx.setSelectedIds([])
    return
  }
  // Typing / Enter / F2 directly on a selected text shape enters editing
  if (selectedIds.length === 1 && document.activeElement === document.body) {
    const entry = editEntry(e, ctx)
    const ctx0 = entry && ctx.findNodeCtx(selectedIds[0]!)
    const n0 = ctx0?.node
    if (entry && n0 && isEditableText(n0)) {
      e.preventDefault()
      ctx.setEditing({
        sourceId: n0.sourceId,
        ...entry,
        ...(ctx0.groupId ? { groupId: ctx0.groupId } : {}),
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
  // ⌘A: every slide with the rail/sorter focused (PowerPoint), otherwise every
  // element on the page (excluding decoration layer/placeholder chips/full-page backgrounds)
  if (mod && !e.altKey && (e.key === 'a' || e.key === 'A')) {
    e.preventDefault()
    if (slideRailHasFocus() && ctx.slides.length > 0) {
      ctx.setSelectedSlides(rangeSelection(0, ctx.slides.length - 1))
      return
    }
    const ids = (slide?.nodes ?? [])
      .filter((n) => !n.decoration && !n.background && n.type !== 'placeholder-chip')
      .map((n) => n.sourceId)
    if (ids.length) ctx.setSelectedIds(ids)
    return
  }
  // Nothing selected: ↑/↓ and PageUp/PageDown switch slides from either the
  // canvas (body focus) or the thumbnail rail (PowerPoint behavior; mac
  // keyboards have no physical PageUp/PageDown). defaultPrevented skips events
  // already consumed by the reading view's capture-phase handler; the focus
  // check keeps ribbon widgets untouched.
  if (selectedIds.length === 0) {
    const railFocus = slideRailHasFocus()
    if (
      !mod &&
      !e.altKey &&
      !e.defaultPrevented &&
      !ctx.masterItems &&
      ctx.slides.length > 0 &&
      (railFocus || document.activeElement === document.body) &&
      (e.key === 'ArrowUp' || e.key === 'ArrowDown' || e.key === 'PageUp' || e.key === 'PageDown')
    ) {
      e.preventDefault()
      const delta = e.key === 'ArrowUp' || e.key === 'PageUp' ? -1 : 1
      const next = Math.max(0, Math.min(ctx.current + delta, ctx.slides.length - 1))
      ctx.setCurrent(next)
      ctx.setSelectedSlides([next])
      return
    }
    // Slide-level Delete/Backspace/⌘C/⌘X/Home/End only with the rail or sorter focused:
    // deselecting on the canvas and pressing Delete must not drop the slide.
    if (!railFocus) return
    // Home/End jump to the first/last slide; with Shift the range grows from the anchor (PowerPoint)
    if (
      !mod &&
      !e.altKey &&
      !e.defaultPrevented &&
      !ctx.masterItems &&
      ctx.slides.length > 0 &&
      (e.key === 'Home' || e.key === 'End')
    ) {
      e.preventDefault()
      const edge = e.key === 'Home' ? 0 : ctx.slides.length - 1
      if (e.shiftKey) {
        ctx.setSelectedSlides(rangeSelection(ctx.current, edge))
      } else {
        ctx.setCurrent(edge)
        ctx.setSelectedSlides([edge])
      }
      return
    }
    // Delete/Backspace removes the selected slides (same action as the thumbnail
    // context menu; deleteSlides keeps ≥1 slide). Not while inking or in
    // reading view (both clear the selection), and not with plain-DOM text
    // dragged (AI panel): the key targets that text.
    if (
      !mod &&
      !e.altKey &&
      !e.defaultPrevented &&
      !ctx.masterItems &&
      ctx.inkTool === 'select' &&
      ctx.viewMode !== 'reading' &&
      ctx.slides.length > 0 &&
      (window.getSelection()?.isCollapsed ?? true) &&
      (e.key === 'Delete' || e.key === 'Backspace')
    ) {
      e.preventDefault()
      void slideActions.deleteSlides(ctx, ctx.selectedSlides)
      return
    }
    if (mod && !e.altKey && !e.shiftKey && (e.key === 'c' || e.key === 'C')) {
      e.preventDefault()
      void clipboardActions.copySlides(ctx, ctx.selectedSlides)
    } else if (mod && !e.altKey && !e.shiftKey && (e.key === 'x' || e.key === 'X')) {
      e.preventDefault()
      void slideActions.cutSlides(ctx, ctx.selectedSlides)
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
  if (e.altKey) return
  let ux = 0
  let uy = 0
  if (e.key === 'ArrowLeft') ux = -1
  else if (e.key === 'ArrowRight') ux = 1
  else if (e.key === 'ArrowUp') uy = -1
  else if (e.key === 'ArrowDown') uy = 1
  else return
  e.preventDefault()
  // PowerPoint tiers: plain arrow = one grid step, ⌘/Ctrl = one screen pixel,
  // Shift = resize about the center; a multi-selection commits as one undo step.
  const step = mod ? 1 / (ctx.zoom || 1) : NUDGE_STEP_PX
  const items: EditTransformMultiOp['items'] = []
  for (const id of selectedIds) {
    const nodeCtx = ctx.findNodeCtx(id)
    if (!nodeCtx) continue
    const b = nodeCtx.node.box
    const box = { x: b.x, y: b.y, w: b.w, h: b.h, rotationDeg: b.rotationDeg }
    if (e.shiftKey && !mod) {
      if (isConnectorNode(nodeCtx.node)) continue
      if (ux) {
        box.w = resizeBy(b.w, ux)
        box.x = b.x + (b.w - box.w) / 2
      } else {
        box.h = resizeBy(b.h, -uy)
        box.y = b.y + (b.h - box.h) / 2
      }
    } else {
      box.x = b.x + ux * step
      box.y = b.y + uy * step
    }
    items.push({
      sourceId: id,
      xPx: box.x,
      yPx: box.y,
      wPx: box.w,
      hPx: box.h,
      rotationDeg: box.rotationDeg,
      ...(nodeCtx.groupId ? { groupId: nodeCtx.groupId } : {}),
    })
  }
  void commitTransforms(ctx, items)
}

/** One element keeps the per-element transform path; several go out as a single batch. */
async function commitTransforms(
  ctx: ActionCtx,
  items: EditTransformMultiOp['items'],
): Promise<void> {
  if (!items.length) return
  if (items.length === 1) {
    const it = items[0]!
    await ctx.onTransform(
      it.sourceId,
      { x: it.xPx, y: it.yPx, w: it.wPx, h: it.hPx, rotationDeg: it.rotationDeg },
      undefined,
      it.groupId,
    )
    return
  }
  const updated = await window.slidesApi.editTransformMulti({
    slideIndex: ctx.current,
    fitWidthPx: FIT_WIDTH,
    items,
  })
  if (updated) ctx.applySlide(ctx.current, updated)
}
