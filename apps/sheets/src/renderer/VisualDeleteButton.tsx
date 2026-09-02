import { useEffect, useRef, type RefObject } from 'react'
import { createPortal } from 'react-dom'

import {
  type BoxRect,
  type GridInsetSource,
  VISUAL_DELETE_BUTTON_GAP,
  VISUAL_DELETE_BUTTON_SIZE,
  gridTopInset,
  visualDeleteButtonPosition,
} from './visual-delete-button'

/// Univer's float-DOM wrapper is the nearest ancestor that clips through an
/// inline style; its parent is the sheet canvas area shared by every float.
function floatDomWrapperOf(host: HTMLElement): HTMLElement | null {
  for (let node = host.parentElement; node; node = node.parentElement) {
    if (node.style.overflow === 'hidden') return node
  }
  return null
}

function gridBounds(grid: HTMLElement | null, topInset: number): BoxRect {
  const rect = grid?.getBoundingClientRect()
  return rect
    ? { left: rect.left, top: rect.top + topInset, right: rect.right, bottom: rect.bottom }
    : { left: 0, top: topInset, right: window.innerWidth, bottom: window.innerHeight }
}

interface Props {
  /// The `.shape-editable` host; the button follows its float-DOM frame.
  readonly hostRef: RefObject<HTMLElement | null>
  readonly worksheet: GridInsetSource
  readonly label: string
  readonly onDelete: () => void
}

/**
 * Body-level ✕ for the selected visual. Univer's float-DOM wrapper clips its
 * children, so a button outside the frame cannot live inside the host: it is
 * portaled to <body>, positioned from the wrapper's client rect (the visible,
 * viewport-clamped frame) and re-measured whenever Univer moves the wrapper
 * (scroll, zoom and row/column resizes all rewrite its inline style), the
 * host or the grid changes size, or the window resizes. Position is written
 * straight to the element so it keeps up with the scroll instead of trailing
 * by a React render.
 */
export function VisualDeleteButton({
  hostRef,
  worksheet,
  label,
  onDelete,
}: Props): React.JSX.Element {
  const buttonRef = useRef<HTMLButtonElement>(null)

  // A passive effect, not a layout effect: the host attaches its ref in the
  // layout phase AFTER its children's layout effects (and re-attaches it on
  // every render), so a layout effect here would still see a null host. The
  // button stays hidden until the first measure, so nothing flashes.
  useEffect(() => {
    const host = hostRef.current
    const button = buttonRef.current
    if (!host || !button) return
    const wrapper = floatDomWrapperOf(host)
    const grid = wrapper?.parentElement ?? null
    const measure = (): void => {
      const frame = (wrapper ?? host).getBoundingClientRect()
      // The float container may lay out a beat after it mounts; the host
      // ResizeObserver brings the button back once the frame has a size.
      if (frame.width <= 0 || frame.height <= 0) {
        button.dataset['ready'] = 'false'
        return
      }
      const position = visualDeleteButtonPosition(
        frame,
        gridBounds(grid, gridTopInset(worksheet)),
        VISUAL_DELETE_BUTTON_SIZE,
        VISUAL_DELETE_BUTTON_GAP,
      )
      button.style.left = `${position.left}px`
      button.style.top = `${position.top}px`
      button.dataset['placement'] = position.placement
      button.dataset['ready'] = 'true'
    }
    measure()
    const wrapperMoves = new MutationObserver(measure)
    if (wrapper) wrapperMoves.observe(wrapper, { attributes: true, attributeFilter: ['style'] })
    const resizes = new ResizeObserver(measure)
    resizes.observe(host)
    if (grid) resizes.observe(grid)
    window.addEventListener('resize', measure)
    return () => {
      wrapperMoves.disconnect()
      resizes.disconnect()
      window.removeEventListener('resize', measure)
    }
  }, [hostRef, worksheet])

  return createPortal(
    <button
      ref={buttonRef}
      type="button"
      className="shape-delete-button"
      tabIndex={-1}
      data-tip={label}
      aria-label={label}
      // Keep focus (and with it Delete/Backspace) on the visual's host.
      onMouseDown={(event) => event.preventDefault()}
      // Portal events still bubble through the React tree to the host, whose
      // pointerdown starts a move/resize drag — stop them here.
      onPointerDown={(event) => event.stopPropagation()}
      onClick={(event) => {
        event.stopPropagation()
        onDelete()
      }}
    >
      ✕
    </button>,
    document.body,
  )
}
