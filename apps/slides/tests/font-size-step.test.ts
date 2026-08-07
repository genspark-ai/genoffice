/**
 * Grow/shrink font while editing: sizes step along the PowerPoint ladder (no ×1.2
 * fractional drift) and the DOM selection survives the span replacement so the
 * highlight stays and repeated clicks keep working.
 */
import { describe, it, expect, beforeAll, afterEach } from 'vitest'
import { resizeSelectionFont, setSelectionFontSizePt } from '../src/renderer/TextEditOverlay'

// jsdom has no execCommand: emulate the browser product — wrap the selection in <font size="7">.
// An element-only selection (what reselectSpans produces) wraps that element's contents, matching
// how Chromium nests the font inside the styled span.
beforeAll(() => {
  document.execCommand = (cmd: string) => {
    if (cmd !== 'fontSize') return true
    const sel = window.getSelection()
    if (!sel?.rangeCount || sel.getRangeAt(0).collapsed) return false
    const range = sel.getRangeAt(0)
    const onlyChild =
      range.startContainer === range.endContainer &&
      range.endOffset - range.startOffset === 1 &&
      range.startContainer.childNodes[range.startOffset]
    const wrap = document.createRange()
    if (onlyChild instanceof HTMLElement) wrap.selectNodeContents(onlyChild)
    else wrap.setStart(range.startContainer, range.startOffset)
    if (!(onlyChild instanceof HTMLElement)) wrap.setEnd(range.endContainer, range.endOffset)
    const font = document.createElement('font')
    font.setAttribute('size', '7')
    wrap.surroundContents(font)
    return true
  }
})

afterEach(() => {
  document.body.innerHTML = ''
  window.getSelection()?.removeAllRanges()
})

function setupEditor(px: number): { root: HTMLElement; span: HTMLElement } {
  const root = document.createElement('div')
  root.tabIndex = 0
  Object.defineProperty(root, 'isContentEditable', { value: true })
  const span = document.createElement('span')
  span.style.fontSize = `${px}px`
  span.textContent = 'Annual timeline'
  root.appendChild(span)
  document.body.appendChild(root)
  root.focus()
  const range = document.createRange()
  range.selectNodeContents(span)
  const sel = window.getSelection()!
  sel.removeAllRanges()
  sel.addRange(range)
  return { root, span }
}

function currentSizePx(root: HTMLElement): number {
  const spans = root.querySelectorAll('span')
  return parseFloat((spans[spans.length - 1] as HTMLElement).style.fontSize)
}

describe('resizeSelectionFont ladder stepping', () => {
  it('grows along the ladder: 45pt → 48 → 54 (not ×1.2)', () => {
    const { root } = setupEditor(60) // 60px = 45pt
    resizeSelectionFont(1)
    expect(currentSizePx(root)).toBe(64) // 48pt
    resizeSelectionFont(1)
    expect(currentSizePx(root)).toBe(72) // 54pt
  })

  it('shrinks along the ladder: 45pt → 44', () => {
    const { root } = setupEditor(60)
    resizeSelectionFont(-1)
    expect(currentSizePx(root)).toBe((44 * 96) / 72)
  })

  it('steps ±10pt beyond the ladder and clamps at 8pt', () => {
    const { root } = setupEditor((100 * 96) / 72)
    resizeSelectionFont(1)
    expect(currentSizePx(root)).toBe((110 * 96) / 72)

    const { root: small } = setupEditor((8 * 96) / 72)
    resizeSelectionFont(-1)
    expect(currentSizePx(small)).toBe((8 * 96) / 72)
  })

  it('keeps the text selected after each step', () => {
    const { root } = setupEditor(60)
    resizeSelectionFont(1)
    const sel = window.getSelection()!
    expect(sel.isCollapsed).toBe(false)
    expect(sel.toString()).toBe('Annual timeline')
    expect(root.contains(sel.getRangeAt(0).commonAncestorContainer)).toBe(true)
  })
})

describe('setSelectionFontSizePt', () => {
  it('applies the absolute size and keeps the selection', () => {
    const { root } = setupEditor(60)
    setSelectionFontSizePt(28)
    expect(currentSizePx(root)).toBe((28 * 96) / 72)
    const sel = window.getSelection()!
    expect(sel.isCollapsed).toBe(false)
    expect(sel.toString()).toBe('Annual timeline')
  })
})
