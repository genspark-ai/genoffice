import { t } from '../i18n/locale'

/** Apply screen-only previews after the card has its final width and is in the DOM. */
export function setCommentPresentation(
  bubble: HTMLElement,
  lines: 15 | 5 | 3 | 1,
  expanded: boolean,
  maxHeight: number,
): number {
  const texts = [...bubble.querySelectorAll<HTMLElement>('.comment-text')]
  for (const text of texts) text.style.removeProperty('max-height')
  const root = texts[0]
  const rootLineHeight = root ? parseFloat(getComputedStyle(root).lineHeight) : 0
  const previewLines =
    lines === 15 && rootLineHeight > 0 && root.scrollHeight > rootLineHeight * 15 + 1 ? 5 : lines
  let truncated = false
  for (const text of texts) {
    const lineHeight = parseFloat(getComputedStyle(text).lineHeight)
    const limit = previewLines === 15 && text.scrollHeight > lineHeight * 15 + 1 ? 5 : previewLines
    if (!expanded && lineHeight > 0 && text.scrollHeight > lineHeight * limit + 1) {
      text.style.maxHeight = `${lineHeight * limit}px`
      truncated = true
    }
  }
  bubble.classList.toggle('comment-expanded', expanded)
  bubble.dataset.commentLines = String(lines)
  if (expanded && maxHeight > 0) bubble.style.maxHeight = `${maxHeight}px`
  else bubble.style.removeProperty('max-height')
  const toggle = bubble.querySelector<HTMLButtonElement>('.comment-reading-toggle')
  if (toggle) {
    toggle.hidden = !expanded && !truncated
    toggle.textContent = expanded ? '⌃' : '…'
    toggle.setAttribute('aria-label', t(expanded ? 'appNavCollapse' : 'appNavExpand'))
    toggle.setAttribute('aria-expanded', String(expanded))
    toggle.title = t(expanded ? 'appNavCollapse' : 'appNavExpand')
  }
  return bubble.offsetHeight
}
