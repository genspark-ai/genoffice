/** Show Comments ▾: Word's Contextual (cards beside the page) or List (the Comments panel) */
export type CommentsMode = 'contextual' | 'list'

const KEY = 'aidocs.commentsMode'

export function loadCommentsMode(): CommentsMode {
  return localStorage.getItem(KEY) === 'list' ? 'list' : 'contextual'
}

export function saveCommentsMode(mode: CommentsMode): void {
  localStorage.setItem(KEY, mode)
}

export interface CommentsView {
  /** contextual cards in the markup column beside the page */
  cards: boolean
  /** the Comments panel on the right edge */
  panel: boolean
}

/**
 * What Show Comments displays. Contextual puts cards beside the page where the
 * margin can host them (not Simple Markup, inline balloons or a non-print view);
 * the panel also opens there while a composer or a thread focus needs a host,
 * and stays until closed.
 */
export function commentsView(
  show: boolean,
  mode: CommentsMode,
  cardsPossible: boolean,
  panelOpen: boolean,
): CommentsView {
  if (!show) return { cards: false, panel: false }
  if (mode === 'list') return { cards: false, panel: true }
  return { cards: cardsPossible, panel: panelOpen }
}

/** anchor-icon click: select the hosted card / panel entry, else pop the thread up beside the page */
export function iconClickTarget(view: CommentsView): 'focus' | 'popup' {
  return view.cards || view.panel ? 'focus' : 'popup'
}
