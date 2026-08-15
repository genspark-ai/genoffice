/** What a plain (no-modifier) navigation key should do in the viewer */
export type NavAction =
  | { type: 'scrollBy'; delta: number }
  | { type: 'scrollViewport'; dir: 1 | -1 }
  | { type: 'scrollEdge'; edge: 'top' | 'bottom' }
  | { type: 'stepPage'; dir: 1 | -1 }

/** Vertical arrows step whole pages while focus sits in the thumbnail sidebar,
    and scroll the document otherwise */
export function navAction(key: string, focusInThumbs: boolean): NavAction | null {
  switch (key) {
    case 'PageDown':
    case ' ':
      return { type: 'scrollViewport', dir: 1 }
    case 'PageUp':
      return { type: 'scrollViewport', dir: -1 }
    case 'Home':
      return { type: 'scrollEdge', edge: 'top' }
    case 'End':
      return { type: 'scrollEdge', edge: 'bottom' }
    case 'ArrowDown':
      return focusInThumbs ? { type: 'stepPage', dir: 1 } : { type: 'scrollBy', delta: 60 }
    case 'ArrowUp':
      return focusInThumbs ? { type: 'stepPage', dir: -1 } : { type: 'scrollBy', delta: -60 }
    case 'ArrowRight':
      return { type: 'stepPage', dir: 1 }
    case 'ArrowLeft':
      return { type: 'stepPage', dir: -1 }
    default:
      return null
  }
}
