/**
 * Fragment links on the export path.
 *
 * Chromium resolves `href="#name"` while printing by looking up the element
 * that holds that id, and takes the first match in DOM order. If that element
 * has no layout box the PDF gets the /Link annotation but no /Dests entry, so
 * the link resolves to nothing. The pagination preview is the print source and
 * holds one full-document clone per page under `.pv-clip { overflow: hidden }`,
 * so every name is copied once per page and the print stylesheet hides the
 * editor canvas ahead of all of them: the winner is always the hidden, box-less
 * canvas copy. Hence one id per name per print, on the block inside the clone of
 * the page that owns it.
 *
 * Ownership is read back from the layout rather than recomputed from the page
 * slices, so pruned and full clones, multi-column regions and repeated table
 * headers all answer the same question — "which clip window shows this anchor".
 *
 * Ids are re-derived from data-bookmark on every call, never tracked
 * incrementally: a second export must not find a name still stranded on the
 * page that lost it.
 */

/** Anchors the bookmark-targets plugin renders, whatever id they carry now. */
const ANCHORS = 'a.doc-bookmark-target[data-bookmark]'

export function assignPreviewBookmarkTargets(root: ParentNode = document): void {
  const pages = [...root.querySelectorAll<HTMLElement>('.pv-page')]
  const anchors = [...root.querySelectorAll<HTMLElement>(ANCHORS)]
  // no preview: the editor canvas is itself the print source and its one copy
  // per name is already correct
  if (pages.length === 0 || anchors.length === 0) return

  for (const a of anchors) a.removeAttribute('id')

  for (const pageEl of pages) {
    for (const clip of pageEl.querySelectorAll<HTMLElement>('.pv-clip')) {
      const window = clip.getBoundingClientRect()
      // a collapsed window is a page hidden for a print chunk (pv-print-skip)
      if (window.height <= 0) continue
      for (const a of clip.querySelectorAll<HTMLElement>(ANCHORS)) {
        const name = a.dataset.bookmark
        // an earlier page already claimed this name; only one id may exist
        if (!name || a.id) continue
        const box = a.getBoundingClientRect()
        // an empty inline anchor keeps its line box; a display:none one
        // collapses to 0x0 at the origin, so geometry alone tells the two apart
        if (box.width <= 0 && box.height <= 0) continue
        // the anchor only prints where its clone's clip window shows it
        if (box.bottom <= window.top || box.top >= window.bottom) continue
        a.id = name
      }
    }
  }
}
