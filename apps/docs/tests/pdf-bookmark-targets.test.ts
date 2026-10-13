import { afterEach, describe, expect, it, vi } from 'vitest'
import { assignPreviewBookmarkTargets } from '../src/renderer/pdf-bookmark-targets'
import { exportPdf } from '../src/renderer/file-actions'

/**
 * Fragment links on the export path: Chromium resolves `href="#name"` against
 * the element holding that id, and when the first match in DOM order has no
 * layout box it emits the /Link annotation but no /Dests entry at all. The
 * pagination preview holds one clone of the document per page and print CSS
 * hides the editor canvas ahead of them, so every name has many copies and the
 * first one is always the hidden canvas. Export therefore moves each id onto the
 * single copy that prints on the page that owns it.
 *
 * jsdom has no layout engine, so the anchors and clips declare their rects.
 */

type Rect = {
  top: number
  bottom: number
  left: number
  right: number
  width: number
  height: number
}

function rect(top: number, height: number, left = 0, width = 400): Rect {
  return { top, bottom: top + height, left, right: left + width, width, height }
}

/** the collapsed rect a display:none element reports */
const NOT_LAID_OUT = rect(0, 0)

function stubRect(el: Element, r: Rect): void {
  ;(el as HTMLElement).getBoundingClientRect = () => r as DOMRect
}

/** one preview page: a clip window with the anchors that fall inside it */
function page(index: number, windowTop: number, windowHeight: number): HTMLElement {
  const el = document.createElement('div')
  el.className = 'pv-page'
  el.dataset.pvPage = String(index)
  const clip = document.createElement('div')
  clip.className = 'pv-clip'
  stubRect(clip, rect(windowTop, windowHeight))
  el.append(clip)
  document.body.append(el)
  return el
}

/** an anchor copy inside `parent`, laid out at `r` (default: collapsed) */
function anchor(parent: Element, name: string, r: Rect = NOT_LAID_OUT): HTMLElement {
  const a = document.createElement('a')
  a.className = 'doc-bookmark-target'
  a.dataset.bookmark = name
  a.id = name
  stubRect(a, r)
  parent.append(a)
  return a
}

function ids(root: ParentNode = document): string[] {
  return [...root.querySelectorAll<HTMLElement>('a.doc-bookmark-target[id]')].map((a) => a.id)
}

/** re-declare layout on anchors rebuilt by an innerHTML rewrite */
function stubAnchors(parent: Element, rects: Rect[]): void {
  const anchors = [...parent.querySelectorAll('a.doc-bookmark-target')]
  anchors.forEach((a, i) => {
    if (rects[i]) stubRect(a, rects[i])
  })
}

afterEach(() => {
  document.body.innerHTML = ''
})

describe('assignPreviewBookmarkTargets', () => {
  describe('the preview is the print source', () => {
    it('puts each id on exactly one copy: the one on the page that owns it', () => {
      const p1 = page(0, 0, 200)
      const p2 = page(1, 200, 200)
      const p3 = page(2, 400, 200)
      const clip1 = p1.querySelector('.pv-clip')!
      const clip2 = p2.querySelector('.pv-clip')!
      const clip3 = p3.querySelector('.pv-clip')!
      // every page clone holds the whole document (small-document preview):
      // the clones are aligned windows onto one flow, so an anchor sits at the
      // same absolute Y in each of them and only the clip window differs
      anchor(clip1, 'chap01', rect(10, 20))
      anchor(clip2, 'chap01', rect(10, 20))
      anchor(clip3, 'chap01', rect(10, 20))
      anchor(clip1, 'chap02', rect(250, 20)) // below page 1's window
      anchor(clip2, 'chap02', rect(250, 20))
      anchor(clip3, 'chap02', rect(250, 20))

      assignPreviewBookmarkTargets()

      // one id each, both on the page whose window shows them
      expect(ids()).toEqual(['chap01', 'chap02'])
      expect(clip1.querySelector('#chap01')).not.toBeNull()
      expect(clip2.querySelector('#chap02')).not.toBeNull()
      expect(clip1.querySelector('#chap02')).toBeNull()
      expect(clip3.querySelector('#chap02')).toBeNull()
    })

    it('strips the id from the hidden editor canvas, whose copy would win in DOM order', () => {
      const canvas = document.createElement('div')
      canvas.className = 'doc-editor'
      anchor(canvas, 'chap01', NOT_LAID_OUT)
      document.body.append(canvas)
      const clip = page(0, 0, 200).querySelector('.pv-clip')!
      anchor(clip, 'chap01', rect(10, 20))

      assignPreviewBookmarkTargets()

      expect(canvas.querySelector('a')!.hasAttribute('id')).toBe(false)
      expect(document.querySelectorAll('#chap01')).toHaveLength(1)
    })

    it('leaves a name with no copy inside any page window without an id', () => {
      const clip = page(0, 0, 200).querySelector('.pv-clip')!
      anchor(clip, 'chap01', rect(900, 20)) // beyond the page: clipped away
      anchor(clip, 'chap02', rect(10, 20))

      assignPreviewBookmarkTargets()

      expect(ids()).toEqual(['chap02'])
    })

    it('re-derives from scratch on every call, so a second export stays correct', () => {
      const p1 = page(0, 0, 200)
      const p2 = page(1, 200, 200)
      anchor(p1.querySelector('.pv-clip')!, 'chap01', rect(10, 20))
      anchor(p2.querySelector('.pv-clip')!, 'chap01', rect(10, 20))

      assignPreviewBookmarkTargets()
      expect(p1.querySelector('#chap01')).not.toBeNull()
      expect(p2.querySelector('#chap01')).toBeNull()

      // a second run must not strand the name on the page that lost it
      assignPreviewBookmarkTargets()
      expect(ids()).toEqual(['chap01'])
      expect(p1.querySelector('#chap01')).not.toBeNull()
    })

    it('repairs the clones after React rewrites them from the cached page HTML', () => {
      // This is the failure the preview has to survive on every commit: each
      // sheet's innerHTML is re-assigned from the same cached document HTML, so
      // every anchor regains its id on every page and Chromium resolves the
      // fragment against the first copy again.
      const p1 = page(0, 0, 200)
      const p2 = page(1, 200, 200)
      const clip1 = p1.querySelector('.pv-clip')!
      const clip2 = p2.querySelector('.pv-clip')!
      anchor(clip1, 'chap01', rect(10, 20))
      anchor(clip2, 'chap01', rect(10, 20))
      const cached = clip1.innerHTML

      assignPreviewBookmarkTargets()
      expect(ids()).toEqual(['chap01'])

      // the re-render replaces the nodes: ids back everywhere, rects re-stubbed
      // (a fresh element reports no layout in jsdom until it says otherwise)
      clip1.innerHTML = cached
      clip2.innerHTML = cached
      stubAnchors(clip1, [rect(10, 20)])
      stubAnchors(clip2, [rect(10, 20)])
      expect(document.querySelectorAll('#chap01')).toHaveLength(2)

      assignPreviewBookmarkTargets()
      expect(ids()).toEqual(['chap01'])
      expect(clip1.querySelector('#chap01')).not.toBeNull()
      expect(clip2.querySelector('#chap01')).toBeNull()
    })

    it('skips a page whose window never got laid out, without stranding other names', () => {
      const broken = page(0, 0, 0) // pv-print-skip: collapsed while its chunk prints
      const p2 = page(1, 200, 200)
      anchor(broken.querySelector('.pv-clip')!, 'chap01', rect(10, 20))
      anchor(p2.querySelector('.pv-clip')!, 'chap02', rect(250, 20))

      assignPreviewBookmarkTargets()

      expect(ids()).toEqual(['chap02'])
      expect(broken.querySelector('#chap01')).toBeNull()
      expect(p2.querySelector('#chap02')).not.toBeNull()
    })

    it('claims a name on the first page that shows it', () => {
      const p1 = page(0, 0, 200)
      const p2 = page(1, 200, 200)
      // a paragraph straddling the break shows its anchor line on both pages
      anchor(p1.querySelector('.pv-clip')!, 'chap01', rect(190, 20))
      anchor(p2.querySelector('.pv-clip')!, 'chap01', rect(190, 20))

      assignPreviewBookmarkTargets()

      expect(p1.querySelector('#chap01')).not.toBeNull()
      expect(p2.querySelector('#chap01')).toBeNull()
    })
  })

  it('leaves the canvas untouched when there is no preview (direct canvas print)', () => {
    const canvas = document.createElement('div')
    anchor(canvas, 'chap01', rect(10, 20))
    anchor(canvas, 'chap02', rect(40, 20))
    document.body.append(canvas)

    assignPreviewBookmarkTargets()

    expect(ids()).toEqual(['chap01', 'chap02'])
  })

  it('does nothing when the document has no bookmark targets', () => {
    page(0, 0, 200)
    expect(() => assignPreviewBookmarkTargets()).not.toThrow()
    expect(ids()).toEqual([])
  })

  // The unit tests above cover the rule; this one covers the call site, which
  // is what decides whether an exported PDF resolves at all.
  it('exportPdf has already moved the ids when it hands the pages to Chromium', async () => {
    const p1 = page(0, 0, 200)
    const p2 = page(1, 200, 200)
    for (const p of [p1, p2]) {
      p.style.width = '400px'
      p.style.height = '200px'
    }
    const clip1 = p1.querySelector('.pv-clip')!
    const clip2 = p2.querySelector('.pv-clip')!
    anchor(clip1, 'chap01', rect(10, 20))
    anchor(clip2, 'chap01', rect(10, 20))
    anchor(clip1, 'chap02', rect(250, 20))
    anchor(clip2, 'chap02', rect(250, 20))

    let atPrintTime: string[] = []
    const print = vi.fn(async () => {
      // exactly the DOM state Chromium would resolve href="#name" against
      atPrintTime = ids()
      return { ok: true, path: '/tmp/out.pdf' }
    })
    ;(window as unknown as { desktop: unknown }).desktop = { exportPdf: print }

    await exportPdf(
      {
        doc: { fileName: 'doc.docx' },
        setStatus: vi.fn(),
        sections: [],
        section: null,
      } as never,
      '/tmp/out.pdf',
    )

    expect(print).toHaveBeenCalledTimes(1)
    expect(atPrintTime).toEqual(['chap01', 'chap02'])
    expect(clip1.querySelector('#chap01')).not.toBeNull()
    expect(clip2.querySelector('#chap02')).not.toBeNull()
  })
})
