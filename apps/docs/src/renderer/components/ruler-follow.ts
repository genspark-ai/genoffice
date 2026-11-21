import { useEffect, useLayoutEffect, useRef } from 'react'
import type { Editor } from '@tiptap/core'

/**
 * Re-run `measure` (at most once per frame) whenever the page under a fixed
 * ruler may have moved: a render, an editor transaction, the editor resizing,
 * the window resizing, or the document view scrolling. The rulers sit outside
 * the scrolling view, as Word's do, and only line their scale up with the page.
 */
export function useRulerFollow(editor: Editor | null, measure: () => void): void {
  const measureRef = useRef(measure)
  measureRef.current = measure
  const rafRef = useRef(0)
  const schedule = useRef(() => {})
  schedule.current = () => {
    if (rafRef.current) return
    rafRef.current = requestAnimationFrame(() => {
      rafRef.current = 0
      measureRef.current()
    })
  }
  useLayoutEffect(() => schedule.current())
  useEffect(() => {
    const run = () => schedule.current()
    const dom = editor && !editor.isDestroyed ? (editor.view.dom as HTMLElement) : null
    const scroller = dom?.closest('.editor-scroll') ?? null
    editor?.on('transaction', run)
    scroller?.addEventListener('scroll', run, { passive: true })
    const ro = typeof ResizeObserver === 'undefined' || !dom ? null : new ResizeObserver(run)
    ro?.observe(dom!)
    window.addEventListener('resize', run)
    return () => {
      editor?.off('transaction', run)
      scroller?.removeEventListener('scroll', run)
      ro?.disconnect()
      window.removeEventListener('resize', run)
      cancelAnimationFrame(rafRef.current)
      rafRef.current = 0
    }
  }, [editor])
}
