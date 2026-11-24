import { useEffect, useRef } from 'react'
import type { Editor } from '@tiptap/core'
import { useI18n } from '../i18n/locale'
import { defaultParagraphStyleId, styleLabel, type StyleMap } from '../style-gallery'

/**
 * The style area of Draft and Outline views: a column left of the text that
 * names the style of each paragraph on its first line. Window chrome only:
 * it is drawn outside the page and no print or export path sees it.
 */
export function StyleArea({
  editor,
  widthPx,
  styles,
}: {
  editor: Editor
  widthPx: number
  styles: StyleMap | undefined
}) {
  const { t } = useI18n()
  const layerRef = useRef<HTMLDivElement>(null)
  const labelRef = useRef((styleId: string | null): string => styleId ?? '')
  labelRef.current = (styleId) => {
    const info = styles?.get(styleId ?? defaultParagraphStyleId(styles) ?? 'Normal')
    return info ? styleLabel(info, t) : (styleId ?? t('ribbonStyleNormal'))
  }

  useEffect(() => {
    const layer = layerRef.current
    const pane = layer?.closest<HTMLElement>('.editor-scroll')
    if (!layer || !pane) return
    const band = layer.firstElementChild as HTMLElement
    let frame = 0
    const sync = () => {
      frame = 0
      if (editor.isDestroyed) return
      const { view } = editor
      const layerTop = layer.getBoundingClientRect().top
      band.style.top = `${pane.getBoundingClientRect().top - pane.scrollTop - layerTop}px`
      band.style.height = `${pane.scrollHeight}px`
      let used = 1
      view.state.doc.forEach((node, offset) => {
        if (!node.isTextblock) return
        const el = view.nodeDOM(offset)
        // Outline view hides body text: a hidden paragraph has no row
        if (!(el instanceof HTMLElement) || el.offsetParent === null) return
        let row = layer.children[used] as HTMLElement | undefined
        if (!row) {
          row = document.createElement('div')
          row.className = 'style-area-name'
          layer.appendChild(row)
        }
        const name = labelRef.current(node.attrs.styleId as string | null)
        if (row.textContent !== name) row.textContent = name
        row.style.top = `${(el.getBoundingClientRect().top - layerTop).toFixed(1)}px`
        used++
      })
      while (layer.children.length > used) layer.lastElementChild?.remove()
    }
    const ask = () => {
      if (!frame) frame = requestAnimationFrame(sync)
    }
    ask()
    editor.on('transaction', ask)
    const resized = new ResizeObserver(ask)
    resized.observe(editor.view.dom)
    return () => {
      cancelAnimationFrame(frame)
      editor.off('transaction', ask)
      resized.disconnect()
    }
  }, [editor, widthPx, styles])

  return (
    <div className="style-area" ref={layerRef} aria-hidden="true">
      <div className="style-area-band" />
    </div>
  )
}
