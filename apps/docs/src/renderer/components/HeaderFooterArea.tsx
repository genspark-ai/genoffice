import { useEffect, useRef, useState } from 'react'
import { TOTAL_PAGES_MARK, type HfImage, type HfParagraph, type Run } from '@genoffice/docx-engine'
import { useI18n } from '../i18n/locale'
import { cssFontFamily } from '../line-metrics'

export interface HfValue {
  text: string
  pageNumber?: boolean
  paras?: HfParagraph[]
}

function runStyle(run: Run): React.CSSProperties {
  const style: React.CSSProperties = {}
  if (run.bold) style.fontWeight = 600
  if (run.italic) style.fontStyle = 'italic'
  if (run.underline) style.textDecoration = 'underline'
  if (run.strike) style.textDecoration = `${style.textDecoration ?? ''} line-through`.trim()
  if (run.color) style.color = `#${run.color}`
  if (run.sizeHalfPoints) style.fontSize = `${run.sizeHalfPoints / 2}pt`
  if (run.font) style.fontFamily = cssFontFamily(run.font)
  return style
}

/** effective paragraphs: rich paras when present, else the legacy single line */
function parasOf(value: HfValue): HfParagraph[] {
  if (value.paras?.length) return value.paras
  const runs: Run[] = value.text ? [{ text: value.text }] : []
  if (value.pageNumber && !value.text.includes('#')) {
    runs.push({ text: runs.length > 0 ? ' #' : '#' })
  }
  return [{ align: 'center', runs }]
}

/**
 * Header / footer zone on the page: renders the rich paragraphs,
 * double-click enters in-place editing (plain text per paragraph; each line
 * keeps its paragraph format and first-run styling), blur commits. The '#'
 * marker stands for the automatic page number.
 */
export function HeaderFooterArea({
  kind,
  value,
  images,
  readOnly,
  onCommit,
  pageNo,
  pageTotal,
}: {
  kind: 'header' | 'footer'
  value: HfValue
  /** logo and other images in the part, display-only (text edits do not affect their saved bytes) */
  images?: HfImage[]
  readOnly?: boolean
  onCommit: (next: HfValue) => void
  /** Page number shown for '#' (may be a section-formatted string); the continuous-flow canvas has no real page number, defaults to 1 */
  pageNo?: number | string
  /** Total page count shown for TOTAL_PAGES_MARK (NUMPAGES field), defaults to 1 */
  pageTotal?: number
}) {
  const { t } = useI18n()
  const [editing, setEditing] = useState(false)
  const editRef = useRef<HTMLDivElement>(null)
  const cancelRef = useRef(false)
  const initialTextRef = useRef('')
  const paras = parasOf(value)

  // The editing surface is a standalone element: content is injected here and React
  // does not manage its children; after commit the whole element unmounts, so text
  // nodes produced while typing don't linger (keeps section/variant switches clean)
  useEffect(() => {
    if (!editing) return
    const el = editRef.current
    if (!el) return
    el.innerText = paras.map((p) => p.runs.map((r) => r.text).join('')).join('\n')
    cancelRef.current = false
    initialTextRef.current = el.innerText
    el.focus()
    const sel = window.getSelection()
    if (sel) {
      sel.selectAllChildren(el)
      sel.collapseToEnd()
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [editing])

  const commit = () => {
    const el = editRef.current
    setEditing(false)
    if (!el) return
    if (cancelRef.current) {
      cancelRef.current = false
      return
    }
    if (el.innerText === initialTextRef.current) return
    const lines = el.innerText.replace(/\n+$/, '').split('\n')
    const nextParas: HfParagraph[] = lines.map((line, i) => {
      const template = paras[Math.min(i, paras.length - 1)]
      const style = template.runs[0] ?? {}
      return { ...template, runs: line === '' ? [] : [{ ...style, text: line }] }
    })
    const text = nextParas.map((p) => p.runs.map((r) => r.text).join('')).join('')
    onCommit({ ...value, text, paras: nextParas })
  }

  const display = (text: string) => {
    const t = text.replaceAll(TOTAL_PAGES_MARK, String(pageTotal ?? 1))
    return value.pageNumber ? t.replace('#', String(pageNo ?? 1)) : t
  }

  return (
    <div
      className={`page-hf page-hf-${kind}${editing ? ' page-hf-editing' : ''}`}
      title={
        readOnly
          ? undefined
          : t(kind === 'header' ? 'appDblclickEditHeader' : 'appDblclickEditFooter') +
            (value.pageNumber ? t('appHfPageNumHint') : '')
      }
      onDoubleClick={() => {
        if (!readOnly && !editing) setEditing(true)
      }}
    >
      {images && images.length > 0 && (
        <div className="page-hf-images" contentEditable={false}>
          {images.map((img, i) => (
            <img
              key={i}
              src={img.dataUrl}
              alt=""
              draggable={false}
              style={{
                ...(img.widthPx ? { width: img.widthPx } : {}),
                ...(img.heightPx ? { height: img.heightPx } : {}),
              }}
            />
          ))}
        </div>
      )}
      {editing ? (
        <div
          ref={editRef}
          className="page-hf-edit-surface"
          contentEditable
          suppressContentEditableWarning
          onBlur={commit}
          onKeyDown={(e) => {
            if (e.key === 'Escape') {
              e.preventDefault()
              e.stopPropagation()
              cancelRef.current = true
              ;(e.target as HTMLElement).blur()
            }
          }}
        />
      ) : (
        <HfContent paras={paras} display={display} />
      )}
    </div>
  )
}

function HfContent({
  paras,
  display,
}: {
  paras: HfParagraph[]
  display: (text: string) => string
}) {
  return (
    <>
      {paras.map((para, i) => (
        <div
          key={i}
          className="page-hf-para"
          style={{
            ...(para.bidi ? { direction: 'rtl' as const } : {}),
            ...(para.align
              ? {
                  textAlign:
                    para.align === 'left' || para.align === 'center' || para.align === 'right'
                      ? para.align
                      : ('justify' as const),
                }
              : {}),
          }}
        >
          {para.runs.length === 0 ? ' ' : null}
          {para.runs.map((run, j) => (
            <span key={j} style={runStyle(run)}>
              {display(run.text)}
            </span>
          ))}
        </div>
      ))}
    </>
  )
}
