import { useEffect, useRef, useState } from 'react'
import {
  PAGE_MARK,
  TOTAL_PAGES_MARK,
  type HfImage,
  type HfParagraph,
  type Run,
} from '@genoffice/docx-engine'
import { useI18n } from '../i18n/locale'
import { hfUsesLegacyHash } from '../editor/hf-dom'
import { cssDualFontFamily, cssFontFamily } from '../line-metrics'

export interface HfValue {
  text: string
  pageNumber?: boolean
  paras?: HfParagraph[]
}

/** visible edit-surface stand-ins for the invisible private-use field sentinels */
const PAGE_TOKEN = '{PAGE}'
const TOTAL_TOKEN = '{NUMPAGES}'

function runStyle(run: Run): React.CSSProperties {
  const style: React.CSSProperties = {}
  if (run.bold) style.fontWeight = 600
  if (run.italic) style.fontStyle = 'italic'
  if (run.underline) style.textDecoration = 'underline'
  if (run.strike) style.textDecoration = `${style.textDecoration ?? ''} line-through`.trim()
  if (run.color) style.color = `#${run.color}`
  if (run.sizeHalfPoints) style.fontSize = `${run.sizeHalfPoints / 2}pt`
  if (run.font && run.fontAscii) style.fontFamily = cssDualFontFamily(run.fontAscii, run.font)
  else if (run.font || run.fontAscii) style.fontFamily = cssFontFamily((run.font ?? run.fontAscii)!)
  return style
}

/** effective paragraphs: rich paras when present, else the legacy single line */
function parasOf(value: HfValue): HfParagraph[] {
  if (value.paras?.length) return value.paras
  const runs: Run[] = value.text ? [{ text: value.text }] : []
  if (value.pageNumber && !value.text.includes('#') && !value.text.includes(PAGE_MARK)) {
    runs.push({ text: runs.length > 0 ? ` ${PAGE_MARK}` : PAGE_MARK })
  }
  return [{ align: 'center', runs }]
}

/**
 * Header / footer zone on the page: renders the rich paragraphs,
 * double-click enters in-place editing (plain text per paragraph; each line
 * keeps its paragraph format and first-run styling), blur commits. PAGE /
 * NUMPAGES sentinels edit as visible {PAGE} / {NUMPAGES} tokens.
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
    // table-row (cells) paragraphs stay out of the text editing flow
    el.innerText = paras
      .filter((p) => !p.cells)
      .map((p) => p.runs.map((r) => r.text).join(''))
      .join('\n')
      .replaceAll(PAGE_MARK, PAGE_TOKEN)
      .replaceAll(TOTAL_PAGES_MARK, TOTAL_TOKEN)
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
    const lines = el.innerText
      .replace(/\n+$/, '')
      .replaceAll(PAGE_TOKEN, PAGE_MARK)
      .replaceAll(TOTAL_TOKEN, TOTAL_PAGES_MARK)
      .split('\n')
    const textParas = paras.filter((p) => !p.cells)
    const templates: HfParagraph[] =
      textParas.length > 0 ? textParas : [{ align: 'center', runs: [] }]
    const edited: HfParagraph[] = lines.map((line, i) => {
      const template = templates[Math.min(i, templates.length - 1)]
      const style = template.runs[0] ?? {}
      return { ...template, runs: line === '' ? [] : [{ ...style, text: line }] }
    })
    // splice cells rows back at their original positions among the text paragraphs
    const nextParas: HfParagraph[] = []
    let ei = 0
    for (const p of paras) {
      if (p.cells) nextParas.push(p)
      else if (ei < edited.length) nextParas.push(edited[ei++])
    }
    nextParas.push(...edited.slice(ei))
    const text = edited.map((p) => p.runs.map((r) => r.text).join('')).join('')
    onCommit({ ...value, text, paras: nextParas })
  }

  const display = (text: string) => {
    const t = text
      .replaceAll(TOTAL_PAGES_MARK, String(pageTotal ?? 1))
      .replaceAll(PAGE_MARK, String(pageNo ?? 1))
    return hfUsesLegacyHash(value) ? t.replace('#', String(pageNo ?? 1)) : t
  }

  return (
    <div
      className={`page-hf page-hf-${kind}${editing ? ' page-hf-editing' : ''}`}
      title={
        readOnly
          ? undefined
          : t(kind === 'header' ? 'appDblclickEditHeader' : 'appDblclickEditFooter') +
            (value.pageNumber
              ? hfUsesLegacyHash(value)
                ? t('appHfPageNumHint')
                : t('appHfPageNumHint').replace('#', PAGE_TOKEN)
              : '')
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
      {paras.map((para, i) =>
        para.cells ? (
          // layout-table row: read-only flex columns (excluded from text editing)
          <div key={i} className="page-hf-para page-hf-row">
            {para.cells.map((cell, j) => (
              <div
                key={j}
                className="page-hf-cell"
                style={{
                  ...(cell.widthPct ? { width: `${cell.widthPct}%` } : {}),
                  ...(cell.align
                    ? {
                        textAlign:
                          cell.align === 'left' || cell.align === 'center' || cell.align === 'right'
                            ? cell.align
                            : ('justify' as const),
                      }
                    : {}),
                }}
              >
                {cell.runs.map((run, k) => (
                  <span key={k} style={runStyle(run)}>
                    {display(run.text)}
                  </span>
                ))}
              </div>
            ))}
          </div>
        ) : (
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
        ),
      )}
    </>
  )
}
