import React, { useEffect, useRef, useState } from 'react'
import {
  PAGE_MARK,
  TOTAL_PAGES_MARK,
  type HfImage,
  type HfParagraph,
  type Run,
} from '@genoffice/docx-engine'
import { useI18n } from '../i18n/locale'
import { spellcheckEnabled } from '../spellcheck-pref'
import {
  hfCellGeometry,
  hfCellParaStyle,
  hfCellSegStyle,
  hfCellTabLines,
  hfDeclaredStrutPt,
  hfLeadIndentCss,
  hfRowStyle,
  hfSegLeftCss,
  hfTabSegments,
  hfBoxAnchorEl,
  hfTextBoxClass,
  hfTextBoxStyle,
  hfUsesLegacyHash,
  paraBorderCss,
  type HfStripGeom,
  hfParaLineHeightCss,
  hfStackedSpacingPx,
} from '../editor/hf-dom'
import { applyHfText, hfEditText, hfParasOf, PAGE_TOKEN } from '../editor/hf-text'
import { dkStyleProps } from '../editor/dark-page'
import { INLINE_RULE_CLASS, inlineRuleStyle } from '../editor/inline-rule'
import { textColorValue } from '../editor/text-color'
import { cssRunFontFamily, fontKerningCss, runLetterSpacingCss } from '../line-metrics'

export interface HfValue {
  text: string
  pageNumber?: boolean
  paras?: HfParagraph[]
}

function runStyle(run: Run): React.CSSProperties {
  const style: React.CSSProperties = {}
  if (run.bold) style.fontWeight = 600
  else if (run.bold === false) style.fontWeight = 'normal'
  if (run.italic) style.fontStyle = 'italic'
  else if (run.italic === false) style.fontStyle = 'normal'
  if (run.underline) style.textDecoration = 'underline'
  if (run.strike) style.textDecoration = `${style.textDecoration ?? ''} line-through`.trim()
  // authored color stays the declaration; the --dk-c twin feeds the dark page
  if (run.color) {
    style.color = textColorValue(run.color)
    if (run.color !== 'auto') Object.assign(style, dkStyleProps({ color: run.color }))
  }
  if (run.sizeHalfPoints) style.fontSize = `${run.sizeHalfPoints / 2}pt`
  const letterSpacing = runLetterSpacingCss(run)
  if (letterSpacing) style.letterSpacing = letterSpacing
  const kerning = fontKerningCss(run)
  if (kerning) style.fontKerning = kerning
  if (run.font || run.fontAscii) style.fontFamily = cssRunFontFamily(run.fontAscii, run.font)
  if (run.caps === 'all') style.textTransform = 'uppercase'
  else if (run.caps === 'small') style.fontVariantCaps = 'small-caps'
  else if (run.caps === 'none') {
    style.textTransform = 'none'
    style.fontVariantCaps = 'normal'
  }
  return style
}

/** document content colors (w:shd / w:pBdr) plus their dark-page twins; mirrors makeGapHfEl */
function paraStyle(para: HfParagraph): React.CSSProperties {
  const style: React.CSSProperties = {}
  const lh = hfParaLineHeightCss(para)
  if (lh) style.lineHeight = lh
  if (para.bidi) style.direction = 'rtl'
  if (para.align) {
    style.textAlign =
      para.align === 'left' || para.align === 'center' || para.align === 'right'
        ? para.align
        : 'justify'
  }
  // frame placement wins over the paragraph's own jc (mirrors makeGapHfEl)
  if (para.frameXAlign) style.textAlign = para.frameXAlign
  if (para.shadingFill) {
    style.backgroundColor = `#${para.shadingFill}`
    Object.assign(style, dkStyleProps({ background: `#${para.shadingFill}` }))
  }
  if (para.borders) {
    const line = (side: 't' | 'b' | 'l' | 'r') => paraBorderCss(para.borderLines?.[side])
    const borders: Partial<Record<'t' | 'b' | 'l' | 'r', string>> = {}
    if (para.borders.includes('t')) style.borderTop = borders.t = line('t')
    if (para.borders.includes('b')) style.borderBottom = borders.b = line('b')
    if (para.borders.includes('l')) style.borderLeft = borders.l = line('l')
    if (para.borders.includes('r')) style.borderRight = borders.r = line('r')
    Object.assign(style, dkStyleProps({ borders }))
    style.padding = '1px 4px'
  }
  return style
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
  style,
  boxGeom,
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
  /** geometry override: on differing-width sections the strip is pinned to its own section's box */
  style?: React.CSSProperties
  /** page geometry: floating textboxes render at their anchor position (absent: stacked) */
  boxGeom?: HfStripGeom
}) {
  const { t } = useI18n()
  const [editing, setEditing] = useState(false)
  const editRef = useRef<HTMLDivElement>(null)
  const cancelRef = useRef(false)
  const initialTextRef = useRef('')
  const paras = hfParasOf(value)

  // The editing surface is a standalone element: content is injected here and React
  // does not manage its children; after commit the whole element unmounts, so text
  // nodes produced while typing don't linger (keeps section/variant switches clean)
  useEffect(() => {
    if (!editing) return
    const el = editRef.current
    if (!el) return
    // table-row (cells) paragraphs stay out of the text editing flow
    el.innerText = hfEditText(value)
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
    onCommit(applyHfText(value, el.innerText))
  }

  const display = (text: string) => {
    const t = text
      .replaceAll(TOTAL_PAGES_MARK, String(pageTotal ?? 1))
      .replaceAll(PAGE_MARK, String(pageNo ?? 1))
    return hfUsesLegacyHash(value) ? t.replace('#', String(pageNo ?? 1)) : t
  }

  // run-declared strip strut (shrink-only), mirroring makeGapHfEl so the
  // preview/export strips lay out like the canvas gaps and the push-down probe
  const strutPt = hfDeclaredStrutPt(paras)
  const hasBoxes = boxGeom != null && paras.some((p) => p.box)
  return (
    <div
      className={`page-hf page-hf-${kind}${editing ? ' page-hf-editing' : ''}${hasBoxes ? ' page-hf-has-boxes' : ''}`}
      style={
        strutPt != null
          ? { fontSize: `min(${strutPt}pt, var(--hf-default-fs, 10.5pt))`, ...style }
          : style
      }
      data-tip={
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
      {images && images.some((im) => !im.floating) && (
        <div
          className="page-hf-images"
          contentEditable={false}
          style={
            images.find((im) => !im.floating)?.align === 'right'
              ? { justifyContent: 'flex-end' }
              : images.find((im) => !im.floating)?.align === 'center'
                ? { justifyContent: 'center' }
                : undefined
          }
        >
          {images
            .filter((img) => !img.floating)
            .map((img, i) => (
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
          spellCheck={spellcheckEnabled()}
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
        <HfContent kind={kind} paras={paras} display={display} boxGeom={boxGeom} />
      )}
    </div>
  )
}

function HfContent({
  kind,
  paras,
  display,
  boxGeom,
}: {
  kind: 'header' | 'footer'
  paras: HfParagraph[]
  display: (text: string) => string
  boxGeom?: HfStripGeom
}) {
  const spacing = hfStackedSpacingPx(paras)
  const margins = (i: number): React.CSSProperties => ({
    ...(spacing[i].top ? { marginTop: `${spacing[i].top}px` } : {}),
    ...(spacing[i].bottom ? { marginBottom: `${spacing[i].bottom}px` } : {}),
  })
  const renderPara = (para: HfParagraph, i: number) =>
    para.cells ? (
      // layout-table row: read-only flex columns (excluded from text editing)
      <div
        key={i}
        className="page-hf-para page-hf-row"
        style={{ ...hfRowStyle(para.row), ...margins(i) }}
      >
        {para.cells.map((cell, j) => {
          const geom = hfCellGeometry(cell)
          const spans = (runs: Run[]) =>
            runs.map((run, l) => (
              <span key={l} style={runStyle(run)}>
                {run.image?.rule && (
                  <span
                    className={INLINE_RULE_CLASS}
                    style={inlineRuleStyle({
                      ...run.image.rule,
                      sizeHalfPoints: run.sizeHalfPoints,
                    })}
                  />
                )}
                {run.image && !run.image.rule && (
                  <img
                    className="page-hf-cell-img"
                    src={run.image.dataUrl}
                    alt=""
                    draggable={false}
                    style={{
                      ...(run.image.widthPx ? { width: run.image.widthPx } : {}),
                      ...(run.image.heightPx ? { height: run.image.heightPx } : {}),
                    }}
                  />
                )}
                {display(run.text)}
              </span>
            ))
          return (
            <div
              key={j}
              className="page-hf-cell"
              style={{
                ...geom.style,
                // document content colors (w:shd / borders) plus their dark-page twins
                ...(cell.fill ? { backgroundColor: `#${cell.fill}` } : {}),
                ...dkStyleProps({
                  ...(cell.fill ? { background: `#${cell.fill}` } : {}),
                  borders: geom.borders,
                }),
              }}
            >
              {/* one block line per cell paragraph (mirrors makeGapHfEl) */}
              {(cell.paras.length > 0 ? cell.paras : [[]]).map((runs, k) => {
                const props = cell.paraProps?.[k]
                const tabLines = hfCellTabLines(runs, props, geom, para.row, display)
                if (!tabLines) {
                  return (
                    <div key={k} className="page-hf-cell-para" style={hfCellParaStyle(props)}>
                      {runs.length === 0 ? ' ' : null}
                      {spans(runs)}
                    </div>
                  )
                }
                return tabLines.map((line, m) => (
                  <div
                    key={`${k}-${m}`}
                    className="page-hf-cell-para page-hf-tabbed"
                    style={{
                      ...hfCellParaStyle(props, {
                        first: m === 0,
                        last: m === tabLines.length - 1,
                      }),
                      textAlign: 'left',
                      ...(line.minHeightPt ? { minHeight: `${line.minHeightPt}pt` } : {}),
                    }}
                  >
                    {spans(line.lead)}
                    {line.segments.map((seg, n) => (
                      <span
                        key={`t${n}`}
                        className={`page-hf-tabseg page-hf-tabseg-${seg.anchor}`}
                        style={hfCellSegStyle(seg)}
                      >
                        {spans(seg.runs)}
                      </span>
                    ))}
                  </div>
                ))
              })}
            </div>
          )
        })}
      </div>
    ) : (
      (() => {
        const tabbed = hfTabSegments(para, display)
        if (!tabbed) {
          return (
            <div
              key={i}
              className={`page-hf-para${para.frameXAlign ? ' page-hf-frame' : ''}`}
              style={{ ...paraStyle(para), ...margins(i) }}
            >
              {para.runs.length === 0 ? ' ' : null}
              {para.runs.map((run, j) => (
                <span key={j} style={runStyle(run)}>
                  {display(run.text)}
                </span>
              ))}
            </div>
          )
        }
        const leadIndent = hfLeadIndentCss(tabbed)
        return (
          <div
            key={i}
            className={`page-hf-para page-hf-tabbed${para.frameXAlign ? ' page-hf-frame' : ''}`}
            style={{
              ...paraStyle(para),
              ...margins(i),
              ...(tabbed.minHeightPt ? { minHeight: `${tabbed.minHeightPt}pt` } : {}),
              // tab layout happens in left-aligned space; w:jc becomes an explicit shift
              textAlign: 'left',
              ...(leadIndent ? { textIndent: leadIndent } : {}),
            }}
          >
            {tabbed.lead.map((run, j) => (
              <span key={j} style={runStyle(run)}>
                {display(run.text)}
              </span>
            ))}
            {tabbed.segments.map((seg, k) => (
              <span
                key={`t${k}`}
                className={`page-hf-tabseg page-hf-tabseg-${seg.anchor}`}
                style={{ left: hfSegLeftCss(seg, tabbed) }}
              >
                {seg.runs.map((run, j) => (
                  <span key={j} style={runStyle(run)}>
                    {display(run.text)}
                  </span>
                ))}
              </span>
            ))}
          </div>
        )
      })()
    )
  // consecutive paragraphs of one floating textbox render inside a positioned
  // box (Word draws them at the anchor); everything else stacks as strip lines.
  // A box sharing its paragraph with text hangs off that paragraph instead.
  const indices = paras.map((_, i) => i)
  const hostedBy = new Map<number, React.ReactNode[]>()
  const out: React.ReactNode[] = []
  for (let i = 0; i < paras.length;) {
    const box = paras[i].box
    const css = box && boxGeom ? hfTextBoxStyle(box, kind, boxGeom) : null
    if (!css) {
      out.push(indices[i])
      i += 1
      continue
    }
    const items: React.ReactNode[] = []
    const start = i
    for (; i < paras.length && paras[i].box?.id === box?.id; i += 1) {
      items.push(renderPara(paras[i], i))
    }
    const node = (
      <div key={`box${start}`} className={hfTextBoxClass(box!)} style={css}>
        {items}
      </div>
    )
    const anchor = hfBoxAnchorEl(box!, indices)
    if (anchor != null) hostedBy.set(anchor, [...(hostedBy.get(anchor) ?? []), node])
    else out.push(node)
  }
  return (
    <>
      {out.map((item) => {
        if (typeof item !== 'number') return item
        const el = renderPara(paras[item], item) as React.ReactElement<{
          className?: string
          children?: React.ReactNode
        }>
        const hosted = hostedBy.get(item)
        if (!hosted) return el
        return React.cloneElement(
          el,
          { className: `${el.props.className ?? ''} page-hf-anchor` },
          el.props.children,
          ...hosted,
        )
      })}
    </>
  )
}
