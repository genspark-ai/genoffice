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
  hfAnchoredImgStyle,
  hfCellParaStyle,
  hfImageHangsOnPara,
  hfCellSegStyle,
  hfCellTabLines,
  hfDeclaredStrutPt,
  hfLeadIndentCss,
  hfRowStyle,
  hfSegLeftCss,
  hfTabLeadNeedsStrut,
  hfTabLines,
  hfTabOverflowPx,
  hfBoxAnchorEl,
  hfTextBoxClass,
  hfTextBoxStyle,
  hfUsesLegacyHash,
  paraBorderCss,
  paraBorderPadding,
  type HfStripGeom,
  type HfTabLayout,
  hfParaLineHeightCss,
  hfParaIndentStyle,
  hfStackedSpacingPx,
} from '../editor/hf-dom'
import { hfParasOf, PAGE_TOKEN } from '../editor/hf-text'
import { mountHfEditor, type HfEditorHandle } from '../editor/hf-editor'
import { dkStyleProps } from '../editor/dark-page'
import { INLINE_RULE_CLASS, inlineRuleStyle } from '../editor/inline-rule'
import { textColorValue } from '../editor/text-color'
import { textOutlineCssValue } from '../editor/text-outline'
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
  if (run.textOutline) style.WebkitTextStroke = textOutlineCssValue(run.textOutline)
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
  Object.assign(style, hfParaIndentStyle(para))
  if (para.align) {
    style.textAlign =
      para.align === 'left' || para.align === 'center' || para.align === 'right'
        ? para.align
        : 'justify'
  }
  // frame placement wins over the paragraph's own jc (mirrors makeGapHfEl)
  if (para.frameXAlign) style.textAlign = para.frameXAlign
  const shdBg = para.shadingDisplay ?? para.shadingFill
  if (shdBg) {
    style.backgroundColor = `#${shdBg}`
    Object.assign(style, dkStyleProps({ background: `#${shdBg}` }))
  } else if (para.shadingClear) style.backgroundColor = 'transparent'
  if (para.borders) {
    const line = (side: 't' | 'b' | 'l' | 'r') => paraBorderCss(para.borderLines?.[side])
    const borders: Partial<Record<'t' | 'b' | 'l' | 'r', string>> = {}
    if (para.borders.includes('t')) style.borderTop = borders.t = line('t')
    if (para.borders.includes('b')) style.borderBottom = borders.b = line('b')
    if (para.borders.includes('l')) style.borderLeft = borders.l = line('l')
    if (para.borders.includes('r')) style.borderRight = borders.r = line('r')
    Object.assign(style, dkStyleProps({ borders }))
    Object.assign(style, paraBorderPadding(para.borders, para.borderLines))
  }
  return style
}

/**
 * Header / footer zone on the page: renders the rich paragraphs; double-click
 * (or an editRequest) mounts a nested rich-text editor over the strip (Word's
 * header editing mode), committing when focus leaves it.
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
  linked,
  sectionLabel,
  editRequest,
  onEditingChange,
}: {
  kind: 'header' | 'footer'
  value: HfValue
  /** logo and other images in the part, display-only (text edits do not affect their saved bytes) */
  images?: HfImage[]
  readOnly?: boolean
  onCommit: (next: HfValue) => void
  /** Word's "Same as Previous" tag: this section inherits the strip from the previous one */
  linked?: boolean | null
  /** "Header -Section 2-" style label shown while editing */
  sectionLabel?: string | null
  /** bump to enter editing from outside (ribbon Edit Header / Go to Footer) */
  editRequest?: number | null
  /** editing mode changes, with the live editor handle while open */
  onEditingChange?: (editing: boolean, handle: HfEditorHandle | null) => void
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
  const rootRef = useRef<HTMLDivElement>(null)
  const hoverRef = useRef(false)
  // the wrap the flag was last set on: rootRef is already null in the unmount cleanup
  const wrapRef = useRef<Element | null>(null)
  // the idle variant chips sit before the page in .page-wrap and reveal while
  // the header is hovered or edited; a data flag on the wrap replaces the
  // .page-wrap:has(...) rule that restyled the whole document on every mutation
  const syncWrapFlag = (on: boolean) => {
    if (kind !== 'header') return
    const wrap = rootRef.current?.closest('.page-wrap') ?? wrapRef.current
    if (!wrap) return
    wrapRef.current = wrap
    if (on) wrap.setAttribute('data-hf-active', '')
    else wrap.removeAttribute('data-hf-active')
  }
  useEffect(() => {
    syncWrapFlag(editing || hoverRef.current)
    return () => syncWrapFlag(false)
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [editing, kind])
  const paras = hfParasOf(value, images)

  useEffect(() => {
    if (editRequest != null && !readOnly) setEditing(true)
  }, [editRequest, readOnly])

  // The editor host is a standalone element React never populates: the nested
  // editor owns its children and the whole element unmounts on exit, so no
  // stale DOM survives section/variant switches
  const latest = useRef({ value, onCommit, onEditingChange, pageNo, pageTotal })
  latest.current = { value, onCommit, onEditingChange, pageNo, pageTotal }
  useEffect(() => {
    if (!editing) return
    const el = editRef.current
    if (!el) return
    const { value: v, pageNo: pn, pageTotal: pt } = latest.current
    const handle = mountHfEditor(el, {
      value: v,
      pageNo: String(pn ?? 1),
      pageTotal: String(pt ?? 1),
      spellcheck: spellcheckEnabled(),
      onCommit: (next) => latest.current.onCommit(next),
      onExit: () => {
        latest.current.onEditingChange?.(false, null)
        setEditing(false)
      },
    })
    latest.current.onEditingChange?.(true, handle)
    return () => handle.exit()
  }, [editing])

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
  const tabOver = Math.max(
    0,
    ...paras.map((p) => (p.cells ? 0 : hfTabOverflowPx(hfTabLines(p, display) ?? [], boxGeom))),
  )
  return (
    <div
      ref={rootRef}
      className={`page-hf page-hf-${kind}${editing ? ' page-hf-editing' : ''}${hasBoxes ? ' page-hf-has-boxes' : ''}`}
      onMouseEnter={() => {
        hoverRef.current = true
        syncWrapFlag(true)
      }}
      onMouseLeave={() => {
        hoverRef.current = false
        syncWrapFlag(editing)
      }}
      style={{
        ...(strutPt != null ? { fontSize: `min(${strutPt}pt, var(--hf-default-fs, 10.5pt))` } : {}),
        ...(tabOver > 0 ? { ['--hf-tab-over' as string]: `${tabOver.toFixed(1)}px` } : {}),
        ...style,
      }}
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
        <div ref={editRef} className="page-hf-edit-surface" />
      ) : (
        <HfContent kind={kind} paras={paras} images={images} display={display} boxGeom={boxGeom} />
      )}
      {editing && (sectionLabel || linked) && (
        <div className="page-hf-tags" contentEditable={false}>
          {sectionLabel && <span className="page-hf-tag">{sectionLabel}</span>}
          {linked && <span className="page-hf-tag">{t('appHfSameAsPrevious')}</span>}
        </div>
      )}
    </div>
  )
}

function HfContent({
  kind,
  paras,
  images,
  display,
  boxGeom,
}: {
  kind: 'header' | 'footer'
  paras: HfParagraph[]
  images?: HfImage[]
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
                    <div
                      key={k}
                      className="page-hf-cell-para"
                      style={hfCellParaStyle(props, undefined, runs)}
                    >
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
                      ...hfCellParaStyle(
                        props,
                        { first: m === 0, last: m === tabLines.length - 1 },
                        runs,
                      ),
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
        const tabLines = hfTabLines(para, display)
        if (!tabLines) {
          return (
            <div
              key={i}
              className={`page-hf-para${para.frameXAlign ? ' page-hf-frame' : ''}`}
              style={{
                ...paraStyle(para),
                ...margins(i),
                ...(para.runs.length === 0 && para.emptyRunSizeHalfPoints
                  ? { fontSize: `${para.emptyRunSizeHalfPoints / 2}pt` }
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
          )
        }
        // tab layout happens in left-aligned space; w:jc becomes an explicit shift
        const lineStyle = (tabbed: HfTabLayout): React.CSSProperties => {
          const leadIndent = hfLeadIndentCss(tabbed)
          return {
            ...(tabbed.minHeightPt ? { minHeight: `${tabbed.minHeightPt}pt` } : {}),
            textAlign: 'left',
            ...(leadIndent ? { textIndent: leadIndent } : {}),
          }
        }
        const lineContent = (tabbed: HfTabLayout) => (
          <>
            {hfTabLeadNeedsStrut(tabbed) ? '\u200b' : null}
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
          </>
        )
        const frame = para.frameXAlign ? ' page-hf-frame' : ''
        if (tabLines.length === 1) {
          return (
            <div
              key={i}
              className={`page-hf-para page-hf-tabbed${frame}`}
              style={{ ...paraStyle(para), ...margins(i), ...lineStyle(tabLines[0]) }}
            >
              {lineContent(tabLines[0])}
            </div>
          )
        }
        // a w:br paragraph stacks one positioned line per break inside the
        // paragraph block (which keeps the spacing and borders)
        return (
          <div
            key={i}
            className={`page-hf-para${frame}`}
            style={{ ...paraStyle(para), ...margins(i) }}
          >
            {tabLines.map((tabbed, m) => (
              <div key={m} className="page-hf-tabbed" style={lineStyle(tabbed)}>
                {lineContent(tabbed)}
              </div>
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
  for (const [n, img] of (images ?? []).filter(hfImageHangsOnPara).entries()) {
    if (img.anchorPara! >= paras.length) continue
    const node = (
      <img
        key={`img${n}`}
        src={img.dataUrl}
        alt=""
        draggable={false}
        style={hfAnchoredImgStyle(img, boxGeom)}
      />
    )
    hostedBy.set(img.anchorPara!, [...(hostedBy.get(img.anchorPara!) ?? []), node])
  }
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
