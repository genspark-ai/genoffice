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
  hfCellParaBox,
  hfImageHangsOnPara,
  hfCellSegStyle,
  hfCellTabLines,
  hfLeadIndentCss,
  hfIndentOverhangPx,
  hfRowGeomOf,
  hfRowRightOverhangPx,
  hfRowStyle,
  hfRunFace,
  hfSegHasImage,
  hfSegLeftCss,
  hfTabLeadNeedsStrut,
  hfTabSegImageHeightPx,
  hfStripColumnPx,
  hfTabLines,
  hfTabOverflowPx,
  hfBoxAnchorEl,
  hfTextBoxClass,
  hfTextBoxStyle,
  hfUsesLegacyHash,
  hfBorderMergeFlags,
  hfParaBorderStyle,
  type ParaBorderMergeFlags,
  type HfStripGeom,
  type HfTabLayout,
  hfParaIndentStyle,
  hfObstacleStyles,
  hfSideWrapObstacles,
  hfStackedSpacingPx,
} from '../editor/hf-dom'
import { hfParasOf, PAGE_TOKEN } from '../editor/hf-text'
import { fromHfParagraph, toStyleRecord } from '../editor/para-layout-spec'
import { hfEditablePara, mountHfEditor, type HfEditorHandle } from '../editor/hf-editor'
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

function runRevClass(run: Run): string | undefined {
  const cls = [run.ins && 'doc-ins', run.del && 'doc-del'].filter(Boolean).join(' ')
  return cls || undefined
}

function runStyle(run: Run): React.CSSProperties {
  const style: React.CSSProperties = {}
  const face = hfRunFace(run)
  if (face.bold) style.fontWeight = 600
  else if (face.bold === false) style.fontWeight = 'normal'
  if (face.italic) style.fontStyle = 'italic'
  else if (face.italic === false) style.fontStyle = 'normal'
  if (run.underline) style.textDecoration = 'underline'
  if (run.strike) style.textDecoration = `${style.textDecoration ?? ''} line-through`.trim()
  // authored color stays the declaration; the --dk-c twin feeds the dark page
  if (run.color) {
    style.color = textColorValue(run.color)
    if (run.color !== 'auto') Object.assign(style, dkStyleProps({ color: run.color }))
  }
  if (face.sizeHalfPoints) style.fontSize = `${face.sizeHalfPoints / 2}pt`
  const letterSpacing = runLetterSpacingCss(run)
  if (letterSpacing) style.letterSpacing = letterSpacing
  const kerning = fontKerningCss(run)
  if (kerning) style.fontKerning = kerning
  if (face.font || face.fontAscii) style.fontFamily = cssRunFontFamily(face.fontAscii, face.font)
  if (run.textOutline) style.WebkitTextStroke = textOutlineCssValue(run.textOutline)
  if (run.caps === 'all') style.textTransform = 'uppercase'
  else if (run.caps === 'small') style.fontVariantCaps = 'small-caps'
  else if (run.caps === 'none') {
    style.textTransform = 'none'
    style.fontVariantCaps = 'normal'
  }
  return style
}

/** line box (para-layout-spec) and document content colors (w:shd / w:pBdr) plus
 *  their dark-page twins; mirrors makeGapHfEl */
function paraBox(
  para: HfParagraph,
  merge?: ParaBorderMergeFlags,
): { style: React.CSSProperties; classes: string[] } {
  const box = fromHfParagraph(para)
  const style: React.CSSProperties = toStyleRecord(box)
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
  const bs = hfParaBorderStyle(para, merge)
  Object.assign(style, bs.style, dkStyleProps({ borders: bs.borders }))
  return { style, classes: box.classes }
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

  const hasBoxes = boxGeom != null && paras.some((p) => p.box)
  const firstEditable = hfEditorSlot(paras)
  const rowGeom = hfRowGeomOf(kind, boxGeom)
  const indOver = hfIndentOverhangPx(paras, rowGeom)
  const tabOver = Math.max(
    hfRowRightOverhangPx(paras, rowGeom),
    ...paras.map((p) =>
      p.cells
        ? 0
        : hfTabOverflowPx(hfTabLines(p, display, hfStripColumnPx(boxGeom)) ?? [], boxGeom),
    ),
  )
  return (
    <div
      ref={rootRef}
      className={`page-hf page-hf-${kind}${editing ? ' page-hf-editing' : ''}${hasBoxes ? ' page-hf-has-boxes' : ''}`}
      style={{
        ...(tabOver > 0 ? { ['--hf-tab-over' as string]: `${tabOver.toFixed(1)}px` } : {}),
        ...(indOver > 0 ? { ['--hf-ind-over' as string]: `${indOver.toFixed(1)}px` } : {}),
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
        // rows and floating boxes stay display-only around the editor (its doc
        // holds the text paragraphs only), so a table-held logo does not vanish
        <>
          <HfContent
            kind={kind}
            paras={paras}
            images={images}
            display={display}
            boxGeom={boxGeom}
            only={(_, i) => i < firstEditable}
          />
          <div ref={editRef} className="page-hf-edit-surface" />
          <HfContent
            kind={kind}
            paras={paras}
            images={images}
            display={display}
            boxGeom={boxGeom}
            only={(_, i) => i >= firstEditable}
          />
        </>
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

/** strip index the editor stands in for; with no editable paragraph it goes
 *  last, where hfValueFromDoc appends typed text */
export function hfEditorSlot(paras: HfParagraph[]): number {
  const i = paras.findIndex(hfEditablePara)
  return i === -1 ? paras.length : i
}

export function HfContent({
  kind,
  paras,
  images,
  display,
  boxGeom,
  only,
}: {
  kind: 'header' | 'footer'
  paras: HfParagraph[]
  images?: HfImage[]
  display: (text: string) => string
  boxGeom?: HfStripGeom
  /** editing: this half of the strip owns these indices (spacing still
   *  measures the whole strip); editable paragraphs live in the editor and
   *  only lend their hosted boxes to the half they fall in */
  only?: (para: HfParagraph, index: number) => boolean
}) {
  const spacing = hfStackedSpacingPx(paras)
  const margins = (i: number): React.CSSProperties => ({
    ...(spacing[i].top ? { marginTop: `${spacing[i].top}px` } : {}),
    ...(spacing[i].bottom ? { marginBottom: `${spacing[i].bottom}px` } : {}),
  })
  const runSpans = (runs: Run[], imgClass: string) =>
    runs.map((run, l) => (
      <span key={l} className={runRevClass(run)} style={runStyle(run)}>
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
            className={imgClass}
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
  const mergeFlags = hfBorderMergeFlags(paras)
  const renderPara = (para: HfParagraph, i: number) =>
    para.cells ? (
      // layout-table row: read-only flex columns (excluded from text editing)
      <div
        key={i}
        className="page-hf-para page-hf-row"
        style={{ ...margins(i), ...hfRowStyle(para.row, hfRowGeomOf(kind, boxGeom), i === 0) }}
      >
        {para.cells.map((cell, j) => {
          const geom = hfCellGeometry(cell)
          const spans = (runs: Run[]) => runSpans(runs, 'page-hf-cell-img')
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
                  backgroundImage: geom.diagonals,
                  borders: geom.borders,
                }),
              }}
            >
              {/* one block line per cell paragraph (mirrors makeGapHfEl) */}
              {(cell.paras.length > 0 ? cell.paras : [[]]).map((runs, k) => {
                const props = cell.paraProps?.[k]
                const tabLines = hfCellTabLines(runs, props, geom, para.row, display)
                if (!tabLines) {
                  const box = hfCellParaBox(props, undefined, runs, para.row)
                  return (
                    <div
                      key={k}
                      className={['page-hf-cell-para', ...box.classes].join(' ')}
                      style={box.style}
                    >
                      {runs.length === 0 ? ' ' : null}
                      {spans(runs)}
                    </div>
                  )
                }
                return tabLines.map((line, m) => {
                  const box = hfCellParaBox(
                    props,
                    { first: m === 0, last: m === tabLines.length - 1 },
                    runs,
                    para.row,
                  )
                  return (
                    <div
                      key={`${k}-${m}`}
                      className={['page-hf-cell-para page-hf-tabbed', ...box.classes].join(' ')}
                      style={{
                        ...box.style,
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
                  )
                })
              })}
            </div>
          )
        })}
      </div>
    ) : (
      (() => {
        const tabLines = hfTabLines(para, display, hfStripColumnPx(boxGeom))
        if (!tabLines) {
          const box = paraBox(para, mergeFlags[i])
          return (
            <div
              key={i}
              className={['page-hf-para', para.frameXAlign && 'page-hf-frame', ...box.classes]
                .filter(Boolean)
                .join(' ')}
              style={{ ...box.style, ...margins(i) }}
            >
              {para.runs.length === 0 ? ' ' : null}
              {runSpans(para.runs, 'page-hf-run-img')}
            </div>
          )
        }
        // tab layout happens in left-aligned space; w:jc becomes an explicit
        // shift (a wrapping tail keeps the paragraph's own jc per line)
        const lineStyle = (tabbed: HfTabLayout): React.CSSProperties => {
          const leadIndent = hfLeadIndentCss(tabbed)
          return {
            ...(tabbed.minHeightPt ? { minHeight: `${tabbed.minHeightPt}pt` } : {}),
            ...(tabbed.flow ? {} : { textAlign: 'left' }),
            ...(leadIndent ? { textIndent: leadIndent } : {}),
          }
        }
        const lineClass = (tabbed: HfTabLayout) =>
          `page-hf-tabbed${tabbed.flow ? ' page-hf-tabbed-flow' : ''}`
        const segImgH = (tabbed: HfTabLayout) => hfTabSegImageHeightPx(tabbed)
        const lineContent = (tabbed: HfTabLayout) => (
          <>
            {hfTabLeadNeedsStrut(tabbed) ? '\u200b' : null}
            {segImgH(tabbed) > 0 ? (
              <span className="page-hf-tab-strut" style={{ height: segImgH(tabbed) }} />
            ) : null}
            {runSpans(tabbed.lead, 'page-hf-run-img')}
            {tabbed.flow && tabbed.flow.gapPx > 0 ? (
              <span className="page-hf-tab-gap" style={{ width: tabbed.flow.gapPx }} />
            ) : null}
            {tabbed.flow ? (
              <span key="flow">{runSpans(tabbed.flow.runs, 'page-hf-run-img')}</span>
            ) : null}
            {tabbed.segments.map((seg, k) => (
              <span
                key={`t${k}`}
                className={`page-hf-tabseg page-hf-tabseg-${seg.anchor}${hfSegHasImage(seg) ? ' page-hf-tabseg-img' : ''}`}
                style={{ left: hfSegLeftCss(seg, tabbed) }}
              >
                {runSpans(seg.runs, 'page-hf-run-img')}
              </span>
            ))}
          </>
        )
        const box = paraBox(para, mergeFlags[i])
        const frame = [para.frameXAlign && ' page-hf-frame', ...box.classes.map((c) => ` ${c}`)]
          .filter(Boolean)
          .join('')
        if (tabLines.length === 1) {
          return (
            <div
              key={i}
              className={`page-hf-para ${lineClass(tabLines[0])}${frame}`}
              style={{ ...box.style, ...margins(i), ...lineStyle(tabLines[0]) }}
            >
              {lineContent(tabLines[0])}
            </div>
          )
        }
        // a w:br paragraph stacks one positioned line per break inside the
        // paragraph block (which keeps the spacing and borders)
        return (
          <div key={i} className={`page-hf-para${frame}`} style={{ ...box.style, ...margins(i) }}>
            {tabLines.map((tabbed, m) => (
              <div key={m} className={lineClass(tabbed)} style={lineStyle(tabbed)}>
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
  const owned = (i: number) => !only || only(paras[i], i)
  const drawn = (i: number) => owned(i) && !(only && hfEditablePara(paras[i]))
  const hostedBy = new Map<number, React.ReactNode[]>()
  for (const [n, img] of (images ?? []).filter(hfImageHangsOnPara).entries()) {
    if (img.anchorPara! >= paras.length || !owned(img.anchorPara!)) continue
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
  if (boxGeom) {
    hfObstacleStyles(hfSideWrapObstacles(kind, images, boxGeom)).forEach((style, n) =>
      out.push(<div key={`ob${n}`} className="page-hf-obstacle" style={style} />),
    )
  }
  for (let i = 0; i < paras.length;) {
    const box = paras[i].box
    const css = box && boxGeom ? hfTextBoxStyle(box, kind, boxGeom) : null
    if (!css) {
      if (drawn(i)) out.push(indices[i])
      i += 1
      continue
    }
    const items: React.ReactNode[] = []
    const start = i
    const anchor = hfBoxAnchorEl(box!, indices)
    if (!owned(anchor ?? start)) {
      for (; i < paras.length && paras[i].box?.id === box?.id; i += 1);
      continue
    }
    for (; i < paras.length && paras[i].box?.id === box?.id; i += 1) {
      items.push(renderPara(paras[i], i))
    }
    const node = (
      <div key={`box${start}`} className={hfTextBoxClass(box!)} style={css}>
        {items}
      </div>
    )
    if (anchor != null && drawn(anchor)) {
      hostedBy.set(anchor, [...(hostedBy.get(anchor) ?? []), node])
    } else out.push(node)
  }
  for (const [host, nodes] of hostedBy) {
    if (!drawn(host)) out.push(...nodes)
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
