// @ts-nocheck — generation layer ported verbatim from untyped JS; it is typed
// file by file without logic changes, and until then strict consumers
// (apps/html, apps/shell) must not fail on it.
const DEFAULT_PAGE_WIDTH_PX = 794
const DEFAULT_PAGE_HEIGHT_PX = 1123
const MARGIN_DXA = 1134

function createRenderContext(docSettings: any = {}, options: any = {}) {
  // Every numeric here arrives from page JavaScript: NaN/Infinity/negatives
  // must degrade to defaults instead of poisoning page geometry (NaN margins
  // make contentDxa NaN; an infinite page size makes every extent infinite).
  const finiteOr = (v: any, fallback: number): number => (Number.isFinite(v) ? v : fallback)
  const marginsPx = docSettings.marginsPx || {}
  const marginPx = (v: any): number => {
    const n = finiteOr(v, MARGIN_DXA / 15)
    return n < 0 ? MARGIN_DXA / 15 : n
  }
  const pageDim = (v: any, fallback: number): number => (Number.isFinite(v) && v > 0 ? v : fallback)
  const pageWidthDxa = Math.max(
    1440,
    Math.round(pageDim(docSettings.pageSizePx?.width, DEFAULT_PAGE_WIDTH_PX) * 15),
  )
  const pageHeightDxa = Math.max(
    1440,
    Math.round(pageDim(docSettings.pageSizePx?.height, DEFAULT_PAGE_HEIGHT_PX) * 15),
  )
  const viewportW = finiteOr(docSettings.viewportWidthPx, 0)
  const baseScale = viewportW > 0 ? Math.min(1, pageWidthDxa / 15 / viewportW) : 1
  // Extractor-provided whole-document squeeze for flow docs slightly taller
  // than one page: shrinks geometry and fonts together a few percent so a
  // two-line tail does not strand on its own page.
  const rawSqueeze = finiteOr(docSettings.pageFitSqueeze, 1)
  const squeeze = rawSqueeze > 0 && rawSqueeze <= 1 ? rawSqueeze : 1
  const measurementScale = baseScale * squeeze
  // Arial is narrower than the browser's common Inter/Roboto fonts. Scaling
  // it as aggressively as box geometry removes authored line wraps, so retain
  // part of the original font size while still fitting the page.
  const fontScale = Math.min(1, baseScale * 1.04) * squeeze
  const pxToTwips = (px) => Math.round(finiteOr(px, 0) * 15 * measurementScale)
  const pageMargins = {
    top: pxToTwips(marginPx(marginsPx.top)),
    bottom: pxToTwips(marginPx(marginsPx.bottom)),
    left: pxToTwips(marginPx(marginsPx.left)),
    right: pxToTwips(marginPx(marginsPx.right)),
    header: docSettings.composed ? 0 : undefined,
    footer: docSettings.composed ? 0 : undefined,
  }

  return {
    pageWidthDxa,
    pageHeightDxa,
    pageMargins,
    contentDxa: pageWidthDxa - pageMargins.left - pageMargins.right,
    measurementScale,
    fontScale,
    marginDxa: MARGIN_DXA,
    pxToHalfPoints: (px) => Math.round(finiteOr(px, 0) * 1.5 * fontScale),
    pxToTwips,
    pxToBorderEighths: (px) => Math.max(4, Math.round(finiteOr(px, 0) * 6 * measurementScale)),
    naturalTableWidth: Boolean(options.naturalTableWidth),
  }
}

export { createRenderContext }
