// @ts-nocheck — generation layer ported verbatim from untyped JS; it is typed
// file by file without logic changes, and until then strict consumers
// (apps/html, apps/shell) must not fail on it.
const DEFAULT_PAGE_WIDTH_PX = 794
const DEFAULT_PAGE_HEIGHT_PX = 1123
const MARGIN_DXA = 1134

function createRenderContext(docSettings: any = {}) {
  const marginsPx = docSettings.marginsPx || {}
  const pageWidthDxa = Math.max(
    1440,
    Math.round((docSettings.pageSizePx?.width || DEFAULT_PAGE_WIDTH_PX) * 15),
  )
  const pageHeightDxa = Math.max(
    1440,
    Math.round((docSettings.pageSizePx?.height || DEFAULT_PAGE_HEIGHT_PX) * 15),
  )
  const baseScale = docSettings.viewportWidthPx
    ? Math.min(1, pageWidthDxa / 15 / docSettings.viewportWidthPx)
    : 1
  // Extractor-provided whole-document squeeze for flow docs slightly taller
  // than one page: shrinks geometry and fonts together a few percent so a
  // two-line tail does not strand on its own page.
  const squeeze = docSettings.pageFitSqueeze || 1
  const measurementScale = baseScale * squeeze
  // Arial is narrower than the browser's common Inter/Roboto fonts. Scaling
  // it as aggressively as box geometry removes authored line wraps, so retain
  // part of the original font size while still fitting the page.
  const fontScale = Math.min(1, baseScale * 1.04) * squeeze
  const pxToTwips = (px) => Math.round(px * 15 * measurementScale)
  const pageMargins = {
    top: pxToTwips(marginsPx.top ?? MARGIN_DXA / 15),
    bottom: pxToTwips(marginsPx.bottom ?? MARGIN_DXA / 15),
    left: pxToTwips(marginsPx.left ?? MARGIN_DXA / 15),
    right: pxToTwips(marginsPx.right ?? MARGIN_DXA / 15),
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
    pxToHalfPoints: (px) => Math.round(px * 1.5 * fontScale),
    pxToTwips,
    pxToBorderEighths: (px) => Math.max(4, Math.round(px * 6 * measurementScale)),
  }
}

export { createRenderContext }
