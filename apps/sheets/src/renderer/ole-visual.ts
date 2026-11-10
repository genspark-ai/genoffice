/**
 * Embedded OLE objects (worksheet `<oleObjects>` + `xl/embeddings/*`).
 *
 * Excel never renders the embedded object itself on the grid; it draws the
 * cached preview picture it stored next to the object (`objectPr/@r:id`, or
 * the legacy VML shape's `v:imagedata` — usually an EMF) at the object's
 * anchor. Only when no preview exists does it fall back to a bordered box
 * with the object's icon and type caption. The sidecar surfaces each object
 * as a read-only `ole` visual; these helpers pick the form to render and the
 * caption text, so the decision is unit-testable without React.
 */

export type OleRenderKind = 'preview' | 'placeholder'

/// Which form an `ole` visual renders as. `previewFailed` is the media
/// fetch/rasterize outcome — a missing or undecodable metafile degrades to
/// the placeholder instead of an empty frame.
export function oleRenderKind(
  visual: { readonly mediaPath?: string | undefined },
  previewFailed: boolean,
): OleRenderKind {
  return visual.mediaPath !== undefined && !previewFailed ? 'preview' : 'placeholder'
}

/// progId prefix → the friendly type name Excel shows under the icon
/// (matched case-insensitively; version suffixes like ".12" are ignored).
const OLE_PROG_ID_NAMES: readonly (readonly [RegExp, string])[] = [
  [/^Word\.Document\b/i, 'Microsoft Word Document'],
  [/^Word\.Template\b/i, 'Microsoft Word Template'],
  [/^Excel\.Sheet\b/i, 'Microsoft Excel Worksheet'],
  [/^Excel\.SheetMacroEnabled\b/i, 'Microsoft Excel Macro-Enabled Worksheet'],
  [/^Excel\.Chart\b/i, 'Microsoft Excel Chart'],
  // Excel 97-2003 embeds carry the bare legacy progId.
  [/^Worksheet$/i, 'Microsoft Excel Worksheet'],
  [/^PowerPoint\.Show\b/i, 'Microsoft PowerPoint Presentation'],
  [/^PowerPoint\.Slide\b/i, 'Microsoft PowerPoint Slide'],
  [/^Visio\.Drawing\b/i, 'Microsoft Visio Drawing'],
  [/^(Acrobat Document|AcroExch\.Document)\b/i, 'Adobe Acrobat Document'],
  [/^Equation\b/i, 'Microsoft Equation'],
  [/^MSGraph\.Chart\b/i, 'Microsoft Graph Chart'],
  [/^(Paint\.Picture|PBrush)\b/i, 'Bitmap Image'],
  [/^Package\b/i, 'Package'],
]

/// Caption for an `ole` visual. Unknown progIds keep their own name minus
/// the numeric version suffix ("Foo.Bar.8" → "Foo.Bar"); an absent progId
/// reads as a generic embedded object.
export function oleCaption(progId: string | undefined): string {
  const id = progId?.trim() ?? ''
  if (id.length === 0) return 'Embedded Object'
  for (const [pattern, name] of OLE_PROG_ID_NAMES) {
    if (pattern.test(id)) return name
  }
  return id.replace(/(\.\d+)+$/, '')
}

/// Inline frame for an `ole` visual from the legacy VML shape's stroke and
/// fill (document colours — Excel's window/windowText defaults, or an explicit
/// `#rrggbb`). A shape with `stroked="f"` / `filled="f"` gets neither; the
/// placeholder then keeps its own themed chrome.
export function oleFrameStyle(visual: {
  readonly lineColor?: string | undefined
  readonly fillColor?: string | undefined
}): { readonly border?: string; readonly background?: string } {
  const style: { border?: string; background?: string } = {}
  if (visual.lineColor !== undefined && visual.lineColor !== 'none') {
    style.border = `1px solid ${visual.lineColor}`
  }
  if (visual.fillColor !== undefined && visual.fillColor !== 'none') {
    style.background = visual.fillColor
  }
  return style
}
