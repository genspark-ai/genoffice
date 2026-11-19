/**
 * Page color (w:background) makes Word lay out one extra empty paragraph in
 * the default paragraph style after the header content of section 1 (every
 * variant) and of later sections inheriting it; never footers, never a section
 * with its own header part; a part-less section 1 gets an empty Header-style
 * line first and a part-less footer one empty Footer-style line. Every one of
 * them inherits its style chain's pBdr (docDefaults `none space` pads the
 * strip like any paragraph). Layout-only: appended to the values the
 * push-down measure and the page strips consume, never to the editable or
 * saved state.
 */
import {
  mergeStyleBorders,
  type DocDefaults,
  type HeaderFooter,
  type HfImage,
  type HfParagraph,
  type StyleDisplay,
  type StyleInfo,
} from '@genoffice/docx-engine'
import { hfParasOf } from './editor/hf-text'
import type { HfKind, HfResolved } from './hf-sections'

export interface HfPhantomSpec {
  normal: HfParagraph
  headerLine: HfParagraph
  footerLine: HfParagraph
}

interface StyleSource {
  styles: Map<string, StyleInfo>
  docDefaults?: DocDefaults
}

function blankPara(d: StyleDisplay | undefined, dd: DocDefaults | undefined): HfParagraph {
  const line = d?.lineRule || d?.lineSpacing || d?.lineRawTwips ? d : dd
  const borderSides = d?.borderSides ?? dd?.borderSides
  return {
    runs: [],
    ...(borderSides ? mergeStyleBorders(borderSides, undefined) : {}),
    align: 'left',
    // no size anywhere: Word's built-in default is 10pt
    emptyRunSizeHalfPoints: d?.sizeHalfPoints ?? dd?.sizeHalfPoints ?? 20,
    ...((d?.fontAscii ?? dd?.asciiFont ?? d?.font)
      ? { emptyRunFontFamily: d?.fontAscii ?? dd?.asciiFont ?? d?.font }
      : {}),
    ...(line?.lineRule ? { lineRule: line.lineRule } : {}),
    ...(line?.lineRawTwips ? { lineRawTwips: line.lineRawTwips } : {}),
    ...(line?.lineSpacing ? { lineSpacing: line.lineSpacing } : {}),
    spaceBefore: d?.spaceBeforeTwips ?? dd?.spaceBeforeTwips ?? 0,
    spaceAfter: d?.spaceAfterTwips ?? dd?.spaceAfterTwips ?? 0,
  }
}

/** the phantom paragraph (and the header line a part-less section 1 gets) from the document's styles */
export function hfPhantomSpec(parsed: StyleSource): HfPhantomSpec {
  let normal: StyleInfo | undefined
  let header: StyleInfo | undefined
  let footer: StyleInfo | undefined
  for (const s of parsed.styles.values()) {
    if (s.type !== 'paragraph') continue
    if (s.isDefault) normal = s
    else if (s.styleId === 'Header' || /^header$/i.test(s.name)) header = s
    else if (s.styleId === 'Footer' || /^footer$/i.test(s.name)) footer = s
  }
  const dd = parsed.docDefaults
  return {
    normal: blankPara(normal?.display, dd),
    headerLine: blankPara(header?.display ?? normal?.display, dd),
    footerLine: blankPara(footer?.display ?? normal?.display, dd),
  }
}

/** a part-less footer: Word lays out one empty Footer-style paragraph */
export function hfFooterLine(value: HeaderFooter | null, spec: HfPhantomSpec): HeaderFooter {
  return value ?? { text: '', paras: [spec.footerLine] }
}

/** `value` with the phantom paragraph appended (null = a section 1 without a header part) */
export function hfWithPhantom(
  value: HeaderFooter | null,
  spec: HfPhantomSpec,
  images?: HfImage[] | null,
): HeaderFooter {
  const paras = value ? hfParasOf(value, images) : [spec.headerLine]
  const flow = paras.filter((p) => !p.boxAnchored)
  const last = flow[flow.length - 1]
  const lastAfter = last && !last.cells ? (last.spaceAfter ?? 0) : 0
  // Word merges the header's after with Normal's before by max; the strip stacks them
  const spaceBefore = Math.max(0, (spec.normal.spaceBefore ?? 0) - lastAfter)
  return { ...(value ?? { text: '' }), paras: [...paras, { ...spec.normal, spaceBefore }] }
}

/**
 * The strip a page lays out: the resolved strip plus the phantom when it is a
 * header owned by (or inherited from) section 1. A header found nowhere
 * inherits section 1's part-less story, which Word still lays out.
 */
export function hfLayoutResolved(
  kind: HfKind,
  resolved: HfResolved,
  spec: HfPhantomSpec | null,
): HfResolved {
  if (!spec) return resolved
  if (kind === 'footer') return { ...resolved, value: hfFooterLine(resolved.value, spec) }
  if (resolved.value && resolved.owner !== 0) return resolved
  return { ...resolved, value: hfWithPhantom(resolved.value, spec, resolved.images) }
}
