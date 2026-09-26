/**
 * The header/footer model every consumer shares (canvas strips, page gaps,
 * pagination heights, print preview, the AI tool, the save path).
 *
 * A section is identified by its break paragraph (lastBlockIndex): positions
 * shift when breaks are inserted or deleted, the paragraph does not. Per
 * (section, kind, variant) the pending state is one of:
 *   - an own part with pending content   → `edits[editKey(...)]`
 *   - an own part deleted (Link to Previous switched on) → `links[linkKey(...)]`
 *   - nothing pending → the parsed reference decides
 * A section with no own part inherits the nearest earlier section's (Word).
 * Sections referencing the same parsed part show one another's pending edit,
 * since the save rewrites that shared part. No section is special-cased.
 */
import type { HeaderFooter, HfImage, HfPartInfo, SectionInfo } from '@genoffice/docx-engine'
import { hfFromPart, type HfVariantKey, type HfVariantsState, type HfView } from './doc-state'

export type HfKind = 'header' | 'footer'

/** parsed document-level strips: what a lone section without references shows */
export interface HfDocStrips {
  header: HeaderFooter | null
  footer: HeaderFooter | null
  variants: HfVariantsState
}

export interface HfParsedParts {
  headerImages?: HfImage[] | null
  footerImages?: HfImage[] | null
  headerFirst?: { images?: HfImage[] } | null
  footerFirst?: { images?: HfImage[] } | null
  headerEven?: { images?: HfImage[] } | null
  footerEven?: { images?: HfImage[] } | null
  hfParts?: Record<string, HfPartInfo>
}

export interface HfSectionState {
  sections: SectionInfo[]
  /** pending own-part content, keyed by editKey */
  edits: Record<string, HeaderFooter>
  /** own part deleted (Link to Previous), keyed by linkKey */
  links: Record<string, true>
  docStrips?: HfDocStrips | null
  parsed?: HfParsedParts | null
}

export interface HfResolved {
  value: HeaderFooter | null
  images?: HfImage[]
  /** section whose part is shown; an edit of this strip writes there */
  owner: number
}

export interface HfCommitTarget {
  /** editKey of the owner's own part */
  key: string
  owner: number
}

export function variantKey(kind: HfKind, variant: Exclude<HfView, 'default'>): HfVariantKey {
  return `${kind}${variant === 'first' ? 'First' : 'Even'}`
}

export function editKey(sectionEnd: number, kind: HfKind, variant: HfView): string {
  return variant === 'default' ? `${sectionEnd}:${kind}` : `${sectionEnd}:${kind}:${variant}`
}

export const linkKey = editKey

export function parseEditKey(key: string): {
  sectionEnd: number
  kind: HfKind
  variant: HfView
} {
  const [end, kind, variant] = key.split(':')
  return {
    sectionEnd: Number(end),
    kind: kind as HfKind,
    variant: (variant as HfView) ?? 'default',
  }
}

export const parseLinkKey = parseEditKey

function refOf(sec: SectionInfo | undefined, kind: HfKind, variant: HfView): string | undefined {
  return sec ? (kind === 'header' ? sec.headerRefs : sec.footerRefs)[variant] : undefined
}

function docStrip(strips: HfDocStrips | null | undefined, kind: HfKind, variant: HfView) {
  if (!strips) return null
  return variant === 'default' ? strips[kind] : strips.variants[variantKey(kind, variant)]
}

function docImages(
  parsed: HfParsedParts | null | undefined,
  kind: HfKind,
  variant: HfView,
): HfImage[] | undefined {
  if (!parsed) return undefined
  if (variant === 'default')
    return (kind === 'header' ? parsed.headerImages : parsed.footerImages) ?? undefined
  const part =
    variant === 'first'
      ? kind === 'header'
        ? parsed.headerFirst
        : parsed.footerFirst
      : kind === 'header'
        ? parsed.headerEven
        : parsed.footerEven
  return part?.images ?? undefined
}

/** pending edit of the parsed part `rId`, made on any section that references it */
function pendingForRef(
  state: HfSectionState,
  kind: HfKind,
  variant: HfView,
  rId: string,
): HeaderFooter | undefined {
  for (const sec of state.sections) {
    if (refOf(sec, kind, variant) !== rId) continue
    if (state.links[linkKey(sec.lastBlockIndex, kind, variant)]) continue
    const edited = state.edits[editKey(sec.lastBlockIndex, kind, variant)]
    if (edited) return edited
  }
  return undefined
}

/** the section's own strip; undefined = it has none and inherits */
export function ownHf(
  state: HfSectionState,
  si: number,
  kind: HfKind,
  variant: HfView,
): { value: HeaderFooter | null; images?: HfImage[] } | undefined {
  const sec = state.sections[si]
  const single = state.sections.length <= 1
  if (!sec && !single) return undefined
  if (sec && state.links[linkKey(sec.lastBlockIndex, kind, variant)]) return undefined
  const rId = refOf(sec, kind, variant)
  const parts = state.parsed?.hfParts
  const images = rId
    ? parts?.[rId]?.images
    : single
      ? docImages(state.parsed, kind, variant)
      : undefined
  const own = sec ? state.edits[editKey(sec.lastBlockIndex, kind, variant)] : undefined
  if (own) return { value: own, images }
  if (rId) {
    const shared = pendingForRef(state, kind, variant, rId)
    return { value: shared ?? hfFromPart(parts?.[rId]), images }
  }
  // a lone section carries the document-level parse (no reference to look up)
  return single ? { value: docStrip(state.docStrips, kind, variant), images } : undefined
}

/** the strip a page of section `si` shows, following inheritance */
export function resolveHf(
  state: HfSectionState,
  si: number,
  kind: HfKind,
  variant: HfView,
): HfResolved {
  const start = Math.min(si, Math.max(0, state.sections.length - 1))
  for (let j = start; j >= 0; j--) {
    const own = ownHf(state, j, kind, variant)
    if (own) return { ...own, owner: j }
  }
  return { value: null, owner: start }
}

/** Word's Link to Previous state; null when there is no previous section */
export function hfLinked(
  state: HfSectionState,
  si: number,
  kind: HfKind,
  variant: HfView,
): boolean | null {
  if (si <= 0 || state.sections.length <= 1) return null
  return ownHf(state, si, kind, variant) === undefined
}

/**
 * Where an edit made on section `si`'s strip lands: the owning section's own
 * part (linked sections write through, like Word). `ownerAtOpen` refuses a write
 * whose owner moved to an earlier section since the editor opened — Link to
 * Previous switched on under an open editor must not overwrite the strip it
 * reconnects to.
 */
export function hfCommitTarget(
  state: HfSectionState,
  si: number,
  kind: HfKind,
  variant: HfView,
  ownerAtOpen?: number,
): HfCommitTarget | null {
  const { owner } = resolveHf(state, si, kind, variant)
  if (ownerAtOpen !== undefined && owner !== ownerAtOpen && owner !== si) return null
  const sec = state.sections[owner]
  if (!sec) return null
  return { key: editKey(sec.lastBlockIndex, kind, variant), owner }
}

/** the pending maps after writing `next` to the strip of section `si` */
export function withHfEdit(
  state: HfSectionState,
  si: number,
  kind: HfKind,
  variant: HfView,
  next: HeaderFooter,
  ownerAtOpen?: number,
): Pick<HfSectionState, 'edits' | 'links'> | null {
  const target = hfCommitTarget(state, si, kind, variant, ownerAtOpen)
  if (!target) return null
  return { edits: { ...state.edits, [target.key]: next }, links: state.links }
}

/**
 * Link to Previous: on = delete the section's own part (its pending edit goes
 * with it; the part it now inherits is untouched); off = the inherited content
 * becomes the section's own part.
 */
export function withHfLink(
  state: HfSectionState,
  si: number,
  kind: HfKind,
  variant: HfView,
  on: boolean,
): Pick<HfSectionState, 'edits' | 'links'> | null {
  const sec = state.sections[si]
  if (!sec || si === 0) return null
  const key = editKey(sec.lastBlockIndex, kind, variant)
  if (on) {
    const { [key]: _dropped, ...edits } = state.edits
    return { edits, links: { ...state.links, [key]: true } }
  }
  const copy = resolveHf(state, si, kind, variant).value ?? { text: '' }
  const { [key]: _unlinked, ...links } = state.links
  return { edits: { ...state.edits, [key]: copy }, links }
}

export interface HfSaveOptions {
  sectionHf?: Array<{
    lastBlockIndex: number
    kind: HfKind
    variant: Exclude<HfView, 'default'> | 'default'
    hf: HeaderFooter
  }>
  sectionHfUnlink?: Array<{
    lastBlockIndex: number
    kind: HfKind
    variant: Exclude<HfView, 'default'> | 'default'
  }>
  header?: HeaderFooter
  footer?: HeaderFooter
  headerFirst?: HeaderFooter
  footerFirst?: HeaderFooter
  headerEven?: HeaderFooter
  footerEven?: HeaderFooter
}

/**
 * The engine's save options for the pending state. The final section's parts
 * go through the trailing-sectPr options (they also carry the watermark and
 * the first/even parts); every other section through `sectionHf`, keyed by
 * its break paragraph. Deleted own parts drop their references.
 */
export function hfSaveOptions(state: HfSectionState): HfSaveOptions {
  const out: HfSaveOptions = {}
  const last = state.sections[state.sections.length - 1]
  const lastEnd = last?.lastBlockIndex
  const sectionHf: NonNullable<HfSaveOptions['sectionHf']> = []
  for (const [key, hf] of Object.entries(state.edits)) {
    const { sectionEnd, kind, variant } = parseEditKey(key)
    if (state.links[key]) continue
    if (sectionEnd === lastEnd) {
      const slot =
        variant === 'default' ? kind : (`${kind}${variant === 'first' ? 'First' : 'Even'}` as const)
      out[slot] = hf
    } else sectionHf.push({ lastBlockIndex: sectionEnd, kind, variant, hf })
  }
  if (sectionHf.length) out.sectionHf = sectionHf
  const unlink = Object.keys(state.links).map((key) => {
    const { sectionEnd, kind, variant } = parseEditKey(key)
    return { lastBlockIndex: sectionEnd, kind, variant }
  })
  if (unlink.length) out.sectionHfUnlink = unlink
  return out
}
