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
import {
  NEW_PICTURE_TOKEN_PREFIX,
  pictureLinkToken,
  retargetImageBlip,
  withPictureLink,
  type HeaderFooter,
  type HfImage,
  type HfParagraph,
  type HfPartInfo,
  type HfPictureEdit,
  type HfPictureLink,
  type Run,
  type SectionInfo,
} from '@genoffice/docx-engine'
import { hfFromPart, type HfVariantKey, type HfVariantsState, type HfView } from './doc-state'
import type { PageFacts } from './editor/field-results'

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

/** results per edit maps and image list: consumers compare image lists by identity */
const editedImagesMemo = new WeakMap<object, WeakMap<object, WeakMap<HfImage[], HfImage[]>>>()
const NO_EDITS = {}

const bytesUrl = (bytes: NonNullable<HfPictureLink['bytes']>): string =>
  `data:${bytes.mime};base64,${bytes.base64}`

/**
 * The part's images with the strip's pending moves / resizes (page-relative boxes) applied,
 * and the bytes Update Now read for a linked one.
 */
export function withHfPictureEdits(
  images: HfImage[] | undefined,
  value: HeaderFooter | null | undefined,
): HfImage[] | undefined {
  const edits = value?.pictureEdits
  const links = value?.pictureLinks
  if (!images || (!edits && !links)) return images
  let byLinks = editedImagesMemo.get(edits ?? NO_EDITS)
  if (!byLinks) editedImagesMemo.set(edits ?? NO_EDITS, (byLinks = new WeakMap()))
  let memo = byLinks.get(links ?? NO_EDITS)
  if (!memo) byLinks.set(links ?? NO_EDITS, (memo = new WeakMap()))
  const hit = memo.get(images)
  if (hit) return hit
  const out = images.map((image): HfImage => {
    const bytes = image.linkDrawing == null ? undefined : links?.[image.linkDrawing]?.bytes
    let img = image
    if (bytes) {
      // the save drops the crop with the old bytes
      const { crop: _old, ...rest } = image
      img = { ...rest, dataUrl: bytesUrl(bytes) }
    }
    const edit = img.drawing == null ? undefined : edits?.[img.drawing]
    if (!edit) return img
    // an axis the picture moved along is pinned to the page; the other keeps
    // its base (an aligned picture realigns at its new size, as in Word)
    let next: HfImage = { ...img, widthPx: edit.widthPx, heightPx: edit.heightPx }
    if (edit.posXPx != null) {
      const { posH: _h, ...rest } = next
      next = { ...rest, posXPx: edit.posXPx, posHRel: 'page' }
    }
    if (edit.posYPx != null) {
      const { posV: _v, anchorPara: _a, ...rest } = next
      next = { ...rest, posYPx: edit.posYPx, posVRel: 'page' }
    }
    return next
  })
  memo.set(images, out)
  return out
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
  if (own) return { value: own, images: withHfPictureEdits(images, own) }
  if (rId) {
    const shared = pendingForRef(state, kind, variant, rId)
    return {
      value: shared ?? hfFromPart(parts?.[rId]),
      images: withHfPictureEdits(images, shared),
    }
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
  return {
    edits: { ...state.edits, [target.key]: hfTextCommit(state.edits[target.key], next) },
    links: state.links,
  }
}

/**
 * A content commit over a strip whose pictures may have moved: the moves stay
 * (the strip editor's value predates the ones made while it was open) and
 * the text now counts as edited.
 */
export function hfTextCommit(prev: HeaderFooter | undefined, next: HeaderFooter): HeaderFooter {
  const { pictureOnly: _textKept, ...rest } = next
  const pictureEdits = prev?.pictureEdits ?? next.pictureEdits
  const pictureLinks = prev?.pictureLinks ?? next.pictureLinks
  // a commit that carries the runs but not what their placeholders stand for keeps the bytes
  const newPictures =
    prev?.newPictures || next.newPictures
      ? { ...prev?.newPictures, ...next.newPictures }
      : undefined
  return {
    ...rest,
    ...(pictureEdits ? { pictureEdits } : {}),
    ...(pictureLinks ? { pictureLinks } : {}),
    ...(newPictures ? { newPictures } : {}),
  }
}

/** one linked picture of a header or footer, as Edit > Links lists it */
export interface HfPictureLinkRow {
  /** the first section whose strip shows the part */
  si: number
  kind: HfKind
  variant: HfView
  drawing: number
  target: string
  docPr?: string
}

const HF_KINDS: readonly HfKind[] = ['header', 'footer']
const HF_VIEWS: readonly HfView[] = ['default', 'first', 'even']

const docPrOf = (xml: string): string | undefined => /<wp:docPr\s[^>]*\bid="([^"]+)"/.exec(xml)?.[1]

const hfRuns = (paras: HfParagraph[] | undefined): Run[] =>
  (paras ?? []).flatMap((p) => [...p.runs, ...(p.cells?.flatMap((c) => c.paras.flat()) ?? [])])

const holdsPicture = (paras: HfParagraph[] | undefined, docPr: string): boolean =>
  hfRuns(paras).some((run) => run.image && docPrOf(run.image.xml) === docPr)

/**
 * The linked pictures of every header and footer part a section shows, each part once, with
 * what Edit > Links did since: a broken link is gone, a changed one names its new file. A
 * picture that stood in the part's text and was deleted from it is gone too.
 */
export function hfPictureLinkRows(state: HfSectionState): HfPictureLinkRow[] {
  const parts = state.parsed?.hfParts
  if (!parts) return []
  const rows: HfPictureLinkRow[] = []
  const seen = new Set<string>()
  state.sections.forEach((sec, si) => {
    for (const kind of HF_KINDS)
      for (const variant of HF_VIEWS) {
        const rId = refOf(sec, kind, variant)
        const linked = rId && !seen.has(rId) ? parts[rId]?.linkedPictures : undefined
        if (!rId || !linked || state.links[linkKey(sec.lastBlockIndex, kind, variant)]) continue
        seen.add(rId)
        const pending =
          state.edits[editKey(sec.lastBlockIndex, kind, variant)] ??
          pendingForRef(state, kind, variant, rId)
        const rewritten = pending && !pending.pictureOnly && !pending.keepContent
        for (const picture of linked) {
          const { docPr } = picture
          const deleted =
            rewritten &&
            docPr !== undefined &&
            holdsPicture(parts[rId].paras, docPr) &&
            !holdsPicture(pending.paras, docPr)
          const edit = pending?.pictureLinks?.[picture.drawing]
          const target = edit ? edit.target : picture.target
          if (target !== null && !deleted)
            rows.push({
              si,
              kind,
              variant,
              drawing: picture.drawing,
              target,
              ...(docPr ? { docPr } : {}),
            })
        }
      }
  })
  return rows
}

// its own range: the placeholders the save numbers, the ones held on nodes and the engine's start lower
const LINK_RUN_TOKEN_BASE = 3_000_000
let linkRunSeq = 0

/**
 * The part's paragraphs with a link change on the run that holds the picture: a part written
 * anew writes its runs, not the bytes the file had. `bytes` collects what new placeholders stand for.
 */
function withLinkOnRuns(
  paras: HfParagraph[] | undefined,
  link: HfPictureLink,
  bytes: NonNullable<HeaderFooter['newPictures']>,
): HfParagraph[] | undefined {
  const { docPr } = link
  if (!paras || docPr === undefined || !holdsPicture(paras, docPr)) return paras
  const onRun = (run: Run): Run => {
    const image = run.image
    if (!image || docPrOf(image.xml) !== docPr) return run
    let { xml, dataUrl } = image
    const { linkTarget: _was, vectorUrl, ...rest } = image
    const fresh = link.bytes && bytesUrl(link.bytes) !== dataUrl
    if (link.bytes && fresh) {
      const token = `${NEW_PICTURE_TOKEN_PREFIX}${LINK_RUN_TOKEN_BASE + ++linkRunSeq}`
      xml = retargetImageBlip(xml, token)
      dataUrl = bytesUrl(link.bytes)
      bytes[token] = link.bytes
    }
    xml = withPictureLink(xml, link.target === null ? null : pictureLinkToken(link.target))
    return {
      ...run,
      image: {
        ...rest,
        xml,
        dataUrl,
        ...(vectorUrl && !fresh ? { vectorUrl } : {}),
        ...(link.target === null ? {} : { linkTarget: link.target }),
      },
    }
  }
  return paras.map((p) => ({
    ...p,
    runs: p.runs.map(onRun),
    ...(p.cells
      ? { cells: p.cells.map((c) => ({ ...c, paras: c.paras.map((runs) => runs.map(onRun)) })) }
      : {}),
  }))
}

/**
 * The pending maps after Edit > Links changed a linked picture of section `si`'s strip:
 * `target` is the file it names from now on (null breaks the link), `bytes` what Update Now
 * read. A strip with no pending content keeps its text bytes (pictureOnly).
 */
export function withHfPictureLink(
  state: HfSectionState,
  row: Pick<HfPictureLinkRow, 'si' | 'kind' | 'variant' | 'drawing' | 'docPr'>,
  change: Pick<HfPictureLink, 'target' | 'bytes'>,
): Pick<HfSectionState, 'edits' | 'links'> | null {
  const { si, kind, variant, drawing, docPr } = row
  const target = hfCommitTarget(state, si, kind, variant)
  if (!target) return null
  const resolved = resolveHf(state, si, kind, variant).value ?? { text: '' }
  const cur =
    state.edits[target.key] ??
    (Object.values(state.edits).includes(resolved) ? resolved : { ...resolved, pictureOnly: true })
  // new bytes stay the picture's own once its link names another file or is broken
  const bytes = change.bytes ?? cur.pictureLinks?.[drawing]?.bytes
  const link: HfPictureLink = {
    target: change.target,
    ...(bytes ? { bytes } : {}),
    ...(docPr ? { docPr } : {}),
  }
  const placed: NonNullable<HeaderFooter['newPictures']> = {}
  const paras = withLinkOnRuns(cur.paras, link, placed)
  const newPictures = { ...cur.newPictures, ...placed }
  return {
    edits: {
      ...state.edits,
      [target.key]: {
        ...cur,
        ...(paras ? { paras } : {}),
        ...(Object.keys(newPictures).length > 0 ? { newPictures } : {}),
        pictureLinks: { ...cur.pictureLinks, [drawing]: link },
      },
    },
    links: state.links,
  }
}

/**
 * The pending maps after floating picture `drawing` of section `si`'s strip
 * was moved or resized. A strip with no pending content keeps its text bytes
 * (pictureOnly); one with an edit takes the move alongside it.
 */
export function withHfPictureEdit(
  state: HfSectionState,
  si: number,
  kind: HfKind,
  variant: HfView,
  drawing: number,
  edit: HfPictureEdit,
): Pick<HfSectionState, 'edits' | 'links'> | null {
  const target = hfCommitTarget(state, si, kind, variant)
  if (!target) return null
  const resolved = resolveHf(state, si, kind, variant).value ?? { text: '' }
  const cur =
    state.edits[target.key] ??
    (Object.values(state.edits).includes(resolved) ? resolved : { ...resolved, pictureOnly: true })
  // a resize after a move keeps the move: an axis the edit leaves out stays as it was
  const prev = cur.pictureEdits?.[drawing]
  const merged: HfPictureEdit = { widthPx: edit.widthPx, heightPx: edit.heightPx }
  const x = edit.posXPx ?? prev?.posXPx
  const y = edit.posYPx ?? prev?.posYPx
  if (x != null) merged.posXPx = x
  if (y != null) merged.posYPx = y
  return {
    edits: {
      ...state.edits,
      [target.key]: { ...cur, pictureEdits: { ...cur.pictureEdits, [drawing]: merged } },
    },
    links: state.links,
  }
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

/** each rewritten part stores the numbers of the first page the layout shows it on */
export function withHfFieldResults(
  options: HfSaveOptions,
  lastEnd: number | undefined,
  pageOf: ((sectionEnd: number, variant: HfView) => PageFacts | null) | undefined,
): HfSaveOptions {
  if (!pageOf) return options
  const settled = (hf: HeaderFooter, sectionEnd: number, variant: HfView): HeaderFooter => {
    const page = pageOf(sectionEnd, variant)
    return page
      ? { ...hf, fieldResults: { page: page.display, sectionPages: page.sectionPages } }
      : hf
  }
  const out = { ...options }
  if (lastEnd !== undefined)
    for (const kind of ['header', 'footer'] as const)
      for (const [variant, suffix] of [
        ['default', ''],
        ['first', 'First'],
        ['even', 'Even'],
      ] as const) {
        const hf = out[`${kind}${suffix}`]
        if (hf) out[`${kind}${suffix}`] = settled(hf, lastEnd, variant)
      }
  if (out.sectionHf)
    out.sectionHf = out.sectionHf.map((e) => ({
      ...e,
      hf: settled(e.hf, e.lastBlockIndex, e.variant),
    }))
  return out
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
