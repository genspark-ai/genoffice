/**
 * Element-level hyperlinks — <a:hlinkClick> under <p:cNvPr>.
 *
 * Two kinds of targets:
 * - External URL: a hyperlink relationship with TargetMode="External" in the slide rels;
 * - In-document slide jump: hlinkClick with action="ppaction://hlinksldjump",
 *   whose relationship (type=slide) points at the target slideN.xml.
 *
 * Implemented via "XML surgery + materialize": edit the cNvPr in the element's
 * current fragment, then reparse the whole slide (same path as appendRawElements).
 */
import type { GroupElement, Slide } from './types'
import { escapeXmlAttr } from './xml-utils'
import { sliceGroupChildXmls } from './parse'
import { relsPathFor, resolveTarget } from './zip'
import { cleanupSupersededSlideResources } from './resource-cleanup'
import { materializeSlide, patchedElementXml, patchSlideXml, type OpenedPptx } from './index'

const HYPERLINK_REL_TYPE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'
const SLIDE_REL_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide'
const SLDJUMP_ACTION = 'ppaction://hlinksldjump'
const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

export type LinkTarget = { kind: 'url'; url: string } | { kind: 'slide'; slideIndex: number }

const EMPTY_RELS =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>'

/** Append a relationship to the slide rels and return the new rId. */
function appendRel(
  opened: OpenedPptx,
  slide: Slide,
  type: string,
  target: string,
  external: boolean,
): string {
  const { archive } = opened
  const relsPath = relsPathFor(slide.path)
  const rels = archive.readText(relsPath) ?? EMPTY_RELS
  let maxRid = 0
  for (const m of rels.matchAll(/Id="rId(\d+)"/g)) maxRid = Math.max(maxRid, Number(m[1]))
  const rid = `rId${maxRid + 1}`
  const mode = external ? ' TargetMode="External"' : ''
  const relXml = `<Relationship Id="${rid}" Type="${type}" Target="${escapeXmlAttr(target)}"${mode}/>`
  archive.entries.set(
    relsPath,
    Buffer.from(rels.replace('</Relationships>', `${relXml}</Relationships>`), 'utf8'),
  )
  return rid
}

/** Strip the hlinkClick inside the element XML's first cNvPr (keeping other children). */
function stripHlink(xml: string): string {
  return xml.replace(/<a:hlinkClick\b[^>]*\/>|<a:hlinkClick\b[^>]*>[\s\S]*?<\/a:hlinkClick>/, '')
}

/**
 * Set/clear an element hyperlink. Relationships no longer used by any live slide
 * XML are pruned after the fragment changes. On success returns the materialized
 * new slide model (all element ids refreshed); on failure returns null.
 */
export function setElementLink(
  opened: OpenedPptx,
  slideIndex: number,
  elementId: string,
  target: LinkTarget | null,
): Slide | null {
  const slide = opened.deck.slides[slideIndex]
  if (!slide) return null
  const el = slide.elements.find((e) => e.id === elementId)
  if (!el) return null

  const previousXml = patchSlideXml(slide)
  let xml = stripHlink(patchedElementXml(el))

  if (target) {
    let hlink: string
    if (target.kind === 'url') {
      const rid = appendRel(opened, slide, HYPERLINK_REL_TYPE, target.url, true)
      hlink = `<a:hlinkClick xmlns:r="${R_NS}" r:id="${rid}"/>`
    } else {
      const dst = opened.deck.slides[target.slideIndex]
      if (!dst) return null
      // Same directory ppt/slides/, so Target is just the file name
      const fileName = dst.path.split('/').pop()!
      const rid = appendRel(opened, slide, SLIDE_REL_TYPE, fileName, false)
      hlink = `<a:hlinkClick xmlns:r="${R_NS}" r:id="${rid}" action="${SLDJUMP_ACTION}"/>`
    }
    // Insert into the first cNvPr (hlinkClick is cNvPr's first valid child)
    const cNvPr = /<p:cNvPr\b((?:"[^"]*"|'[^']*'|[^"'>])*?)(\/?)>/.exec(xml)
    if (!cNvPr) return null
    if (cNvPr[2] === '/') {
      // Self-closing → expand into a tag pair
      xml =
        xml.slice(0, cNvPr.index) +
        `<p:cNvPr${cNvPr[1]}>${hlink}</p:cNvPr>` +
        xml.slice(cNvPr.index + cNvPr[0].length)
    } else {
      const at = cNvPr.index + cNvPr[0].length
      xml = xml.slice(0, at) + hlink + xml.slice(at)
    }
  }

  el.anchor = { ...el.anchor, originalXml: xml }
  // Post-surgery fragment is the sole truth: clear text dirty so patches don't overwrite; structural rebuild persists it
  el.dirty = false
  slide.structureDirty = true
  cleanupSupersededSlideResources(opened, slide, previousXml, patchSlideXml(slide))
  return materializeSlide(opened, slideIndex)
}

/** Parse a TextRun.hyperlink encoded target ("slide:N" or url); null on bad input. */
export function decodeRunLink(s: string): LinkTarget | null {
  const m = /^slide:(\d+)$/.exec(s)
  if (m) return { kind: 'slide', slideIndex: Number(m[1]) }
  return s ? { kind: 'url', url: s } : null
}

/** TextRun.hyperlink encoding of a link target. */
export function encodeRunLink(target: LinkTarget): string {
  return target.kind === 'slide' ? `slide:${target.slideIndex}` : target.url
}

/**
 * Allocate slide-rels relationships for runs whose hyperlink was set this session
 * (hyperlink present, hyperlinkRId absent — the edit path clears the rId on change).
 * The caller reconciles the before/after slide XML after its text patch so cleared
 * or replaced relationship ids can be removed safely. Returns whether anything
 * was allocated.
 */
export function ensureRunLinkRels(
  opened: OpenedPptx,
  slideIndex: number,
  paragraphs: Array<{
    runs: Array<{ hyperlink?: string; hyperlinkRId?: string; hyperlinkAction?: string }>
  }>,
): boolean {
  const slide = opened.deck.slides[slideIndex]
  if (!slide) return false
  let changed = false
  for (const p of paragraphs) {
    for (const run of p.runs) {
      if (!run.hyperlink || run.hyperlinkRId) continue
      const target = decodeRunLink(run.hyperlink)
      if (!target) continue
      if (target.kind === 'url') {
        run.hyperlinkRId = appendRel(opened, slide, HYPERLINK_REL_TYPE, target.url, true)
        delete run.hyperlinkAction
      } else {
        const dst = opened.deck.slides[target.slideIndex]
        if (!dst) continue
        run.hyperlinkRId = appendRel(
          opened,
          slide,
          SLIDE_REL_TYPE,
          dst.path.split('/').pop()!,
          false,
        )
        run.hyperlinkAction = SLDJUMP_ACTION
      }
      changed = true
    }
  }
  return changed
}

/**
 * All run-level hyperlinks on a slide, resolved live against the rels (parse-time
 * TextRun.hyperlink can go stale after slide reorders). Keyed by element + paragraph +
 * run indexes; the slideshow matches these against layout glyph runs to hit-test clicks.
 */
export function getRunLinks(
  opened: OpenedPptx,
  slideIndex: number,
): Array<{ elementId: string; paraIndex: number; runIndex: number; target: LinkTarget }> {
  const slide = opened.deck.slides[slideIndex]
  if (!slide) return []
  const out: Array<{ elementId: string; paraIndex: number; runIndex: number; target: LinkTarget }> =
    []
  const rels = opened.archive.readRels(slide.path)
  const resolve = (rid: string): LinkTarget | null => {
    const rel = rels.get(rid)
    if (!rel) return null
    if (rel.type === HYPERLINK_REL_TYPE) return { kind: 'url', url: rel.target }
    if (rel.type === SLIDE_REL_TYPE) {
      const abs = resolveTarget(slide.path, rel.target)
      const idx = opened.deck.slides.findIndex((s) => s.path === abs)
      if (idx >= 0) return { kind: 'slide', slideIndex: idx }
    }
    return null
  }
  const walk = (elements: Slide['elements']) => {
    for (const el of elements) {
      if (el.type === 'group') {
        walk((el as GroupElement).children)
        continue
      }
      if ((el.type !== 'text' && el.type !== 'shape') || !('text' in el) || !el.text) continue
      el.text.paragraphs.forEach((p, paraIndex) => {
        p.runs.forEach((run, runIndex) => {
          if (!run.hyperlinkRId) return
          const target = resolve(run.hyperlinkRId)
          if (target) out.push({ elementId: el.id, paraIndex, runIndex, target })
        })
      })
    }
  }
  walk(slide.elements)
  return out
}

/** Resolve the hlinkClick rId in an element fragment against the slide rels; null if none/unresolvable. */
function resolveLinkInXml(opened: OpenedPptx, slide: Slide, xml: string): LinkTarget | null {
  const m = /<a:hlinkClick\b[^>]*\br:id="(rId\d+)"/.exec(xml)
  if (!m) return null
  const rel = opened.archive.readRels(slide.path).get(m[1]!)
  if (!rel) return null
  if (rel.type === HYPERLINK_REL_TYPE) return { kind: 'url', url: rel.target }
  if (rel.type === SLIDE_REL_TYPE) {
    const abs = resolveTarget(slide.path, rel.target)
    const idx = opened.deck.slides.findIndex((s) => s.path === abs)
    if (idx >= 0) return { kind: 'slide', slideIndex: idx }
  }
  return null
}

/** Read an element's current hyperlink (for UI echo-back); returns null if none. */
export function getElementLink(
  opened: OpenedPptx,
  slideIndex: number,
  elementId: string,
): LinkTarget | null {
  const slide = opened.deck.slides[slideIndex]
  const el = slide?.elements.find((e) => e.id === elementId)
  if (!slide || !el) return null
  return resolveLinkInXml(opened, slide, el.anchor.originalXml)
}

/**
 * All element hyperlinks on a slide (groups scanned recursively) — the slideshow
 * uses this to hit-test clicks and follow slide jumps (Zoom) / external URLs.
 * Group children have empty anchors (their bytes live only inside the group's
 * fragment), so each child's XML is sliced out of the group XML by index.
 */
export function getSlideLinks(
  opened: OpenedPptx,
  slideIndex: number,
): Array<{ elementId: string; target: LinkTarget }> {
  const slide = opened.deck.slides[slideIndex]
  if (!slide) return []
  const out: Array<{ elementId: string; target: LinkTarget }> = []
  const walk = (elements: Slide['elements'], xmlOf: (i: number) => string) => {
    for (let i = 0; i < elements.length; i++) {
      const el = elements[i]!
      const xml = xmlOf(i)
      if (el.type === 'group') {
        // Restrict the group's own match to its nvGrpSpPr header so a child link
        // doesn't make the whole group clickable
        const own = /<p:nvGrpSpPr>[\s\S]*?<\/p:nvGrpSpPr>/.exec(xml)?.[0] ?? ''
        const target = resolveLinkInXml(opened, slide, own)
        if (target) out.push({ elementId: el.id, target })
        // Child fragments are in document order, matching (grp as GroupElement).children
        const childXmls = sliceGroupChildXmls(xml)
        walk((el as GroupElement).children, (j) => childXmls[j] ?? '')
        continue
      }
      const target = resolveLinkInXml(opened, slide, xml)
      if (target) out.push({ elementId: el.id, target })
    }
  }
  walk(slide.elements, (i) => slide.elements[i]!.anchor.originalXml)
  return out
}
