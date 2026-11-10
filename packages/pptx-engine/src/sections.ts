/**
 * Section read/write — modeled after PowerPoint's "sections".
 *
 * OOXML storage: inside <p:extLst> at the end of ppt/presentation.xml:
 *   <p:ext uri="{521415D9-36F7-43E2-AB2F-B90AF26B5E84}">
 *     <p14:sectionLst xmlns:p14="…/powerpoint/2010/main">
 *       <p14:section name="…" id="{GUID}">
 *         <p14:sldIdLst><p14:sldId id="256"/>…</p14:sldIdLst>
 *       </p14:section>…
 *
 * Note: a p14:sldId's id is the numeric slide id from p:sldIdLst (what PowerPoint
 * actually writes), not an rId; parsing tolerates the odd generator that writes
 * rIds by mistake.
 *
 * After slide-order changes (adding/deleting slides), section data may be stale
 * (dangling references / incomplete coverage), so every edit operation first does
 * a positional normalization: each section covers [its first slide, next
 * section's first slide), and slides before the first section count as
 * unsectioned — an intuitive rule that naturally self-heals stale
 * references.
 */
import { randomUUID } from 'node:crypto'
import { resolveTarget } from './zip'
import { escapeXmlAttr } from './xml-utils'
import { unescapeXml } from './notes'
import type { OpenedPptx } from './index'

export interface SectionInfo {
  /** Section GUID (with braces, PowerPoint style) */
  id: string
  name: string
  /** Indices in deck.slides of the slides in this section (in slide order) */
  slideIndices: number[]
}

const PRES_PATH = 'ppt/presentation.xml'
/** Fixed uri of the extension hosting sectionLst (defined in [MS-PPTX]) */
const SECTION_EXT_URI = '{521415D9-36F7-43E2-AB2F-B90AF26B5E84}'
const P14_NS = 'http://schemas.microsoft.com/office/powerpoint/2010/main'

/** New section GUID: uppercase with braces (the format found in pptx files). */
function newSectionId(): string {
  return `{${randomUUID().toUpperCase()}}`
}

/** Index array for [start, end). */
function range(start: number, end: number): number[] {
  const out: number[] = []
  for (let i = start; i < end; i++) out.push(i)
  return out
}

/** Ordered entries of presentation.xml sldIdLst: numeric sldId / rId / deck index. */
interface SlideEntry {
  sldId: string
  rId: string
  deckIndex: number
}

function readSlideEntries(opened: OpenedPptx): SlideEntry[] {
  const { archive, deck } = opened
  const pres = archive.readText(PRES_PATH)
  if (!pres) return []
  const inner = /<p:sldIdLst>([\s\S]*?)<\/p:sldIdLst>/.exec(pres)?.[1] ?? ''
  const rels = archive.readRels(PRES_PATH)
  const out: SlideEntry[] = []
  for (const m of inner.matchAll(/<p:sldId\b[^>]*\/>/g)) {
    const sldId = /\bid="(\d+)"/.exec(m[0])?.[1]
    const rId = /\br:id="([^"]+)"/.exec(m[0])?.[1]
    if (!sldId || !rId) continue
    const rel = rels.get(rId)
    const path = rel ? resolveTarget(PRES_PATH, rel.target) : undefined
    const deckIndex = path ? deck.slides.findIndex((s) => s.path === path) : -1
    out.push({ sldId, rId, deckIndex })
  }
  return out
}

/**
 * Parse presentation.xml's p14:sectionLst. Returns [] if there are no sections
 * (older documents). sldIds referencing deleted slides are silently skipped.
 */
export function getSections(opened: OpenedPptx): SectionInfo[] {
  const pres = opened.archive.readText(PRES_PATH)
  if (!pres) return []
  const lst = /<p14:sectionLst[^>]*>([\s\S]*?)<\/p14:sectionLst>/.exec(pres)?.[1]
  if (!lst) return []
  const entries = readSlideEntries(opened)
  const bySldId = new Map(entries.map((e) => [e.sldId, e.deckIndex]))
  const byRid = new Map(entries.map((e) => [e.rId, e.deckIndex]))
  const out: SectionInfo[] = []
  for (const m of lst.matchAll(/<p14:section\b([^>]*?)(?:\/>|>([\s\S]*?)<\/p14:section>)/g)) {
    const attrs = m[1] ?? ''
    const inner = m[2] ?? ''
    const name = unescapeXml(/\bname="([^"]*)"/.exec(attrs)?.[1] ?? '')
    const id = /\bid="([^"]+)"/.exec(attrs)?.[1] ?? newSectionId()
    const slideIndices: number[] = []
    for (const sm of inner.matchAll(/<p14:sldId\s[^>]*id="([^"]+)"/g)) {
      const idx = bySldId.get(sm[1]!) ?? byRid.get(sm[1]!) ?? -1
      if (idx >= 0) slideIndices.push(idx)
    }
    out.push({ id, name, slideIndices })
  }
  return out
}

/**
 * Write the section structure back into presentation.xml's extLst (created if
 * missing). With empty sections the whole sectionLst extension is removed (and
 * the extLst too if it becomes empty). Only presentation.xml is touched;
 * [Content_Types]/rels are unaffected.
 */
export function setSections(opened: OpenedPptx, sections: SectionInfo[]): void {
  const { archive } = opened
  const pres = archive.readText(PRES_PATH)
  if (!pres) return
  const byIndex = new Map(readSlideEntries(opened).map((e) => [e.deckIndex, e.sldId]))

  // First remove the whole p:ext hosting the existing sectionLst; drop the extLst too if it becomes empty
  let next = pres.replace(
    /<p:ext\b[^>]*>\s*<p14:sectionLst[\s\S]*?<\/p14:sectionLst>\s*<\/p:ext>/,
    '',
  )
  next = next.replace(/<p:extLst>\s*<\/p:extLst>/, '')

  if (sections.length) {
    const secXml = sections
      .map((s) => {
        const ids = s.slideIndices.map((i) => byIndex.get(i)).filter((v): v is string => v != null)
        const lst = ids.length
          ? `<p14:sldIdLst>${ids.map((id) => `<p14:sldId id="${id}"/>`).join('')}</p14:sldIdLst>`
          : '<p14:sldIdLst/>'
        return `<p14:section name="${escapeXmlAttr(s.name)}" id="${s.id || newSectionId()}">${lst}</p14:section>`
      })
      .join('')
    const ext = `<p:ext uri="${SECTION_EXT_URI}"><p14:sectionLst xmlns:p14="${P14_NS}">${secXml}</p14:sectionLst></p:ext>`
    if (/<\/p:extLst>/.test(next)) {
      next = next.replace('</p:extLst>', () => `${ext}</p:extLst>`)
    } else if (/<p:extLst\/>/.test(next)) {
      next = next.replace('<p:extLst/>', () => `<p:extLst>${ext}</p:extLst>`)
    } else {
      next = next.replace('</p:presentation>', () => `<p:extLst>${ext}</p:extLst></p:presentation>`)
    }
  }
  archive.entries.set(PRES_PATH, Buffer.from(next, 'utf8'))
}

/**
 * Positional normalization: each section's start = its minimum index (empty
 * sections take the next section's start), section i covers
 * [starts[i], starts[i+1]), the last section runs to total. Returns the
 * unsectioned leading slides + the normalized sections.
 */
export function normalizeSections(
  sections: SectionInfo[],
  total: number,
): { lead: number[]; starts: number[]; sections: SectionInfo[] } {
  if (!sections.length) return { lead: range(0, total), starts: [], sections: [] }
  const starts = new Array<number>(sections.length)
  let nextStart = total
  for (let i = sections.length - 1; i >= 0; i--) {
    const indices = sections[i]!.slideIndices
    let own = nextStart
    for (let k = 0; k < indices.length; k++) own = Math.min(own, indices[k]!)
    starts[i] = Math.min(own, nextStart)
    nextStart = starts[i]!
  }
  const norm = sections.map((s, i) => ({
    ...s,
    slideIndices: range(starts[i]!, i + 1 < sections.length ? starts[i + 1]! : total),
  }))
  return { lead: range(0, starts[0]!), starts, sections: norm }
}

/**
 * Add a section before slide atSlideIndex (covering atSlideIndex through the end
 * of its containing section). For documents without sections: slides before
 * atSlideIndex go into an auto-created "default section".
 * Returns the updated section list; null for an invalid index.
 */
export function addSection(
  opened: OpenedPptx,
  atSlideIndex: number,
  name: string,
): SectionInfo[] | null {
  const total = opened.deck.slides.length
  if (atSlideIndex < 0 || atSlideIndex >= total) return null
  const cur = getSections(opened)
  const newSec: SectionInfo = { id: newSectionId(), name, slideIndices: [] }
  let next: SectionInfo[]
  if (!cur.length) {
    next = []
    if (atSlideIndex > 0)
      next.push({
        id: newSectionId(),
        name: 'Default Section',
        slideIndices: range(0, atSlideIndex),
      })
    newSec.slideIndices = range(atSlideIndex, total)
    next.push(newSec)
  } else {
    const { starts, sections } = normalizeSections(cur, total)
    if (atSlideIndex < starts[0]!) {
      // Inserting within the unsectioned leading slides: new section covers [at, first section's start)
      newSec.slideIndices = range(atSlideIndex, starts[0]!)
      next = [newSec, ...sections]
    } else {
      // Find section g containing at, split into [g.start, at) + new section [at, g.end)
      let g = sections.length - 1
      for (let i = 0; i < sections.length; i++) {
        const end = i + 1 < sections.length ? starts[i + 1]! : total
        if (atSlideIndex >= starts[i]! && atSlideIndex < end) {
          g = i
          break
        }
      }
      const end = g + 1 < sections.length ? starts[g + 1]! : total
      newSec.slideIndices = range(atSlideIndex, end)
      const gSec = { ...sections[g]!, slideIndices: range(starts[g]!, atSlideIndex) }
      next = [...sections.slice(0, g), gSec, newSec, ...sections.slice(g + 1)]
    }
  }
  setSections(opened, next)
  return getSections(opened)
}

/** Rename a section; returns null if the id does not exist. */
export function renameSection(opened: OpenedPptx, id: string, name: string): SectionInfo[] | null {
  const cur = getSections(opened)
  const s = cur.find((x) => x.id === id)
  if (!s) return null
  s.name = name
  setSections(opened, cur)
  return getSections(opened)
}

/**
 * Delete a section but keep its slides: the section's slides merge into the
 * previous section (or the next one for the first section); when no sections
 * remain the whole sectionLst is removed. Returns null if the id does not exist.
 */
export function removeSection(
  opened: OpenedPptx,
  id: string,
  _opts?: { keepSlides?: boolean },
): SectionInfo[] | null {
  const total = opened.deck.slides.length
  const cur = getSections(opened)
  const i = cur.findIndex((x) => x.id === id)
  if (i < 0) return null
  const { sections } = normalizeSections(cur, total)
  const removed = sections.splice(i, 1)[0]!
  if (sections.length) {
    if (i > 0) {
      sections[i - 1]!.slideIndices = [...sections[i - 1]!.slideIndices, ...removed.slideIndices]
    } else {
      sections[0]!.slideIndices = [...removed.slideIndices, ...sections[0]!.slideIndices]
    }
  }
  setSections(opened, sections)
  return getSections(opened)
}

/**
 * Move slide fromIndex to toIndex (the insertion index after removal, i.e.
 * drag-drop semantics). Keeps presentation.xml's sldIdLst order and deck.slides
 * in sync; with sections, the moved slide joins the section at the drop position
 * (a section emptied by the move stays as an empty section).
 */
export function moveSlide(opened: OpenedPptx, fromIndex: number, toIndex: number): boolean {
  const { deck, archive } = opened
  const total = deck.slides.length
  if (fromIndex < 0 || fromIndex >= total) return false
  const to = Math.max(0, Math.min(toIndex, total - 1))
  if (fromIndex === to) return false

  // Before moving, label each slide with its section by position (-1 = unsectioned leading slide)
  const cur = getSections(opened)
  let labels: number[] | null = null
  if (cur.length) {
    const { sections } = normalizeSections(cur, total)
    labels = new Array<number>(total).fill(-1)
    sections.forEach((s, si) => {
      for (const i of s.slideIndices) labels![i] = si
    })
  }

  // presentation.xml sldIdLst: move the sldId tag wholesale (id/rId travel with it)
  const pres = archive.readText(PRES_PATH)
  if (!pres) return false
  const m = /<p:sldIdLst>([\s\S]*?)<\/p:sldIdLst>/.exec(pres)
  if (!m) return false
  const tags = [...m[1]!.matchAll(/<p:sldId\b[^>]*\/>/g)].map((t) => t[0])
  if (tags.length !== total) return false
  const [tag] = tags.splice(fromIndex, 1)
  tags.splice(to, 0, tag!)
  archive.entries.set(
    PRES_PATH,
    Buffer.from(
      pres.replace(m[0], () => `<p:sldIdLst>${tags.join('')}</p:sldIdLst>`),
      'utf8',
    ),
  )

  // Sync deck.slides
  const [slide] = deck.slides.splice(fromIndex, 1)
  deck.slides.splice(to, 0, slide!)

  // Section membership: the moved slide adopts the label of the slide it displaces at the drop point (keeping sections contiguous), written back in the new order
  if (labels) {
    labels.splice(fromIndex, 1)
    const adopt = to < labels.length ? labels[to]! : (labels[labels.length - 1] ?? -1)
    labels.splice(to, 0, adopt)
    const rebuilt = cur.map((s, si) => ({
      ...s,
      slideIndices: labels!.flatMap((l, pos) => (l === si ? [pos] : [])),
    }))
    setSections(opened, rebuilt)
  }
  return true
}

/**
 * Move a whole section up/down: swaps places with the adjacent section, keeping
 * presentation.xml's sldIdLst order and deck.slides order in sync (unsectioned
 * leading slides stay pinned at the front). Returns the updated section list.
 */
export function moveSection(
  opened: OpenedPptx,
  id: string,
  dir: 'up' | 'down',
): SectionInfo[] | null {
  const { deck, archive } = opened
  const total = deck.slides.length
  const cur = getSections(opened)
  const idx = cur.findIndex((s) => s.id === id)
  if (idx < 0) return null
  const j = dir === 'up' ? idx - 1 : idx + 1
  if (j < 0 || j >= cur.length) return null

  const { lead, sections } = normalizeSections(cur, total)
  const order = [...sections]
  ;[order[idx], order[j]] = [order[j]!, order[idx]!]
  const newOldOrder = [...lead, ...order.flatMap((s) => s.slideIndices)]
  if (newOldOrder.length !== total) return null

  // Sync presentation.xml sldIdLst order (sldId tags moved wholesale, id/rId travel with them)
  const pres = archive.readText(PRES_PATH)
  if (!pres) return null
  const m = /<p:sldIdLst>([\s\S]*?)<\/p:sldIdLst>/.exec(pres)
  if (!m) return null
  const tags = [...m[1]!.matchAll(/<p:sldId\b[^>]*\/>/g)].map((t) => t[0])
  if (tags.length !== total) return null
  const newInner = newOldOrder.map((i) => tags[i]!).join('')
  archive.entries.set(
    PRES_PATH,
    Buffer.from(
      pres.replace(m[0], () => `<p:sldIdLst>${newInner}</p:sldIdLst>`),
      'utf8',
    ),
  )

  // Sync deck.slides order
  deck.slides = newOldOrder.map((i) => deck.slides[i]!)

  // Recompute section indices for the new slide order and write back
  let pos = lead.length
  const rebuilt = order.map((s) => {
    const indices = range(pos, pos + s.slideIndices.length)
    pos += s.slideIndices.length
    return { ...s, slideIndices: indices }
  })
  setSections(opened, rebuilt)
  return getSections(opened)
}
