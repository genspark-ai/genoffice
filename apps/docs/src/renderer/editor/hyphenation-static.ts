import { DOC_CSS_COMMITTED_EVENT } from './cjk-punct-shrink'
import {
  docHyphenation,
  enforceLineRules,
  hyphenWords,
  hyphenationReady,
  loadHyphenation,
  onDocHyphenationChange,
} from './hyphenation'

/**
 * The document's break points in text drawn without an editor (header and
 * footer strips, notes, display-only text boxes): the same `.doc-shy` spans
 * the body carries, so one stylesheet and one line-rule pass serve every place.
 */

function breakOffsets(text: string): number[] {
  const { dict, caps } = docHyphenation()
  if (!dict || !text) return []
  if (!hyphenationReady(dict)) {
    void loadHyphenation(dict)
    return []
  }
  const out: number[] = []
  for (const word of hyphenWords(text, dict, caps))
    for (const p of word.points) out.push(word.from + p)
  return out
}

/**
 * The texts of one paragraph's runs, each cut where a word may break (a word
 * across runs breaks as one word; a break at a run's start leaves an empty
 * first piece). null when nothing may break.
 */
export function hyphenPieces(texts: readonly string[]): string[][] | null {
  const offsets = breakOffsets(texts.join(''))
  if (offsets.length === 0) return null
  let start = 0
  let next = 0
  return texts.map((text) => {
    const pieces: string[] = []
    let from = 0
    // a point on the seam of two runs goes to the run that follows
    while (next < offsets.length && offsets[next] < start + text.length) {
      const cut = offsets[next++] - start
      pieces.push(text.slice(from, cut))
      from = cut
    }
    pieces.push(text.slice(from))
    start += text.length
    return pieces
  })
}

const SHY_CLASS = 'doc-shy'

/** paragraphs under `root` holding break points of their own; one hosting a hanging box counts beside the box's */
function ownBlocks(root: HTMLElement): HTMLElement[] {
  const out = new Set<HTMLElement>()
  for (const span of root.querySelectorAll(`.${SHY_CLASS}`)) {
    const block = span.closest<HTMLElement>(BLOCKS)
    if (block && (block === root || root.contains(block)) && !block.closest(SUB_EDITORS))
      out.add(block)
  }
  // the host first: the box's own pass then has the last word on its points
  const all = [...out]
  const hosts = all.filter((el) => all.some((other) => other !== el && el.contains(other)))
  return [...hosts, ...all.filter((el) => !hosts.includes(el))]
}

export function appendPieces(host: Node, pieces: readonly string[]): void {
  pieces.forEach((piece, i) => {
    if (i > 0) {
      const shy = document.createElement('span')
      shy.className = SHY_CLASS
      host.appendChild(shy)
    }
    if (piece) host.appendChild(document.createTextNode(piece))
  })
}

/** class of a container whose text carries our break points: Chromium's dictionary stays out */
export const STATIC_HYPH_CLASS = 'doc-hyph-in'

export const staticHyphenationOn = (): boolean => docHyphenation().dict !== null

const BLOCKS =
  '.page-hf-para, .page-hf-cell-para, .doc-textbox-para, .page-gap-note, .pv-footnote, .page-note'
/** the editors of a text box and of a strip place and check their own break points */
const SUB_EDITORS = '.doc-textbox-editor, .page-hf-editor'

/** innermost only */
function staticBlocks(root: HTMLElement): HTMLElement[] {
  const found = Array.from(root.querySelectorAll<HTMLElement>(BLOCKS))
  if (root.matches(BLOCKS)) found.unshift(root)
  return found.filter((el) => !el.querySelector(BLOCKS) && !el.closest(SUB_EDITORS))
}

/** cuts markup built elsewhere (a display-only text box) in place */
export function hyphenateDom(root: HTMLElement): void {
  if (!staticHyphenationOn()) return
  root.classList.add(STATIC_HYPH_CLASS)
  for (const block of staticBlocks(root)) {
    if (block.querySelector(`.${SHY_CLASS}`)) continue
    const nodes: Text[] = []
    const walker = document.createTreeWalker(block, NodeFilter.SHOW_TEXT)
    for (let t = walker.nextNode() as Text | null; t; t = walker.nextNode() as Text | null)
      if (t.data) nodes.push(t)
    // text nodes of one inline flow form a group; a tab segment or a pad starts another
    let group: Text[] = []
    const flush = () => {
      const pieces = hyphenPieces(group.map((t) => t.data))
      pieces?.forEach((cut, i) => {
        if (cut.length === 1) return
        const frag = document.createDocumentFragment()
        appendPieces(frag, cut)
        group[i].replaceWith(frag)
      })
      group = []
    }
    for (const node of nodes) {
      const prev = group[group.length - 1]
      if (prev && !flowsOn(prev, node)) flush()
      group.push(node)
    }
    flush()
  }
}

/** `b` continues the inline text of `a`: same parent, or neighbouring run spans */
function flowsOn(a: Text, b: Text): boolean {
  if (a.parentNode === b.parentNode) return a.nextSibling === b
  const pa = a.parentElement
  const pb = b.parentElement
  return (
    !!pa &&
    !!pb &&
    pa.parentNode === pb.parentNode &&
    pa.nextSibling === pb &&
    a.nextSibling === null &&
    b.previousSibling === null
  )
}

export function unhyphenateDom(root: HTMLElement): void {
  root.classList.remove(STATIC_HYPH_CLASS)
  const spans = root.querySelectorAll(`.${SHY_CLASS}`)
  if (spans.length === 0) return
  const parents = new Set<Node>()
  for (const span of spans) {
    if (span.closest(SUB_EDITORS)) continue
    if (span.parentNode) parents.add(span.parentNode)
    span.remove()
  }
  for (const parent of parents) parent.normalize()
}

const settled = new WeakMap<Element, string>()

const signature = (el: HTMLElement, limit: number, zonePx: number): string => {
  const parts: string[] = [`${limit}:${zonePx}:${el.offsetWidth}`]
  el.querySelectorAll(`.${SHY_CLASS}`).forEach((s, i) => {
    if (s.classList.contains('doc-shy-off')) parts.push(`o${i}`)
    else if (s.getBoundingClientRect().width > 0.01) parts.push(`b${i}`)
  })
  return parts.join(',')
}

/** the consecutive limit and the zone on the static paragraphs under `root`, which must be laid out */
export function enforceStaticLineRules(root: HTMLElement): void {
  const { dict, limit, zonePx } = docHyphenation()
  if (!dict) return
  const active = limit > 0 || zonePx > 0
  if (!active && root.querySelector('.doc-shy-off') === null) return
  for (const block of ownBlocks(root)) {
    if (settled.get(block) === signature(block, limit, zonePx)) continue
    enforceLineRules(block, limit, zonePx)
    settled.set(block, signature(block, limit, zonePx))
  }
}

let sweepFrame: number | undefined

/** builders call it for markup they are about to attach */
export function scheduleStaticLineRules(): void {
  if (!staticHyphenationOn()) return
  if (sweepFrame !== undefined || typeof requestAnimationFrame === 'undefined') return
  sweepFrame = requestAnimationFrame(() => {
    sweepFrame = undefined
    // PDF export lays its pages out itself (PaginationPreview)
    if (document.querySelector('.app.pv-exporting')) return
    for (const root of document.querySelectorAll<HTMLElement>(`.${STATIC_HYPH_CLASS}`))
      enforceStaticLineRules(root)
  })
}

/** cuts now and again when the rules change; returns the cleanup */
export function followStaticHyphenation(root: HTMLElement): () => void {
  const apply = () => {
    unhyphenateDom(root)
    hyphenateDom(root)
    scheduleStaticLineRules()
  }
  if (staticHyphenationOn()) apply()
  return onDocHyphenationChange(apply)
}

if (typeof document !== 'undefined') {
  document.addEventListener(DOC_CSS_COMMITTED_EVENT, scheduleStaticLineRules)
  document.fonts?.addEventListener('loadingdone', scheduleStaticLineRules)
}
