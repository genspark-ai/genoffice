import { Extension } from '@tiptap/core'
import type { Editor } from '@tiptap/core'
import createHyphenator from 'hyphen'
import type { Node as ProseMirrorNode } from '@tiptap/pm/model'
import { Plugin, PluginKey, type EditorState, type Transaction } from '@tiptap/pm/state'
import { Decoration, DecorationSet, type EditorView } from '@tiptap/pm/view'
import { rangeSlot } from '../dom-range'
import { PHASED_CONTENT_SETTLED_EVENT, isPhasedContentPending } from '../phased-content'
import {
  DOC_CSS_COMMITTED_EVENT,
  WRAP_WIDTHS_COMMITTED_EVENT,
  cjkPunctShrinkPluginKey,
} from './cjk-punct-shrink'
import { justifyShrinkPluginKey, sameLine } from './justify-shrink'

/**
 * Automatic hyphenation with our own break points. Chromium's `hyphens: auto`
 * leaves words with a capital letter whole and offers no control over how many
 * lines in a row end in a hyphen or how much ragged edge is tolerated, so the
 * break points come from TeX hyphenation patterns instead: every candidate
 * point of a body paragraph carries a soft hyphen as a widget decoration (the
 * document text is untouched) and Chromium's own line breaker picks among
 * them. The consecutive limit and the zone are enforced after layout by
 * switching off the candidates of the offending word.
 */

type PatternModule = { default: unknown }

const PATTERNS: Record<string, () => Promise<PatternModule>> = {
  'en-us': () => import('hyphen/patterns/en-us'),
  'en-gb': () => import('hyphen/patterns/en-gb'),
  de: () => import('hyphen/patterns/de-1996'),
  fr: () => import('hyphen/patterns/fr'),
  es: () => import('hyphen/patterns/es'),
  pt: () => import('hyphen/patterns/pt'),
  it: () => import('hyphen/patterns/it'),
  nl: () => import('hyphen/patterns/nl'),
}

const BRITISH = new Set(['gb', 'au', 'nz', 'ie', 'za', 'in'])

/** the pattern set for a w:lang value; a document that names no language is read as US English */
export function hyphenationDictFor(lang: string | null | undefined): string | null {
  if (!lang) return 'en-us'
  const [primary, region] = lang.toLowerCase().split(/[-_]/)
  if (primary === 'en') return region && BRITISH.has(region) ? 'en-gb' : 'en-us'
  return primary in PATTERNS ? primary : null
}

const SHY = '\u00ad'
const hyphenators = new Map<string, (text: string) => string>()
const loading = new Map<string, Promise<boolean>>()
const pointCache = new Map<string, Map<string, readonly number[]>>()
const loadListeners = new Set<() => void>()

export function hyphenationReady(dict: string): boolean {
  return hyphenators.has(dict)
}

export function loadHyphenation(dict: string): Promise<boolean> {
  if (hyphenators.has(dict)) return Promise.resolve(true)
  const loader = PATTERNS[dict]
  if (!loader) return Promise.resolve(false)
  let pending = loading.get(dict)
  if (!pending) {
    pending = loader().then(
      (mod) => {
        hyphenators.set(dict, createHyphenator(mod.default, { minWordLength: MIN_WORD }))
        for (const fn of loadListeners) fn()
        if (docConfig.dict === dict) for (const fn of configListeners) fn()
        return true
      },
      () => false,
    )
    loading.set(dict, pending)
  }
  return pending
}

const MIN_WORD = 5

/** offsets inside `word` where it may break; empty until the dictionary has loaded */
export function breakPoints(dict: string, word: string): readonly number[] {
  const hyphenate = hyphenators.get(dict)
  if (!hyphenate) return []
  let cache = pointCache.get(dict)
  if (!cache) pointCache.set(dict, (cache = new Map()))
  let points = cache.get(word)
  if (!points) {
    const out: number[] = []
    let offset = 0
    for (const ch of hyphenate(word)) {
      if (ch === SHY) out.push(offset)
      else offset += ch.length
    }
    points = out
    if (cache.size > 50000) cache.clear()
    cache.set(word, points)
  }
  return points
}

const TOKEN_RE = /\S+/g
const LETTERS_RE = /\p{L}[\p{L}\p{M}]*/gu
// an optional hyphen already says where the word breaks; addresses and codes stay whole
const WHOLE_TOKEN_RE = /[\u00ad\d@/\\_]|^www\./i

const isAllCaps = (word: string) => word === word.toUpperCase() && word !== word.toLowerCase()

export interface HyphenWord {
  /** offsets in the text handed in */
  from: number
  to: number
  text: string
  points: readonly number[]
}

/** the words of a stretch of text that have break points */
export function hyphenWords(text: string, dict: string, caps: boolean): HyphenWord[] {
  const out: HyphenWord[] = []
  for (const token of text.matchAll(TOKEN_RE)) {
    if (WHOLE_TOKEN_RE.test(token[0])) continue
    for (const m of token[0].matchAll(LETTERS_RE)) {
      const word = m[0]
      if (word.length < MIN_WORD || (!caps && isAllCaps(word))) continue
      const points = breakPoints(dict, word)
      if (points.length === 0) continue
      const from = token.index + m.index
      out.push({ from, to: from + word.length, text: word, points })
    }
  }
  return out
}

export interface HyphenationConfig {
  /** pattern set the document hyphenates with; null = our hyphenation is off */
  dict: string | null
  /** Hyphenate words in CAPS */
  caps: boolean
  /** lines in a row that may end in a hyphen; 0 = no limit */
  limit: number
  /** Word's hyphenation zone in CSS px; 0 = none */
  zonePx: number
}

export type HyphenationStorage = HyphenationConfig

/** the open document's rules for text outside the body editor (text boxes, strips, notes) */
const docConfig: HyphenationConfig = { dict: null, caps: true, limit: 0, zonePx: 0 }
const followers = new Set<EditorView>()
const configListeners = new Set<() => void>()

export const docHyphenation = (): Readonly<HyphenationConfig> => docConfig

/** what was cut under another value is stale; empty when those places do not hyphenate */
export function docHyphenationSig(): string {
  const { dict, caps, limit, zonePx } = docConfig
  if (!dict) return ''
  return `${dict}:${hyphenationReady(dict) ? 1 : 0}:${caps ? 1 : 0}:${limit}:${zonePx}`
}

export function onDocHyphenationChange(fn: () => void): () => void {
  configListeners.add(fn)
  return () => configListeners.delete(fn)
}

function publishDocHyphenation(): void {
  for (const view of followers) {
    if (view.isDestroyed) continue
    view.dispatch(
      view.state.tr
        .setMeta(hyphenationPluginKey, { reset: true, relayout: true } satisfies HyphenationMeta)
        .setMeta('addToHistory', false),
    )
  }
  for (const fn of configListeners) fn()
}

declare module '@tiptap/core' {
  interface Storage {
    hyphenation: HyphenationStorage
  }
}

interface HyphenationState {
  decos: DecorationSet
  /** word starts answered No in a manual run: they keep no candidates */
  skip: readonly number[]
}

type HyphenationMeta = { reset?: true; skip?: readonly number[]; relayout?: true }

export const hyphenationPluginKey = new PluginKey<HyphenationState>('hyphenation')

const runDict = (node: ProseMirrorNode, docDict: string): string | null => {
  for (const mark of node.marks) {
    const raw = mark.attrs.lang as string | null | undefined
    if (!raw) continue
    try {
      const val = (JSON.parse(raw) as { val?: string }).val
      if (val) return hyphenationDictFor(val)
    } catch {
      return docDict
    }
  }
  return docDict
}

/** words with break points inside one textblock; offsets are relative to its content */
export function blockHyphenWords(block: ProseMirrorNode, config: HyphenationConfig): HyphenWord[] {
  const docDict = config.dict
  if (!docDict) return []
  const out: HyphenWord[] = []
  let start = 0
  let text = ''
  let dict: string | null = docDict
  const flush = () => {
    if (text && dict) {
      if (hyphenationReady(dict)) {
        for (const w of hyphenWords(text, dict, config.caps))
          out.push({ ...w, from: w.from + start, to: w.to + start })
      } else void loadHyphenation(dict)
    }
    text = ''
  }
  block.forEach((child, offset) => {
    const childDict = child.isText ? runDict(child, docDict) : null
    if (!child.isText || !child.text || (text && childDict !== dict)) flush()
    if (!child.isText || !child.text) return
    if (!text) {
      start = offset
      dict = childDict
    }
    text += child.text
  })
  flush()
  return out
}

const shyWidget = () => {
  const span = document.createElement('span')
  span.className = 'doc-shy'
  return span
}
const WIDGET_SPEC = { side: -1, ignoreSelection: true }

function blockDecorations(
  block: ProseMirrorNode,
  pos: number,
  config: HyphenationConfig,
  skip: readonly number[],
): Decoration[] {
  const out: Decoration[] = [Decoration.node(pos, pos + block.nodeSize, { class: 'doc-hyph' })]
  for (const word of blockHyphenWords(block, config)) {
    const from = pos + 1 + word.from
    if (skip.includes(from)) continue
    for (const p of word.points) out.push(Decoration.widget(from + p, shyWidget, WIDGET_SPEC))
  }
  return out
}

const hyphenates = (node: ProseMirrorNode) => node.isTextblock && !node.type.spec.code

function buildAll(
  doc: ProseMirrorNode,
  config: HyphenationConfig,
  skip: readonly number[],
): DecorationSet {
  if (!config.dict) return DecorationSet.empty
  const decos: Decoration[] = []
  doc.descendants((node, pos) => {
    if (!node.isTextblock) return true
    if (hyphenates(node)) decos.push(...blockDecorations(node, pos, config, skip))
    return false
  })
  return DecorationSet.create(doc, decos)
}

function rebuildChanged(
  old: DecorationSet,
  tr: Transaction,
  config: HyphenationConfig,
  skip: readonly number[],
): DecorationSet {
  const blocks = new Map<number, ProseMirrorNode>()
  tr.mapping.maps.forEach((map, i) => {
    const rest = tr.mapping.slice(i + 1)
    map.forEach((_oldFrom, _oldTo, newFrom, newTo) => {
      const from = Math.max(0, rest.map(newFrom, -1) - 1)
      const to = Math.min(tr.doc.content.size, rest.map(newTo, 1) + 1)
      tr.doc.nodesBetween(from, to, (node, pos) => {
        if (!node.isTextblock) return true
        if (hyphenates(node)) blocks.set(pos, node)
        return false
      })
    })
  })
  let set = old.map(tr.mapping, tr.doc)
  const fresh: Decoration[] = []
  for (const [pos, node] of blocks) {
    const end = pos + node.nodeSize
    // find() also returns the neighbours' decorations that only touch the block
    const stale = set.find(pos, end).filter((d) => d.from >= pos && d.to <= end)
    if (stale.length) set = set.remove(stale)
    fresh.push(...blockDecorations(node, pos, config, skip))
  }
  return fresh.length ? set.add(tr.doc, fresh) : set
}

const LETTER_RE = /[\p{L}\p{M}]/u
const ONLY_LETTERS_RE = /^[\p{L}\p{M}]*$/u
const probeRange = rangeSlot()
const NOTIFY_MAX = 6

/** box of the first character after `span`: it opens the next line when the word broke there */
function tailRect(span: Element, root: Element): DOMRect | null {
  const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT)
  walker.currentNode = span
  for (let t = walker.nextNode() as Text | null; t; t = walker.nextNode() as Text | null) {
    if (!t.data) continue
    const range = probeRange()
    range.setStart(t, 0)
    range.setEnd(t, 1)
    return range.getBoundingClientRect()
  }
  return null
}

/** left edge of the word `span` sits in, where that word starts on the line */
function wordLeft(span: Element, root: Element): number | null {
  const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT)
  walker.currentNode = span
  let first: { node: Text; offset: number } | null = null
  scan: for (
    let t = walker.previousNode() as Text | null;
    t;
    t = walker.previousNode() as Text | null
  ) {
    for (let i = t.data.length - 1; i >= 0; i--) {
      if (!LETTER_RE.test(t.data[i])) break scan
      first = { node: t, offset: i }
    }
  }
  if (!first) return null
  const range = probeRange()
  range.setStart(first.node, first.offset)
  range.setEnd(first.node, first.offset + 1)
  return range.getBoundingClientRect().left
}

/** the candidates sharing a word with spans[index] */
function wordSpans(spans: Element[], index: number): Element[] {
  const joined = (a: Element, b: Element) => {
    const range = probeRange()
    range.setStartAfter(a)
    range.setEndBefore(b)
    return ONLY_LETTERS_RE.test(range.toString())
  }
  let lo = index
  while (lo > 0 && joined(spans[lo - 1], spans[lo])) lo--
  let hi = index
  while (hi < spans.length - 1 && joined(spans[hi], spans[hi + 1])) hi++
  return spans.slice(lo, hi + 1)
}

/**
 * Switches off the candidates of words that may not break under the consecutive
 * limit or the zone, top line first: taking a word off only moves the lines below it.
 * Returns whether the set of switched-off candidates changed.
 */
export function enforceLineRules(el: HTMLElement, limit: number, zonePx: number): boolean {
  const spans = Array.from(el.querySelectorAll('.doc-shy'))
  const before = spans.filter((s) => s.classList.contains('doc-shy-off'))
  for (const s of before) s.classList.remove('doc-shy-off')
  const off: Element[] = []
  const box = el.getBoundingClientRect()
  if ((limit || zonePx) && spans.length > 0 && box.width > 0 && el.offsetWidth > 0) {
    const zoom = box.width / el.offsetWidth
    const cs = getComputedStyle(el)
    const lineRight =
      box.right -
      ((parseFloat(cs.paddingRight) || 0) + (parseFloat(cs.borderRightWidth) || 0)) * zoom
    let run = 0
    let prevTail: DOMRect | null = null
    for (let i = 0; i < spans.length; i++) {
      const rect = spans[i].getBoundingClientRect()
      if (rect.width <= 0.01) continue
      const nextRun = prevTail && sameLine(rect, prevTail) ? run + 1 : 1
      let whole = limit > 0 && nextRun > limit
      if (!whole && zonePx) {
        const left = wordLeft(spans[i], el)
        // Word leaves the word whole when the gap it would close is no wider than the zone
        whole = left !== null && lineRight - left <= zonePx * zoom
      }
      if (whole) {
        // the line now ends before this word: the hyphen count of the lines above stands
        const group = wordSpans(spans, i)
        for (const s of group) s.classList.add('doc-shy-off')
        off.push(...group)
        i = spans.indexOf(group[group.length - 1])
        continue
      }
      run = nextRun
      prevTail = tailRect(spans[i], el)
    }
  }
  return off.length !== before.length || off.some((s, i) => s !== before[i])
}

/** where the paragraph breaks at our candidates and which of them are off: the same again = nothing moved */
function layoutSignature(el: HTMLElement): string {
  const parts: string[] = [String(el.offsetWidth)]
  el.querySelectorAll('.doc-shy').forEach((s, i) => {
    if (s.classList.contains('doc-shy-off')) parts.push(`o${i}`)
    else if (s.getBoundingClientRect().width > 0.01) parts.push(`b${i}`)
  })
  return parts.join(',')
}

class HyphenationView {
  private notified = 0
  private resizeObserver?: ResizeObserver
  private lastDomWidth = -1
  /** paragraphs whose line ends were checked, with the layout they were checked in */
  private settled = new WeakMap<ProseMirrorNode, string>()
  private recent: ProseMirrorNode[] = []
  private onLayoutInput = () => this.pass('all')
  private lastBox = ''
  // every text box editor hears a document-wide layout input: it re-checks only when its own box changed
  private onDocLayout = () => {
    const { view, storage } = this
    if (storage === docConfig) {
      // a pass that cannot run yet must not count as this box's check
      if (!storage.dict || !view.dom.isConnected || isPhasedContentPending()) return
      const box = `${view.dom.offsetWidth}:${view.dom.offsetHeight}`
      if (box === this.lastBox) return
      this.lastBox = box
    }
    this.pass('all')
  }
  private onLoaded = () => {
    const { view } = this
    if (!this.storage.dict || view.isDestroyed) return
    view.dispatch(
      view.state.tr
        .setMeta(hyphenationPluginKey, { reset: true, relayout: true } satisfies HyphenationMeta)
        .setMeta('addToHistory', false),
    )
  }

  constructor(
    private view: EditorView,
    private storage: HyphenationStorage,
  ) {
    if (storage === docConfig) followers.add(view)
    loadListeners.add(this.onLoaded)
    document.fonts?.addEventListener('loadingdone', this.onLayoutInput)
    document.addEventListener(DOC_CSS_COMMITTED_EVENT, this.onDocLayout)
    document.addEventListener(WRAP_WIDTHS_COMMITTED_EVENT, this.onDocLayout)
    document.addEventListener(PHASED_CONTENT_SETTLED_EVENT, this.onDocLayout)
    if (typeof ResizeObserver !== 'undefined') {
      this.resizeObserver = new ResizeObserver(() => {
        const w = this.view.dom.offsetWidth
        if (w === this.lastDomWidth) return
        this.lastDomWidth = w
        this.pass('all')
      })
      this.resizeObserver.observe(view.dom)
    }
    this.pass('all')
  }

  update(view: EditorView, prev: EditorState) {
    const changed = (key: PluginKey) => key.getState(view.state) !== key.getState(prev)
    if (changed(hyphenationPluginKey) && view.state.doc === prev.doc) {
      // other rules: what was switched off under the old ones is checked again
      this.settled = new WeakMap()
      this.pass('all')
    } else if (view.state.doc !== prev.doc) {
      this.notified = 0
      this.pass('new')
    }
    // their spacing moves the line ends of the paragraphs just edited
    else if (changed(justifyShrinkPluginKey) || changed(cjkPunctShrinkPluginKey))
      this.pass('recent')
  }

  destroy() {
    followers.delete(this.view)
    loadListeners.delete(this.onLoaded)
    document.fonts?.removeEventListener('loadingdone', this.onLayoutInput)
    document.removeEventListener(DOC_CSS_COMMITTED_EVENT, this.onDocLayout)
    document.removeEventListener(WRAP_WIDTHS_COMMITTED_EVENT, this.onDocLayout)
    document.removeEventListener(PHASED_CONTENT_SETTLED_EVENT, this.onDocLayout)
    this.resizeObserver?.disconnect()
  }

  /**
   * 'new': the paragraphs the document change touched; 'recent': those again;
   * 'all': every paragraph whose line ends are not where they were checked.
   */
  private pass(scope: 'new' | 'recent' | 'all') {
    const { view, storage } = this
    if (!storage.dict || !view.dom.isConnected || isPhasedContentPending()) return
    // PDF export parks the editor subtree: a layout read would re-lay it out per print chunk
    if (view.dom.closest('.app.pv-exporting')) return
    const active = storage.limit > 0 || storage.zonePx > 0
    if (!active && view.dom.querySelector('.doc-shy-off') === null) return
    const todo: Array<{ node: ProseMirrorNode; el: HTMLElement }> = []
    const recent = new Set(this.recent)
    view.state.doc.descendants((node, pos) => {
      if (!node.isTextblock) return true
      const known = this.settled.get(node)
      if (scope === 'new' ? known !== undefined : scope === 'recent' && !recent.has(node)) {
        return false
      }
      const el = view.nodeDOM(pos)
      if (!(el instanceof HTMLElement) || !el.classList.contains('doc-hyph')) return false
      if (scope === 'all' && known !== undefined && known === layoutSignature(el)) return false
      todo.push({ node, el })
      return false
    })
    if (scope === 'new') this.recent = todo.map((t) => t.node)
    let moved = false
    for (const { node, el } of todo) {
      if (enforceLineRules(el, storage.limit, storage.zonePx)) moved = true
      this.settled.set(node, layoutSignature(el))
    }
    // a text box keeps its size and a strip is measured again on commit: nothing follows their lines
    if (!moved || storage === docConfig || this.notified >= NOTIFY_MAX) return
    this.notified++
    // lines moved with no change to the document: the measuring views and the pagination follow
    view.dispatch(
      view.state.tr
        .setMeta(hyphenationPluginKey, { relayout: true } satisfies HyphenationMeta)
        .setMeta('addToHistory', false),
    )
    document.dispatchEvent(new Event(DOC_CSS_COMMITTED_EVENT))
  }
}

export const HyphenationExtension = Extension.create<{ followDocument: boolean }>({
  name: 'hyphenation',
  addOptions() {
    return { followDocument: false }
  },
  addStorage(): HyphenationStorage {
    return { dict: null, caps: true, limit: 0, zonePx: 0 }
  },
  addProseMirrorPlugins() {
    // the editor of a text box or a strip reads the document's rules, not its own
    const storage = this.options.followDocument ? docConfig : (this.storage as HyphenationStorage)
    return [
      new Plugin<HyphenationState>({
        key: hyphenationPluginKey,
        state: {
          init: (_config, state) => ({ decos: buildAll(state.doc, storage, []), skip: [] }),
          apply(tr, old) {
            const meta = tr.getMeta(hyphenationPluginKey) as HyphenationMeta | undefined
            const skip =
              meta?.skip ?? (tr.docChanged ? old.skip.map((p) => tr.mapping.map(p)) : old.skip)
            if (meta?.reset || meta?.skip) return { decos: buildAll(tr.doc, storage, skip), skip }
            if (!tr.docChanged) return old
            if (!storage.dict) return { decos: DecorationSet.empty, skip }
            return { decos: rebuildChanged(old.decos, tr, storage, skip), skip }
          },
        },
        props: {
          decorations(state) {
            return this.getState(state)?.decos
          },
        },
        view: (editorView) => new HyphenationView(editorView, storage),
      }),
    ]
  },
})

/** true when a transaction of this plugin moved line breaks without changing the document */
export function hyphenationRelayout(tr: Transaction): boolean {
  return (tr.getMeta(hyphenationPluginKey) as HyphenationMeta | undefined)?.relayout === true
}

const sameConfig = (a: HyphenationConfig, b: HyphenationConfig) =>
  a.dict === b.dict && a.caps === b.caps && a.limit === b.limit && a.zonePx === b.zonePx

export function setDocHyphenation(config: HyphenationConfig): void {
  if (sameConfig(docConfig, config)) return
  Object.assign(docConfig, config)
  publishDocHyphenation()
}

/**
 * Hands the document's hyphenation rules to the editor; the page follows at once.
 * `elsewhere` goes to text boxes, strips and notes: a manual run probes the body alone.
 */
export function setHyphenation(
  editor: Editor,
  config: HyphenationConfig,
  elsewhere: HyphenationConfig = config,
): void {
  const storage = editor.storage.hyphenation
  setDocHyphenation(elsewhere)
  // a manual run's answers go with the rules they were given under
  if (!hyphenationPluginKey.getState(editor.state)?.skip.length && sameConfig(storage, config))
    return
  Object.assign(storage, config)
  const { view } = editor
  view.dispatch(
    view.state.tr
      .setMeta(hyphenationPluginKey, {
        reset: true,
        skip: [],
        relayout: true,
      } satisfies HyphenationMeta)
      .setMeta('addToHistory', false),
  )
}

export interface ManualCandidate {
  /** document positions of the word */
  from: number
  to: number
  text: string
  points: readonly number[]
  /** the point the line breaks at as laid out */
  chosen: number
  el: Element
}

const isBroken = (span: Element) =>
  !span.classList.contains('doc-shy-off') && span.getBoundingClientRect().width > 0.01

/** the first word after `after` that ends a line at one of its candidates */
export function nextManualCandidate(editor: Editor, after: number): ManualCandidate | null {
  const { view } = editor
  const config = editor.storage.hyphenation
  for (const span of Array.from(view.dom.querySelectorAll('.doc-hyph .doc-shy'))) {
    if (!isBroken(span)) continue
    let pos: number
    try {
      pos = view.posAtDOM(span, 0)
    } catch {
      continue
    }
    if (pos <= after) continue
    const $pos = view.state.doc.resolve(pos)
    const block = $pos.parent
    if (!block.isTextblock) continue
    const offset = $pos.parentOffset
    const word = blockHyphenWords(block, config).find((w) => w.from < offset && offset < w.to)
    if (!word) continue
    const start = $pos.start()
    return {
      from: start + word.from,
      to: start + word.to,
      text: word.text,
      points: word.points,
      chosen: offset - word.from,
      el: span,
    }
  }
  return null
}

/** a manual run's No: the word keeps no candidates for the rest of the run */
export function skipManualWord(editor: Editor, from: number): void {
  const { view } = editor
  const skip = [...(hyphenationPluginKey.getState(view.state)?.skip ?? []), from]
  view.dispatch(
    view.state.tr
      .setMeta(hyphenationPluginKey, { skip } satisfies HyphenationMeta)
      .setMeta('addToHistory', false),
  )
}

/** the finer rules of Tools > Hyphenation as the document carries them */
export interface HyphenationSettings {
  caps: boolean
  /** 0 = no limit */
  limit: number
  zoneTwips: number
}

interface HyphenationDocFacts {
  autoHyphenation?: boolean
  doNotHyphenateCaps?: boolean
  consecutiveHyphenLimit?: number
  hyphenationZoneTwips?: number
  compatibilityMode?: number
  docDefaults?: { lang?: string }
}

export const DEFAULT_HYPHENATION_ZONE_TWIPS = 360

export function hyphenationSettingsOf(
  parsed: HyphenationDocFacts | null | undefined,
  edit?: Partial<HyphenationSettings> | null,
): HyphenationSettings {
  return {
    caps: edit?.caps ?? parsed?.doNotHyphenateCaps !== true,
    limit: edit?.limit ?? parsed?.consecutiveHyphenLimit ?? 0,
    zoneTwips: edit?.zoneTwips ?? parsed?.hyphenationZoneTwips ?? DEFAULT_HYPHENATION_ZONE_TWIPS,
  }
}

/** Word 2013+ layout breaks lines without a hyphenation zone; earlier modes keep it */
export const hyphenationZoneApplies = (parsed: HyphenationDocFacts | null | undefined): boolean =>
  (parsed?.compatibilityMode ?? 0) < 15

export function hyphenationConfigOf(
  parsed: HyphenationDocFacts,
  on: boolean,
  settings: HyphenationSettings = hyphenationSettingsOf(parsed),
): HyphenationConfig {
  return {
    dict: on ? hyphenationDictFor(parsed.docDefaults?.lang) : null,
    caps: settings.caps,
    limit: settings.limit,
    zonePx: hyphenationZoneApplies(parsed) ? settings.zoneTwips / 15 : 0,
  }
}
