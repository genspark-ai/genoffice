/**
 * Cross-references to a numbered paragraph (REF \r, \n, \w): the number the list shows
 * for the bookmarked paragraph, kept in step with the numbering as the document changes.
 */
import { Extension } from '@tiptap/core'
import type { Node as PmNode } from '@tiptap/pm/model'
import { Plugin, PluginKey, type Transaction } from '@tiptap/pm/state'
import type { Transform } from '@tiptap/pm/transform'
import {
  computeListNumberSites,
  paragraphNumberText,
  type ListItemRef,
  type ListNumberSite,
  type NumberingDef,
  type ParagraphNumberContext,
} from '@genoffice/docx-engine'
import { t } from '../i18n/locale'
import { collectFields, isFieldLocked, type FieldRange } from './field-codes'
import { writeField } from './field-results'
import { TEXT_FORMAT_SWITCHES, formatFieldText, formatSwitchOf } from './field-switches'
import { TRACK_IGNORE } from './revisions'

export interface NumberRefSwitches {
  context: ParagraphNumberContext
  /** REF \d */
  separator?: string
  /** REF \p: where the paragraph is from here goes after the number */
  aboveBelow: boolean
  /** the instruction, when its `\* Upper`, `Lower`, `FirstCap` or `Caps` shapes the result */
  cased?: string
}

export interface NumberRefWords {
  above: string
  below: string
}

/** the switches of a REF that shows a paragraph number; null for any other field, or one whose format is Word's to apply */
export function numberRefSwitches(instr: string | null | undefined): NumberRefSwitches | null {
  if (!instr || !/^\s*REF\s/i.test(instr)) return null
  const bare = instr.replace(/"[^"]*"/g, '""')
  const context: ParagraphNumberContext | null = /\\w\b/i.test(bare)
    ? 'full'
    : /\\r\b/i.test(bare)
      ? 'relative'
      : /\\n\b/i.test(bare)
        ? 'none'
        : null
  if (!context) return null
  const format = formatSwitchOf(bare)
  const cased =
    !!format && TEXT_FORMAT_SWITCHES.some((s) => s.toLowerCase() === format.toLowerCase())
  if (/\\[#@]/.test(bare) || /\\\*\s*DBCHAR\b/i.test(bare) || (format && !cased)) return null
  const separator = /\\d\s+(?:"([^"]*)"|([^\s\\]+))/i.exec(instr)
  return {
    context,
    ...(separator ? { separator: separator[1] ?? separator[2] } : {}),
    aboveBelow: /\\p\b/i.test(bare),
    ...(cased ? { cased: bare } : {}),
  }
}

export interface NumberedItem {
  pos: number
  node: PmNode
  site: ListNumberSite | null
}

/** the list items of the document in the order the numbering counts them */
export function numberedItems(doc: PmNode, defs: Map<string, NumberingDef>): NumberedItem[] {
  const refs: ListItemRef[] = []
  const nodes: Array<{ pos: number; node: PmNode }> = []
  doc.descendants((node, pos) => {
    if (node.type.name === 'docListItem') {
      const deleted =
        !!node.attrs.paraMarkDel ||
        (node.attrs.blockRevision as { kind?: string } | null)?.kind === 'del'
      refs.push({
        numId: (node.attrs.numId as string | null) ?? null,
        ilvl: Number(node.attrs.ilvl) || 0,
        ...(deleted ? { deleted: true } : {}),
      })
      nodes.push({ pos, node })
      return false
    }
    const stray =
      node.type.name === 'docProtected' ? (node.attrs.strayList as ListItemRef | null) : null
    if (stray) {
      refs.push({ numId: stray.numId, ilvl: stray.ilvl })
      nodes.push({ pos, node })
      return false
    }
    return true
  })
  const sites = computeListNumberSites(refs, defs)
  return nodes.map((n, i) => ({ ...n, site: sites[i] }))
}

function itemAt(items: NumberedItem[], pos: number): number {
  let lo = 0
  let hi = items.length - 1
  while (lo <= hi) {
    const mid = (lo + hi) >> 1
    const item = items[mid]
    if (pos < item.pos) hi = mid - 1
    else if (pos >= item.pos + item.node.nodeSize) lo = mid + 1
    else return mid
  }
  return -1
}

function numberText(
  items: NumberedItem[],
  sites: ReadonlyArray<ListNumberSite | null>,
  target: number,
  from: number,
  switches: NumberRefSwitches,
  words: NumberRefWords | undefined,
): string | null {
  const here = itemAt(items, from)
  const number = paragraphNumberText(sites, target, switches.context, {
    ...(here >= 0 ? { from: here } : {}),
    ...(switches.separator !== undefined ? { separator: switches.separator } : {}),
  })
  const shaped = (text: string) => (switches.cased ? formatFieldText(switches.cased, text) : text)
  if (number === null) return null
  if (!switches.aboveBelow) return shaped(number)
  if (!words) return null
  return shaped(`${number} ${items[target].pos < from ? words.above : words.below}`)
}

const sitesOf = (items: NumberedItem[]) => items.map((i) => i.site)

/** the result of a reference standing at `from` to the numbered paragraph at `targetPos` */
export function paragraphNumberAt(
  doc: PmNode,
  defs: Map<string, NumberingDef>,
  targetPos: number,
  from: number,
  switches: NumberRefSwitches,
  words?: NumberRefWords,
): string | null {
  const items = numberedItems(doc, defs)
  const target = items.findIndex((i) => i.pos === targetPos)
  return target < 0 ? null : numberText(items, sitesOf(items), target, from, switches, words)
}

interface NumberRef {
  range: FieldRange
  key: string
  /** what the numbering gives; null when it cannot tell */
  text: string | null
  agrees: boolean
}

const isNumberRefMark = (mark: { type: { name: string }; attrs: Record<string, unknown> }) =>
  mark.type.name === 'refField' && !!numberRefSwitches(mark.attrs.instr as string | null)

const hasNumberRef = (node: PmNode): boolean => {
  let found = false
  node.descendants((child) => {
    if (found) return false
    if (child.isText) found = child.marks.some(isNumberRefMark)
    return !found
  })
  return found
}

function numberRefs(
  doc: PmNode,
  defs: Map<string, NumberingDef>,
  words: NumberRefWords | undefined,
): NumberRef[] {
  const markType = doc.type.schema.marks.refField
  const ranges = markType
    ? collectFields(doc, markType).filter((r) => numberRefSwitches(r.instr))
    : []
  if (ranges.length === 0) return []
  const items = numberedItems(doc, defs)
  const sites = sitesOf(items)
  const byName = new Map<string, number>()
  items.forEach((item, i) => {
    for (const attr of ['bookmarks', 'hiddenBookmarks'] as const)
      for (const name of (item.node.attrs[attr] as string[] | null) ?? []) byName.set(name, i)
  })
  return ranges.map((range) => {
    const name = String(range.mark.attrs.name)
    const target = byName.get(name)
    const text =
      target === undefined
        ? null
        : numberText(items, sites, target, range.from, numberRefSwitches(range.instr)!, words)
    const shown = doc.textBetween(range.from, range.to, '').trim()
    return { range, key: `${name}\n${range.instr.trim()}`, text, agrees: text === shown }
  })
}

/** per bookmark and instruction, whether each reference in document order shows what the numbering gives */
type Agreement = ReadonlyMap<string, readonly boolean[]>

function agreementOf(refs: NumberRef[]): Agreement {
  const out = new Map<string, boolean[]>()
  for (const ref of refs) out.set(ref.key, [...(out.get(ref.key) ?? []), ref.agrees])
  return out
}

/**
 * Gives every paragraph-number reference in `tr` the number its paragraph shows. With
 * `trusted`, only the references that agreed there follow: one whose number in the file was
 * not what the numbering gives is left as the file has it. Returns how many changed.
 */
export function refreshNumberRefs(
  tr: Transform,
  defs: Map<string, NumberingDef>,
  options: { words?: NumberRefWords; trusted?: Agreement } = {},
): number {
  const refs = numberRefs(tr.doc, defs, options.words)
  const now = agreementOf(refs)
  const seen = new Map<string, number>()
  const follows = (ref: NumberRef): boolean => {
    const nth = seen.get(ref.key) ?? 0
    seen.set(ref.key, nth + 1)
    const before = options.trusted?.get(ref.key)
    if (!options.trusted) return true
    if (!before) return false
    // one came or went: which is which is not known, so all of them must have agreed
    return before.length === now.get(ref.key)!.length ? before[nth] : before.every(Boolean)
  }
  const jobs = refs.filter(
    (ref) =>
      follows(ref) &&
      ref.text !== null &&
      !isFieldLocked(ref.range.mark) &&
      (!ref.agrees || ref.range.mark.attrs.dirty === true),
  )
  for (const job of [...jobs].reverse()) writeField(tr, job.range, job.text!)
  return jobs.filter((j) => !j.agrees).length
}

interface NumberRefState {
  /** the document holds a paragraph-number reference */
  has: boolean
  /** which references show what the numbering gives */
  trusted: Agreement
}

const NONE: NumberRefState = { has: false, trusted: new Map() }
const numberRefKey = new PluginKey<NumberRefState>('numberRefs')

const interfaceWords = (): NumberRefWords => ({
  above: t('ribbonCrossRefAbove'),
  below: t('ribbonCrossRefBelow'),
})

const bringsNumberRef = (tr: Transaction): boolean =>
  tr.steps.some((step) => {
    const { slice, mark } = step as {
      slice?: { content: PmNode['content'] }
      mark?: Parameters<typeof isNumberRefMark>[0]
    }
    let found = !!mark && isNumberRefMark(mark)
    slice?.content.forEach((node) => {
      found ||= node.isText ? node.marks.some(isNumberRefMark) : hasNumberRef(node)
    })
    return found
  })

export const NumberRefsExtension = Extension.create({
  name: 'numberRefs',
  addProseMirrorPlugins() {
    const defsOf = () => this.editor.storage.listNumbering.defs
    let seenDefs = defsOf()
    const stateOf = (doc: PmNode): NumberRefState => {
      const refs = numberRefs(doc, defsOf(), interfaceWords())
      return refs.length === 0 ? NONE : { has: true, trusted: agreementOf(refs) }
    }
    return [
      new Plugin<NumberRefState>({
        key: numberRefKey,
        state: {
          init: (_, state) => (hasNumberRef(state.doc) ? stateOf(state.doc) : NONE),
          apply: (tr, old) => {
            if (!tr.docChanged) return old
            return old.has || bringsNumberRef(tr) ? stateOf(tr.doc) : old
          },
        },
        appendTransaction(transactions, oldState, newState) {
          const defs = defsOf()
          const redefined = defs !== seenDefs
          seenDefs = defs
          if (!redefined && !transactions.some((tr) => tr.docChanged)) return null
          if (!numberRefKey.getState(newState)?.has) return null
          // a reference that agreed with its paragraph before the change follows it
          const trusted = numberRefKey.getState(oldState)?.trusted ?? NONE.trusted
          if (trusted.size === 0) return null
          const tr = newState.tr.setMeta(TRACK_IGNORE, true)
          refreshNumberRefs(tr, defs, { words: interfaceWords(), trusted })
          return tr.docChanged ? tr : null
        },
      }),
    ]
  },
})
