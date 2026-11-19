/**
 * Tracked changes outside the main body (Word's "stories"): footnotes,
 * endnotes, header / footer parts and textboxes. The editor holds these as
 * display models (NoteRun / Run lists with ins / del / rPrChange records), not
 * as ProseMirror content, so they are enumerated from the models here, listed
 * next to the body's revisions, and accepted / rejected by rewriting the runs.
 * The save path replays the same decisions on the part's original bytes
 * (NoteInfo / HeaderFooter revisionsResolved).
 */
import type {
  HeaderFooter,
  HfParagraph,
  NoteInfo,
  NoteRun,
  RevisionInfo,
  RevisionResolution,
  Run,
  TextboxDisplay,
  TextboxParaDisplay,
} from '@genoffice/docx-engine'
import type { Node as PmNode } from '@tiptap/pm/model'
import type { HfView } from '../doc-state'
import { editKey, ownHf, type HfKind, type HfSectionState } from '../hf-sections'
import type { StringKey } from '../i18n/locale'
import type { RevisionRange } from './revisions'

export type StoryKind = 'footnote' | 'endnote' | 'header' | 'footer' | 'textbox'

export type StoryRevisionType = 'ins' | 'del' | 'rPrChange'

export interface StoryLocation {
  kind: StoryKind
  /** unique within the current models; re-resolved against them before an action */
  id: string
  /** the changed text as the list shows it */
  text: string
  record: RevisionInfo
  type: StoryRevisionType
  /** footnote / endnote: the note's w:id */
  noteId?: string
  /** header / footer: the strip slot the part is shown in */
  hf?: { section: number; kind: HfKind; variant: HfView }
  /** textbox: PM position of the anchor paragraph node, index of the box in it */
  textbox?: { pos: number; box: number }
  para: number
  /** layout-table cell inside the paragraph row and the paragraph in it */
  cell?: [number, number]
  /** run index range [start, end) in that paragraph */
  runs: [number, number]
}

export type StoryRevision = RevisionRange & { story: StoryLocation }

export const STORY_LABEL: Record<StoryKind, StringKey> = {
  footnote: 'appRevLocFootnote',
  endnote: 'appRevLocEndnote',
  header: 'appRevLocHeader',
  footer: 'appRevLocFooter',
  textbox: 'appRevLocTextbox',
}

const sameRecord = (a: RevisionInfo | undefined, b: RevisionInfo | undefined) =>
  !!a && !!b && a.author === b.author && a.date === b.date && a.id === b.id && a.move === b.move

type AnyRun = Run | NoteRun

interface RunSite {
  kind: StoryKind
  partKey: string
  para: number
  cell?: [number, number]
  from: number
  to: number
  noteId?: string
  hf?: StoryLocation['hf']
  textbox?: StoryLocation['textbox']
}

/** contiguous runs sharing one ins / del / rPrChange record become one revision */
function runRevisions(runs: readonly AnyRun[], site: RunSite): StoryRevision[] {
  const out: StoryRevision[] = []
  const push = (type: StoryRevisionType, record: RevisionInfo, start: number, end: number) => {
    const text = runs
      .slice(start, end)
      .map((r) => r.text)
      .join('')
      .replace(/\s+/g, ' ')
      .trim()
    const kind: RevisionRange['kind'] =
      type === 'rPrChange'
        ? 'rPrChange'
        : record.move
          ? type === 'ins'
            ? 'moveTo'
            : 'moveFrom'
          : type
    const cell = site.cell ? `${site.cell[0]}.${site.cell[1]}` : '-'
    const overlapped = runs.slice(start, end).some((r) => r.ins && r.del)
    out.push({
      from: site.from,
      to: site.to,
      kind,
      author: record.author,
      ...(record.date ? { date: record.date } : {}),
      ...(record.move ? { moveId: record.move } : {}),
      ...(overlapped && type !== 'rPrChange' ? { overlapped: true } : {}),
      story: {
        kind: site.kind,
        id: `${site.kind}:${site.partKey}:${site.para}:${cell}:${start}:${type}`,
        text,
        record,
        type,
        ...(site.noteId ? { noteId: site.noteId } : {}),
        ...(site.hf ? { hf: site.hf } : {}),
        ...(site.textbox ? { textbox: site.textbox } : {}),
        para: site.para,
        ...(site.cell ? { cell: site.cell } : {}),
        runs: [start, end],
      },
    })
  }
  for (const type of ['ins', 'del', 'rPrChange'] as const) {
    let start = -1
    let record: RevisionInfo | undefined
    runs.forEach((run, i) => {
      const cur = (run as Run)[type]
      if (cur && start >= 0 && sameRecord(cur, record)) return
      if (start >= 0) push(type, record!, start, i)
      start = cur ? i : -1
      record = cur
    })
    if (start >= 0) push(type, record!, start, runs.length)
  }
  return out
}

function paraRevisions(
  para: Pick<HfParagraph, 'runs' | 'cells'> | Pick<TextboxParaDisplay, 'runs' | 'cells'>,
  site: Omit<RunSite, 'cell'>,
): StoryRevision[] {
  const out = runRevisions(para.runs, site)
  para.cells?.forEach((cell, ci) => {
    cell.paras.forEach((runs, pi) => out.push(...runRevisions(runs, { ...site, cell: [ci, pi] })))
  })
  return out
}

/** w:id of every docNoteRef of `kind`, mapped to its position in the body */
function noteRefPositions(doc: PmNode, kind: 'footnote' | 'endnote'): Map<string, number> {
  const out = new Map<string, number>()
  doc.descendants((node, pos) => {
    if (
      node.type.name === 'docNoteRef' &&
      node.attrs.kind === kind &&
      !out.has(String(node.attrs.id))
    )
      out.set(String(node.attrs.id), pos)
  })
  return out
}

export function collectNoteRevisions(
  doc: PmNode,
  kind: 'footnote' | 'endnote',
  notes: readonly NoteInfo[],
): StoryRevision[] {
  const out: StoryRevision[] = []
  const refs = noteRefPositions(doc, kind)
  for (const note of notes) {
    if (!note.richParas) continue
    const at = refs.get(note.id) ?? doc.content.size
    note.richParas.forEach((runs, pi) => {
      out.push(
        ...runRevisions(runs, {
          kind,
          partKey: note.id,
          para: pi,
          from: at,
          to: at,
          noteId: note.id,
        }),
      )
    })
  }
  // notes list in body-reference order, like the footnote area
  return out.sort((a, b) => a.from - b.from)
}

const HF_VIEWS: readonly HfView[] = ['default', 'first', 'even']

export interface HfSlot {
  si: number
  kind: HfKind
  variant: HfView
  value: HeaderFooter
}

/** the part a slot shows: the referenced part, else the section's own (unlinked / new) part */
function hfSlotKey(state: HfSectionState, si: number, kind: HfKind, variant: HfView): string {
  const sec = state.sections[si]
  const rId = sec ? (kind === 'header' ? sec.headerRefs : sec.footerRefs)[variant] : undefined
  if (rId) return `ref:${rId}`
  return sec ? `edit:${editKey(sec.lastBlockIndex, kind, variant)}` : 'doc'
}

/** every own header / footer part once (sections sharing a part list it under the first) */
export function ownHfSlots(state: HfSectionState): HfSlot[] {
  const out: HfSlot[] = []
  const seen = new Set<string>()
  for (let si = 0; si < Math.max(1, state.sections.length); si++) {
    for (const kind of ['header', 'footer'] as const) {
      for (const variant of HF_VIEWS) {
        const value = ownHf(state, si, kind, variant)?.value
        const key = hfSlotKey(state, si, kind, variant)
        if (!value?.paras || seen.has(key)) continue
        seen.add(key)
        out.push({ si, kind, variant, value })
      }
    }
  }
  return out
}

export function collectHfRevisions(state: HfSectionState): StoryRevision[] {
  const out: StoryRevision[] = []
  for (const { si, kind, variant, value } of ownHfSlots(state)) {
    value.paras!.forEach((para, pi) => {
      out.push(
        ...paraRevisions(para, {
          kind,
          partKey: `${si}:${variant}`,
          para: pi,
          from: -1,
          to: -1,
          hf: { section: si, kind, variant },
        }),
      )
    })
  }
  return out
}

export function collectTextboxRevisions(doc: PmNode): StoryRevision[] {
  const out: StoryRevision[] = []
  doc.descendants((node, pos) => {
    const boxes = node.attrs?.textboxes as TextboxDisplay[] | null | undefined
    if (!boxes) return
    boxes.forEach((box, bi) => {
      box.paras.forEach((para, pi) => {
        out.push(
          ...paraRevisions(para, {
            kind: 'textbox',
            partKey: `${pos}:${bi}`,
            para: pi,
            from: pos + 1,
            to: pos + 1,
            textbox: { pos, box: bi },
          }),
        )
      })
    })
  })
  return out
}

/** the current models' revision matching a (possibly stale) entry: by id, else by record and text */
export function liveStoryRevision(
  live: readonly StoryRevision[],
  target: StoryRevision,
): StoryRevision | undefined {
  const s = target.story
  return (
    live.find((r) => r.story.id === s.id) ??
    live.find(
      (r) =>
        r.story.kind === s.kind &&
        r.story.type === s.type &&
        r.story.noteId === s.noteId &&
        r.story.textbox?.pos === s.textbox?.pos &&
        sameRecord(r.story.record, s.record) &&
        r.story.text === s.text,
    )
  )
}

type Mode = 'accept' | 'reject'

function resolveRun<T extends AnyRun>(run: T, types: Set<StoryRevisionType>, mode: Mode): T | null {
  let next: T = run
  const strip = (field: 'ins' | 'del' | 'rPrChange') => {
    const { [field]: _gone, ...rest } = next as Run
    next = rest as T
  }
  if (types.has('ins')) {
    if (mode === 'reject') return null
    strip('ins')
  }
  if (types.has('del')) {
    if (mode === 'accept') return null
    strip('del')
  }
  if (types.has('rPrChange') && (next as Run).rPrChange) {
    const old = (next as Run).rPrChange!.old
    strip('rPrChange')
    if (mode === 'reject' && old) {
      const restored = { ...(next as Run) } as Record<string, unknown>
      for (const key of ['bold', 'italic', 'underline', 'strike'] as const) {
        if (old[key]) restored[key] = true
        else delete restored[key]
      }
      for (const key of [
        'color',
        'sizeHalfPoints',
        'font',
        'fontAscii',
        'charSpacingTwips',
        'charScalePct',
        'highlight',
        'vertAlign',
        'styleId',
      ] as const) {
        if (old[key] !== undefined) restored[key] = old[key]
        else delete restored[key]
      }
      next = restored as unknown as T
    }
  }
  return next
}

/** the run list after the given revisions (all located in it) are accepted / rejected */
export function resolveRuns<T extends AnyRun>(
  runs: readonly T[],
  targets: readonly StoryLocation[],
  mode: Mode,
): T[] {
  const out: T[] = []
  runs.forEach((run, i) => {
    const types = new Set<StoryRevisionType>()
    for (const t of targets) if (i >= t.runs[0] && i < t.runs[1]) types.add(t.type)
    const next = types.size ? resolveRun(run, types, mode) : run
    if (next) out.push(next)
  })
  return out
}

function resolvePara<P extends Pick<HfParagraph, 'runs' | 'cells'>>(
  para: P,
  targets: readonly StoryLocation[],
  mode: Mode,
): P {
  const direct = targets.filter((t) => !t.cell)
  const next: P = direct.length ? { ...para, runs: resolveRuns(para.runs, direct, mode) } : para
  if (!para.cells || targets.every((t) => !t.cell)) return next
  return {
    ...next,
    cells: para.cells.map((cell, ci) => {
      const mine = targets.filter((t) => t.cell?.[0] === ci)
      if (!mine.length) return cell
      return {
        ...cell,
        rewrite: true as const,
        paras: cell.paras.map((runs, pi) =>
          resolveRuns(
            runs,
            mine.filter((t) => t.cell![1] === pi),
            mode,
          ),
        ),
      }
    }),
  }
}

/** how the save path replays the decision on the part bytes */
export function resolutionOf(
  targets: readonly StoryLocation[],
  mode: Mode,
  all = false,
): RevisionResolution[] {
  if (all) return [{ mode }]
  const ids = targets.map((t) => t.record.id).filter((id): id is string => !!id)
  const out: RevisionResolution[] = []
  if (ids.length) out.push({ mode, filter: { ids: [...new Set(ids)] } })
  for (const t of targets) {
    if (t.record.id) continue
    out.push({
      mode,
      filter: { author: t.record.author, ...(t.record.date ? { date: t.record.date } : {}) },
    })
  }
  return out
}

const noteText = (richParas: NoteRun[][]) =>
  richParas
    .map((runs) =>
      runs
        .filter((r) => !r.del)
        .map((r) => r.text)
        .join(''),
    )
    .join('\n')

export function resolveNotes(
  notes: readonly NoteInfo[],
  targets: readonly StoryLocation[],
  mode: Mode,
  all = false,
): NoteInfo[] {
  return notes.map((note) => {
    const mine = targets.filter((t) => t.noteId === note.id)
    if (!mine.length || !note.richParas) return note
    const richParas = note.richParas.map((runs, pi) =>
      resolveRuns(
        runs,
        mine.filter((t) => t.para === pi),
        mode,
      ),
    )
    return {
      ...note,
      richParas,
      text: noteText(richParas),
      revisionsResolved: [...(note.revisionsResolved ?? []), ...resolutionOf(mine, mode, all)],
    }
  })
}

/** `pending`: the value is already an unsaved edit (regenerates on save); else the part bytes are kept */
export function resolveHfValue(
  value: HeaderFooter,
  targets: readonly StoryLocation[],
  mode: Mode,
  pending: boolean,
  all = false,
): HeaderFooter {
  if (!value.paras) return value
  const paras = value.paras.map((para, pi) =>
    resolvePara(
      para,
      targets.filter((t) => t.para === pi),
      mode,
    ),
  )
  const text = paras
    .filter((p) => !p.cells)
    .map((p) => p.runs.map((r) => r.text).join(''))
    .join('')
  return {
    ...value,
    text,
    paras,
    revisionsResolved: [...(value.revisionsResolved ?? []), ...resolutionOf(targets, mode, all)],
    ...(!pending || value.keepContent ? { keepContent: true as const } : {}),
  }
}

export function resolveTextboxes(
  boxes: readonly TextboxDisplay[],
  targets: readonly StoryLocation[],
  mode: Mode,
): TextboxDisplay[] {
  return boxes.map((box, bi) => {
    const mine = targets.filter((t) => t.textbox?.box === bi)
    if (!mine.length) return box
    return {
      ...box,
      paras: box.paras.map((para, pi) =>
        resolvePara(
          para,
          mine.filter((t) => t.para === pi),
          mode,
        ),
      ),
    }
  })
}
