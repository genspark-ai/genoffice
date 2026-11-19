import type { Editor } from '@tiptap/core'
import type { Node as PMNode } from '@tiptap/pm/model'
import type { Transaction } from '@tiptap/pm/state'
import { Step, StepResult } from '@tiptap/pm/transform'
import type { NoteProps, SectionInfo, SectionSettings } from '@genoffice/docx-engine'
import type { HyphenationSettings } from './hyphenation'
import { TRACK_IGNORE } from './revisions'

type NoteKind = 'footnote' | 'endnote'
export type NotePropsByKind = Record<NoteKind, NoteProps>

/**
 * Page setup lives in app state, not in the editor document. A change to it
 * enters the editor's undo history as a step that leaves the content alone and
 * carries what to put back, so Undo and Redo replay it in order with text edits.
 */
export interface PageSetupState {
  /** the last section's settings, as the save writes them */
  section: SectionSettings | null
  sections: SectionInfo[]
  trailingStartType: SectionInfo['startType'] | null
  mirrorMargins: boolean
  gutterAtTop: boolean
  titlePg: boolean
  evenOddHf: boolean
  /** the document-wide note properties in effect: settings.xml under this session's edit */
  notes: NotePropsByKind
  /** the note properties in effect in each section */
  sectionNotes: NotePropsByKind[]
  autoHyphenation: boolean
  hyphenation: HyphenationSettings
}

type SettingsPatch = Partial<SectionSettings>

export interface SectionPatch {
  settings?: SettingsPatch
  startType?: SectionInfo['startType']
  /** a section behind an unsaved break: the sectPr its break paragraph holds */
  sectPrXml?: string
}

/** the values one side of a change holds, only where the two sides differ */
export interface PageSetupPatch {
  section?: SettingsPatch
  sections?: Record<number, SectionPatch>
  trailingStartType?: SectionInfo['startType'] | null
  mirrorMargins?: boolean
  gutterAtTop?: boolean
  titlePg?: boolean
  evenOddHf?: boolean
  notes?: Partial<NotePropsByKind>
  autoHyphenation?: boolean
  hyphenation?: Partial<HyphenationSettings>
}

const same = (a: unknown, b: unknown): boolean => a === b || JSON.stringify(a) === JSON.stringify(b)

// absent settings the save reads as "what the file has": a save may have come
// in between, so putting one back has to say the value
const WRITTEN_WHEN_ABSENT: Record<string, unknown> = { gutter: 0, colEqualWidth: true }

function settingsDiff(
  before: SectionSettings,
  after: SectionSettings,
): [SettingsPatch, SettingsPatch] | null {
  const undo: Record<string, unknown> = {}
  const redo: Record<string, unknown> = {}
  const a = before as unknown as Record<string, unknown>
  const b = after as unknown as Record<string, unknown>
  for (const key of new Set([...Object.keys(a), ...Object.keys(b)])) {
    if (same(a[key], b[key])) continue
    undo[key] = a[key] ?? WRITTEN_WHEN_ABSENT[key]
    redo[key] = b[key] ?? WRITTEN_WHEN_ABSENT[key]
  }
  return Object.keys(undo).length > 0 ? [undo as SettingsPatch, redo as SettingsPatch] : null
}

const SCALARS = [
  'trailingStartType',
  'mirrorMargins',
  'gutterAtTop',
  'titlePg',
  'evenOddHf',
  'autoHyphenation',
] as const

const NOTE_KINDS = ['footnote', 'endnote'] as const
const NOTE_FIELDS = ['pos', 'numFmt', 'numStart', 'numRestart'] as const
const NOTE_DEFAULTS: Record<NoteKind, Required<NoteProps>> = {
  footnote: { pos: 'pageBottom', numFmt: 'decimal', numStart: 1, numRestart: 'continuous' },
  endnote: { pos: 'docEnd', numFmt: 'lowerRoman', numStart: 1, numRestart: 'continuous' },
}

// a note field the save is not told keeps what the file has, so a patch states
// every field its change touched, with the value in effect where none was set
function statedNoteProps(
  kind: NoteKind,
  fields: readonly (typeof NOTE_FIELDS)[number][],
  own: NoteProps | undefined,
  inEffect: NoteProps | undefined,
): NoteProps {
  return Object.fromEntries(
    fields.map((f) => [f, own?.[f] ?? inEffect?.[f] ?? NOTE_DEFAULTS[kind][f]]),
  )
}

function stateSectionNotes(
  undo: SettingsPatch,
  redo: SettingsPatch,
  before: NotePropsByKind | undefined,
  after: NotePropsByKind | undefined,
): void {
  for (const kind of NOTE_KINDS) {
    const key = `${kind}Pr` as const
    if (!(key in undo)) continue
    const fields = NOTE_FIELDS.filter(
      (f) => undo[key]?.[f] !== undefined || redo[key]?.[f] !== undefined,
    )
    undo[key] = statedNoteProps(kind, fields, undo[key], before?.[kind])
    redo[key] = statedNoteProps(kind, fields, redo[key], after?.[kind])
  }
}

/**
 * What undoing and redoing a change put back; null when nothing changed, or
 * when sections were added or removed (their settings follow the break
 * paragraph in the document, not a patch by index).
 */
export function pageSetupChange(
  before: PageSetupState,
  after: PageSetupState,
): { undo: PageSetupPatch; redo: PageSetupPatch } | null {
  if (before.sections.length !== after.sections.length) return null
  const undo: PageSetupPatch = {}
  const redo: PageSetupPatch = {}
  if (before.section && after.section) {
    const diff = settingsDiff(before.section, after.section)
    if (diff) {
      ;[undo.section, redo.section] = diff
      stateSectionNotes(diff[0], diff[1], before.sectionNotes.at(-1), after.sectionNotes.at(-1))
    }
  }
  before.sections.forEach((was, i) => {
    const now = after.sections[i]!
    const u: SectionPatch = {}
    const r: SectionPatch = {}
    const diff = was.settings === now.settings ? null : settingsDiff(was.settings, now.settings)
    if (diff) {
      ;[u.settings, r.settings] = diff
      stateSectionNotes(diff[0], diff[1], before.sectionNotes[i], after.sectionNotes[i])
    }
    if (was.startType !== now.startType) {
      u.startType = was.startType
      r.startType = now.startType
    }
    if (was.pendingBreak && now.pendingBreak && was.sectPrXml !== now.sectPrXml) {
      u.sectPrXml = was.sectPrXml
      r.sectPrXml = now.sectPrXml
    }
    if (Object.keys(u).length === 0) return
    ;(undo.sections ??= {})[i] = u
    ;(redo.sections ??= {})[i] = r
  })
  for (const key of SCALARS) {
    if (before[key] === after[key]) continue
    Object.assign(undo, { [key]: before[key] })
    Object.assign(redo, { [key]: after[key] })
  }
  for (const kind of NOTE_KINDS) {
    const fields = NOTE_FIELDS.filter((f) => before.notes[kind][f] !== after.notes[kind][f])
    if (fields.length === 0) continue
    ;(undo.notes ??= {})[kind] = statedNoteProps(kind, fields, before.notes[kind], undefined)
    ;(redo.notes ??= {})[kind] = statedNoteProps(kind, fields, after.notes[kind], undefined)
  }
  for (const key of ['caps', 'limit', 'zoneTwips'] as const) {
    if (before.hyphenation[key] === after.hyphenation[key]) continue
    undo.hyphenation = { ...undo.hyphenation, [key]: before.hyphenation[key] }
    redo.hyphenation = { ...redo.hyphenation, [key]: after.hyphenation[key] }
  }
  return Object.keys(undo).length > 0 ? { undo, redo } : null
}

/** the section list with a patch's values put back; sections the patch names that are gone are skipped */
export function patchSections(sections: SectionInfo[], patch: PageSetupPatch): SectionInfo[] {
  const byIndex = patch.sections
  if (!byIndex) return sections
  return sections.map((s, i) => {
    const p = byIndex[i]
    if (!p) return s
    return {
      ...s,
      ...(p.settings && { settings: { ...s.settings, ...p.settings } }),
      ...(p.startType && { startType: p.startType }),
      ...(p.sectPrXml !== undefined && s.pendingBreak && { sectPrXml: p.sectPrXml }),
    }
  })
}

export class PageSetupStep extends Step {
  constructor(
    /** what applying this step puts in place, and what undoing it puts back */
    readonly put: PageSetupPatch,
    readonly back: PageSetupPatch,
  ) {
    super()
  }

  // a new document object over the same content: the change counts as an edit
  // for everything that watches the document's identity (Repeat, a save in flight)
  apply(doc: PMNode): StepResult {
    return StepResult.ok(doc.type.create(doc.attrs, doc.content, doc.marks))
  }

  invert(): Step {
    return new PageSetupStep(this.back, this.put)
  }

  map(): Step {
    return this
  }

  toJSON(): { stepType: string } {
    return { stepType: 'pageSetup' }
  }
}

const RECORDING = 'pageSetupRecording'

/**
 * Puts a change that app state already holds into the undo history as its own
 * step. `joinPrevious`: the command also edited the document (a break
 * paragraph's sectPr), and the two undo together.
 */
export function recordPageSetupChange(
  editor: Editor,
  change: { undo: PageSetupPatch; redo: PageSetupPatch },
  joinPrevious: boolean,
): void {
  const tr = editor.state.tr
    .step(new PageSetupStep(change.redo, change.undo))
    .setMeta(RECORDING, true)
    .setMeta(TRACK_IGNORE, true)
  if (joinPrevious) tr.setMeta('appendedTransaction', editor.state.tr)
  editor.view.dispatch(tr)
}

/** the patches an undo or redo transaction asks app state to take, in order */
export function replayedPageSetup(tr: Transaction): PageSetupPatch[] {
  if (tr.getMeta(RECORDING)) return []
  return tr.steps.flatMap((step) => (step instanceof PageSetupStep ? [step.put] : []))
}
