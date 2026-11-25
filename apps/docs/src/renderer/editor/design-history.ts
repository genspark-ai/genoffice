import type { Editor } from '@tiptap/core'
import type { Node as PMNode } from '@tiptap/pm/model'
import type { Transaction } from '@tiptap/pm/state'
import { Step, StepResult } from '@tiptap/pm/transform'
import {
  DEFAULT_THEME_COLORS,
  listStylesXml,
  readDocxStylesXml,
  readDocxThemeFormat,
  type StyleUpsert,
  type ThemeColors,
  type ThemeFonts,
  type ThemeFormat,
} from '@genoffice/docx-engine'
import { TRACK_IGNORE } from './revisions'

/**
 * The theme and the style definitions live in app state, not in the editor
 * document. A Design pick enters the editor's undo history as a step that
 * leaves the content alone and carries what to put back, the way page setup does.
 */
export interface DesignTheme {
  fonts: ThemeFonts | null
  fontsDirty: boolean
  colors: ThemeColors | null
  colorsDirty: boolean
  /** the format scheme of a whole-theme pick the save has not written yet */
  format: ThemeFormat | null
}

export interface DesignState extends DesignTheme {
  /** style definitions the save has not written yet, by style id */
  styles: Record<string, StyleUpsert>
}

/** one side of a change: the theme when it changed, and the pending definition of each style it touched */
export interface DesignSide {
  theme?: DesignTheme
  styles?: Record<string, StyleUpsert | null>
}

/** a side said without reference to the file: what to write when a save came in between */
interface StatedSide {
  /** each touched style as styles.xml held it; null where it did not exist */
  styles: Record<string, string | null>
  format: ThemeFormat | null
}

export interface DesignChange {
  /** the reading of the file the change was made on */
  parsed: object
  undo: DesignSide
  redo: DesignSide
  stated: Promise<{ undo: StatedSide; redo: StatedSide }>
}

/** the pending definitions with `up` queued: a whole definition stands alone, a partial one merges into what is pending */
export function withStyleUpsert(
  pending: Record<string, StyleUpsert>,
  up: StyleUpsert,
): Record<string, StyleUpsert> {
  const prev = pending[up.styleId]
  // a removal or a copied definition stands alone; a style made again after its removal starts clean
  const whole = up.remove || up.sourceXml !== undefined
  return {
    ...pending,
    [up.styleId]: whole
      ? up
      : prev?.remove
        ? { ...up, replace: true as const }
        : prev
          ? {
              ...prev,
              ...up,
              pPr: up.pPr || prev.pPr ? { ...prev.pPr, ...up.pPr } : undefined,
              rPr: up.rPr || prev.rPr ? { ...prev.rPr, ...up.rPr } : undefined,
            }
          : up,
  }
}

const same = (a: unknown, b: unknown): boolean => a === b || JSON.stringify(a) === JSON.stringify(b)

const themeOf = ({ fonts, fontsDirty, colors, colorsDirty, format }: DesignTheme): DesignTheme => ({
  fonts,
  fontsDirty,
  colors,
  colorsDirty,
  format,
})

/** what undoing and redoing a change put back; null when nothing changed */
export function designSides(
  before: DesignState,
  after: DesignState,
): { undo: DesignSide; redo: DesignSide } | null {
  const undo: DesignSide = {}
  const redo: DesignSide = {}
  if (!same(themeOf(before), themeOf(after))) {
    undo.theme = themeOf(before)
    redo.theme = themeOf(after)
  }
  const touched = [
    ...new Set([...Object.keys(before.styles), ...Object.keys(after.styles)]),
  ].filter((id) => !same(before.styles[id], after.styles[id]))
  if (touched.length > 0) {
    undo.styles = Object.fromEntries(touched.map((id) => [id, before.styles[id] ?? null]))
    redo.styles = Object.fromEntries(touched.map((id) => [id, after.styles[id] ?? null]))
  }
  return undo.theme || undo.styles ? { undo, redo } : null
}

async function statedSide(
  bytes: Uint8Array,
  state: DesignState,
  touched: string[],
  fileFormat: Promise<ThemeFormat | null> | null,
): Promise<StatedSide> {
  const styles: Record<string, string | null> = {}
  if (touched.length > 0) {
    const xml = await readDocxStylesXml(bytes, Object.values(state.styles))
    const byId = new Map(listStylesXml(xml).map((s) => [s.styleId, s.xml]))
    for (const id of touched) styles[id] = byId.get(id) ?? null
  }
  return { styles, format: state.format ?? (fileFormat ? await fileFormat : null) }
}

/** the change between two states of the document read as `parsed`, whose bytes are `bytes` */
export function designChange(
  parsed: object,
  bytes: Uint8Array,
  before: DesignState,
  after: DesignState,
): DesignChange | null {
  const sides = designSides(before, after)
  if (!sides) return null
  const touched = Object.keys(sides.undo.styles ?? {})
  const fileFormat = same(before.format, after.format) ? null : readDocxThemeFormat(bytes)
  const stated = Promise.all([
    statedSide(bytes, before, touched, fileFormat),
    statedSide(bytes, after, touched, fileFormat),
  ]).then(([undo, redo]) => ({ undo, redo }))
  // a file that cannot be read again leaves the step working on its own reading
  stated.catch(() => {})
  return { parsed, ...sides, stated }
}

const DEFAULT_FONTS: ThemeFonts = { major: 'Calibri Light', minor: 'Calibri', eastAsia: '' }

/**
 * `state` with one side of a change put back. `saved`: a save came in between,
 * so the file now holds the other side and every value has to be said in full.
 */
export function withDesignSide(
  state: DesignState,
  side: DesignSide,
  saved: StatedSide | null,
): DesignState {
  const next = { ...state }
  if (side.theme) {
    Object.assign(next, side.theme)
    if (saved) {
      next.fonts = side.theme.fonts ?? DEFAULT_FONTS
      next.colors = side.theme.colors ?? { ...DEFAULT_THEME_COLORS }
      next.fontsDirty = true
      next.colorsDirty = true
      next.format = saved.format
    }
  }
  if (side.styles) {
    const styles = { ...state.styles }
    for (const [id, pending] of Object.entries(side.styles)) {
      if (saved) {
        const xml = saved.styles[id]
        styles[id] = xml ? { styleId: id, sourceXml: xml } : { styleId: id, remove: true }
      } else if (pending) styles[id] = pending
      else delete styles[id]
    }
    next.styles = styles
  }
  return next
}

/** a step's change arrives once its command has read what it needs; until then the step puts nothing back */
export interface DesignSlot {
  change: DesignChange | null
  /** the step stands undone while its change is still to arrive: the change waits for a redo */
  undone?: boolean
}

/** counts what a waiting style change gives way to: every style change made, every undo and redo of a Design step that holds its change */
export interface DesignClock {
  current: number
}

/**
 * A change that has to read the file before it can apply. The returned check
 * says what becomes of it once the read is done: it applies, it is kept for a
 * redo because its own step was undone meanwhile, or (null) it is dropped
 * after a newer style change or an undo or redo of another Design step.
 */
export function pendingDesign(clock: DesignClock, slot: DesignSlot): () => 'apply' | 'keep' | null {
  const at = ++clock.current
  return () => (clock.current !== at ? null : slot.undone ? 'keep' : 'apply')
}

export class DesignStep extends Step {
  constructor(
    readonly slot: DesignSlot,
    /** which side applying this step puts in place */
    readonly side: 'undo' | 'redo',
  ) {
    super()
  }

  // a new document object over the same content, so Repeat and a save in flight see an edit
  apply(doc: PMNode): StepResult {
    return StepResult.ok(doc.type.create(doc.attrs, doc.content, doc.marks))
  }

  invert(): Step {
    return new DesignStep(this.slot, this.side === 'redo' ? 'undo' : 'redo')
  }

  map(): Step {
    return this
  }

  toJSON(): { stepType: string } {
    return { stepType: 'design' }
  }
}

const RECORDING = 'designRecording'

/** puts a Design change app state already holds into the undo history as one step */
export function recordDesignStep(editor: Editor, slot: DesignSlot): void {
  editor.view.dispatch(
    editor.state.tr
      .step(new DesignStep(slot, 'redo'))
      .setMeta(RECORDING, true)
      .setMeta(TRACK_IGNORE, true),
  )
}

/** the sides an undo or redo transaction asks app state to take, in order */
export function replayedDesign(
  tr: Transaction,
): Array<{ change: DesignChange; side: 'undo' | 'redo' }> {
  if (tr.getMeta(RECORDING)) return []
  return tr.steps.flatMap((step) => {
    if (!(step instanceof DesignStep)) return []
    if (step.slot.change) return [{ change: step.slot.change, side: step.side }]
    step.slot.undone = step.side === 'undo'
    return []
  })
}

/** an undo or redo of a Design step that holds its change; one still waiting only notes where it stands */
export const replaysDesign = (tr: Transaction): boolean =>
  !tr.getMeta(RECORDING) &&
  tr.steps.some((step) => step instanceof DesignStep && step.slot.change !== null)
