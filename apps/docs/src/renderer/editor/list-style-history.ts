import type { Editor } from '@tiptap/core'
import type { Node as PMNode } from '@tiptap/pm/model'
import type { Transaction } from '@tiptap/pm/state'
import { Step, StepResult } from '@tiptap/pm/transform'
import type { CustomNumberingLevel, NumberingDef, StyleUpsert } from '@genoffice/docx-engine'
import type { PendingNumbering } from '../doc-state'
import { nineListLevels, type ListStyleEntry } from '../list-styles'
import { modifyListStyle, type NumberingContext } from '../numbering-actions'
import { TRACK_IGNORE } from './revisions'

type StylePending = Pick<PendingNumbering, 'levelEdits' | 'styleDefs' | 'styleNums'>

/**
 * A list style's levels live in app state and in the numbering the editor
 * draws from, not in the editor document. Modify Style enters the undo history
 * as a step that leaves the content alone and carries both sides of the change.
 */
export interface ListStyleSide {
  name: string
  /** the style's nine levels: what to modify it to once a save has written the other side */
  levels: CustomNumberingLevel[]
  /** the unsaved numbering entries that define the style */
  pending: StylePending
  /** the unsaved definition of the style itself (its name); null when there is none */
  upsert: StyleUpsert | null
  /** the levels the editor draws each list of the style with, by numId */
  drawn: Record<string, NumberingDef['levels']>
}

export interface ListStyleChange {
  /** the reading of the file the change was made on */
  parsed: object
  styleId: string
  /** the definition the style's levels are written to, when the document holds the style */
  abstractNumId?: string
  undo: ListStyleSide
  redo: ListStyleSide
}

const same = (a: unknown, b: unknown): boolean => JSON.stringify(a) === JSON.stringify(b)

const defsOf = (editor: Editor): Map<string, NumberingDef> =>
  (editor.storage.listNumbering as { defs: Map<string, NumberingDef> }).defs

function pendingOf(p: PendingNumbering, styleId: string, abstractNumId?: string): StylePending {
  return {
    levelEdits: p.levelEdits.filter(
      (e) => abstractNumId !== undefined && e.abstractNumId === abstractNumId,
    ),
    styleDefs: p.styleDefs.filter((s) => s.styleId === styleId),
    styleNums: p.styleNums.filter((l) => l.styleId === styleId),
  }
}

function drawnOf(editor: Editor, styleId: string, abstractNumId?: string) {
  const drawn: Record<string, NumberingDef['levels']> = {}
  for (const [numId, def] of defsOf(editor))
    if (
      def.listStyleId === styleId ||
      (abstractNumId !== undefined && def.abstractNumId === abstractNumId)
    )
      drawn[numId] = def.levels
  return drawn
}

/** the unsaved numbering with one side's entries for the style in place of the ones it holds */
export function withListStylePending(
  p: PendingNumbering,
  { styleId, abstractNumId }: Pick<ListStyleChange, 'styleId' | 'abstractNumId'>,
  side: StylePending,
): PendingNumbering {
  const def = side.styleDefs[0]
  const made = p.styleDefs.some((s) => s.styleId === styleId)
  return {
    ...p,
    levelEdits: [
      ...p.levelEdits.filter(
        (e) => abstractNumId === undefined || e.abstractNumId !== abstractNumId,
      ),
      ...side.levelEdits,
    ],
    styleDefs: made
      ? p.styleDefs.flatMap((s) => (s.styleId !== styleId ? [s] : def ? [def] : []))
      : def
        ? [...p.styleDefs, def]
        : p.styleDefs,
    styleNums: p.styleNums.map(
      (l) => (l.styleId === styleId && side.styleNums.find((n) => n.numId === l.numId)) || l,
    ),
  }
}

/** the lists of the style drawn with one side's levels again */
export function drawListStyleSide(editor: Editor, side: ListStyleSide): void {
  const store = editor.storage.listNumbering as { defs: Map<string, NumberingDef> }
  const next = new Map(store.defs)
  for (const [numId, levels] of Object.entries(side.drawn)) {
    const def = next.get(numId)
    if (def) next.set(numId, { ...def, levels })
  }
  store.defs = next
  editor.view.dispatch(editor.state.tr)
}

export class ListStyleStep extends Step {
  constructor(
    readonly change: ListStyleChange,
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
    return new ListStyleStep(this.change, this.side === 'redo' ? 'undo' : 'redo')
  }

  map(): Step {
    return this
  }

  toJSON(): { stepType: string } {
    return { stepType: 'listStyle' }
  }
}

const RECORDING = 'listStyleRecording'

/**
 * Modify Style on a list style, as one step of the undo history. `upsertOf`
 * reads the unsaved definition of a style. False when the style cannot be modified.
 */
export function modifyListStyleAsStep(
  ctx: NumberingContext,
  style: ListStyleEntry,
  name: string,
  levels: CustomNumberingLevel[],
  upsertOf: (styleId: string) => StyleUpsert | undefined,
): boolean {
  const { editor, doc } = ctx
  if (!editor || !doc) return false
  const { styleId } = style
  const own = style.inDocument
    ? doc.parsed.listStyles?.find((s) => s.styleId === styleId)
    : undefined
  const abstractNumId =
    own?.numId === undefined ? undefined : doc.parsed.numbering.get(own.numId)?.abstractNumId
  const before: ListStyleSide = {
    name: style.name,
    levels: nineListLevels(style.levels),
    pending: pendingOf(ctx.pendingNumbering, styleId, abstractNumId),
    upsert: upsertOf(styleId) ?? null,
    drawn: drawnOf(editor, styleId, abstractNumId),
  }
  let pending = ctx.pendingNumbering
  let renamed: string | null = null
  const modified = modifyListStyle(
    {
      ...ctx,
      setPendingNumbering: (update) => {
        pending = update(pending)
        ctx.setPendingNumbering(update)
      },
      renameStyle: (id, next) => {
        renamed = next
        ctx.renameStyle?.(id, next)
      },
    },
    style,
    name,
    levels,
  )
  if (!modified) return false
  const after: ListStyleSide = {
    name: style.custom ? name.trim() || style.name : style.name,
    levels,
    pending: pendingOf(pending, styleId, abstractNumId),
    upsert:
      renamed === null
        ? before.upsert
        : { ...(before.upsert ?? { styleId }), styleId, name: renamed },
    drawn: drawnOf(editor, styleId, abstractNumId),
  }
  if (same(before, after)) return true
  const change: ListStyleChange = {
    parsed: doc.parsed,
    styleId,
    ...(abstractNumId !== undefined ? { abstractNumId } : {}),
    undo: before,
    redo: after,
  }
  editor.view.dispatch(
    editor.state.tr
      .step(new ListStyleStep(change, 'redo'))
      .setMeta(RECORDING, true)
      .setMeta(TRACK_IGNORE, true),
  )
  return true
}

/** the sides an undo or redo transaction asks app state to take, in order */
export function replayedListStyles(
  tr: Transaction,
): Array<{ change: ListStyleChange; side: 'undo' | 'redo' }> {
  if (tr.getMeta(RECORDING)) return []
  return tr.steps.flatMap((step) =>
    step instanceof ListStyleStep ? [{ change: step.change, side: step.side }] : [],
  )
}
