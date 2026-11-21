import type { Node as PMNode } from '@tiptap/pm/model'
import type { Transaction } from '@tiptap/pm/state'
import { Step, StepResult } from '@tiptap/pm/transform'
import type { WrittenFontConversion } from '@genoffice/docx-engine'

/**
 * Convert Permanently reaches parts the editor does not hold (styles, headers,
 * notes, numbering, theme), so the conversion is written at the save. Its
 * place in the undo history is a step beside the run changes, so Undo and Redo
 * take the save's share along with the editor's.
 */
export interface FontConversions {
  /** missing font -> its stand-in, to write at the next save */
  pending: Record<string, string>
  /** what earlier saves of this session wrote, oldest first */
  written: readonly WrittenFontConversion[]
}

type FontMap = Readonly<Record<string, string>>

const without = (fonts: FontMap, gone: FontMap): Record<string, string> =>
  Object.fromEntries(Object.entries(fonts).filter(([name]) => !(name in gone)))

const holds = (fonts: FontMap, all: FontMap): boolean =>
  Object.entries(all).every(([name, to]) => fonts[name] === to)

const settled = (state: FontConversions): FontConversions | null =>
  Object.keys(state.pending).length === 0 && state.written.length === 0 ? null : state

function replaceLast<T>(list: readonly T[], match: (item: T) => boolean, next: (item: T) => T) {
  let at = list.length - 1
  while (at >= 0 && !match(list[at]!)) at--
  return at < 0 ? null : list.map((item, i) => (i === at ? next(item) : item))
}

/** the state after Convert Permanently, or after a redo of it */
export function withFontsConverted(
  state: FontConversions | null,
  fonts: FontMap,
): FontConversions | null {
  const { pending, written } = state ?? { pending: {}, written: [] }
  // an undo no save has written yet: the file still holds the conversion
  const kept = replaceLast(
    written,
    (w) => holds(w.fonts, fonts) && !Object.keys(fonts).some((name) => name in w.kept),
    (w) => {
      const next = { ...w.kept, ...fonts }
      return { ...w, kept: holds(next, w.fonts) ? w.fonts : next }
    },
  )
  return settled(kept ? { pending, written: kept } : { pending: { ...pending, ...fonts }, written })
}

/** the state after an undo of Convert Permanently */
export function withFontsReverted(
  state: FontConversions | null,
  fonts: FontMap,
): FontConversions | null {
  if (!state) return null
  const { pending, written } = state
  if (Object.keys(fonts).some((name) => name in pending))
    return settled({ pending: without(pending, fonts), written })
  const put = replaceLast(
    written,
    (w) => holds(w.kept, fonts),
    (w) => ({ ...w, kept: without(w.kept, fonts) }),
  )
  return put ? { pending, written: put } : state
}

/** something for the save to write: a conversion, or an undone one to put back */
export function fontsToWrite(state: FontConversions | null): boolean {
  return (
    !!state &&
    (Object.keys(state.pending).length > 0 || state.written.some((w) => w.kept !== w.fonts))
  )
}

/** fonts an undo brought back into the editor that the file does not name again until the save */
export function fontsBeingRestored(state: FontConversions | null): string[] {
  return (state?.written ?? []).flatMap((w) =>
    Object.keys(w.fonts).filter((name) => !(name in w.kept)),
  )
}

export class FontConversionStep extends Step {
  constructor(
    readonly fonts: FontMap,
    /** false: applying this step takes the conversion back */
    readonly converting: boolean,
  ) {
    super()
  }

  // a new document object over the same content, as a page setup step: fonts
  // only styles or headers name still make an edit to undo
  apply(doc: PMNode): StepResult {
    return StepResult.ok(doc.type.create(doc.attrs, doc.content, doc.marks))
  }

  invert(): Step {
    return new FontConversionStep(this.fonts, !this.converting)
  }

  map(): Step {
    return this
  }

  toJSON(): { stepType: string } {
    return { stepType: 'fontConversion' }
  }
}

export const FONT_CONVERSION_RECORDING = 'fontConversionRecording'

/** the conversions an undo or redo transaction puts in place or takes back, in order */
export function replayedFontConversions(tr: Transaction): FontConversionStep[] {
  if (tr.getMeta(FONT_CONVERSION_RECORDING)) return []
  return tr.steps.filter((step): step is FontConversionStep => step instanceof FontConversionStep)
}
