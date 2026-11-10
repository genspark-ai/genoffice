/** Notes text typed but not yet written back, with the stored text it started from. */
export interface NotesDraft {
  index: number
  text: string
  /** Stored notes of that slide when typing began; a deck edit that changes them wins over the draft. */
  base: string
  partPath?: string
}

export interface LoadedNotes {
  index: number
  text: string
}

/**
 * The draft after another keystroke: `base` stays what the draft began from,
 * or is the slide's last loaded/flushed text when this keystroke starts one.
 */
export function startNotesDraft(
  prev: NotesDraft | null,
  loaded: LoadedNotes | null,
  index: number,
  text: string,
  partPath: string | undefined,
): NotesDraft {
  const base = prev?.index === index ? prev.base : loaded?.index === index ? loaded.text : ''
  return { index, text, base, partPath }
}

/**
 * The baseline after a flush was read back from the document. A flush for a
 * slide the baseline has moved away from (a late slide-switch write) leaves the
 * current slide's entry alone.
 */
export function notesBaselineAfterFlush(
  loaded: LoadedNotes | null,
  index: number,
  storedText: string,
): LoadedNotes | null {
  if (loaded && loaded.index !== index) return loaded
  return { index, text: storedText }
}

/** Index of the draft's slide in a replaced deck, or -1 when it is gone. */
export function notesDraftIndex(
  draft: Pick<NotesDraft, 'index' | 'partPath'>,
  slides: readonly { partPath?: string }[],
): number {
  if (draft.partPath) return slides.findIndex((s) => s.partPath === draft.partPath)
  return draft.index < slides.length ? draft.index : -1
}

/**
 * The draft after the deck was replaced under it (AI batch, restore): dropped
 * when its slide is gone or the stored notes changed since typing began,
 * re-indexed when the slide moved, otherwise unchanged.
 */
export function retargetNotesDraft(
  draft: NotesDraft,
  slides: readonly { partPath?: string }[],
  storedNotes: string,
): NotesDraft | null {
  const index = notesDraftIndex(draft, slides)
  if (index < 0 || storedNotes !== draft.base) return null
  return index === draft.index ? draft : { ...draft, index }
}
