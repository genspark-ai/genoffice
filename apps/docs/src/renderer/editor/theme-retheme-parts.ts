import type { NoteInfo, ParsedDocFull } from '@genoffice/docx-engine'
import { hfVariantsFromParsed } from '../doc-state'
import type { HfDocStrips, HfParsedParts } from '../hf-sections'

/**
 * Headers, footers and notes hold theme colours and faces resolved when the
 * file was read, as the body does (theme-retheme.ts). They live in App state
 * as plain data, so a theme pick is put on them value by value: where the live
 * value still is what the file read as, it takes what the file reads as under
 * the picked theme. Nothing here marks a part as edited.
 */

/** one reading of the file's parts outside the body */
export interface PartsReading {
  strips: HfDocStrips
  /** the parts as parsed: what sections with their own references draw from */
  parsed: HfParsedParts
  footnotes: NoteInfo[]
  endnotes: NoteInfo[]
}

type Strip = NonNullable<HfDocStrips['header']>

const docStripOf = (
  text: string | null | undefined,
  pageNumber: boolean | undefined,
  paras: Strip['paras'],
): Strip | null =>
  text || pageNumber || paras?.length
    ? { text: text ?? '', pageNumber: pageNumber ?? false, paras: paras ?? undefined }
    : null

export function partsReading(parsed: ParsedDocFull): PartsReading {
  return {
    strips: {
      header: docStripOf(
        parsed.headerText,
        parsed.headerHasPageNumber,
        parsed.headerParas ?? undefined,
      ),
      footer: docStripOf(
        parsed.footerText,
        parsed.footerHasPageNumber,
        parsed.footerParas ?? undefined,
      ),
      variants: hfVariantsFromParsed(parsed),
    },
    parsed: {
      headerImages: parsed.headerImages,
      footerImages: parsed.footerImages,
      headerFirst: parsed.headerFirst,
      footerFirst: parsed.footerFirst,
      headerEven: parsed.headerEven,
      footerEven: parsed.footerEven,
      hfParts: parsed.hfParts,
    },
    footnotes: parsed.footnotes,
    endnotes: parsed.endnotes,
  }
}

const isPlain = (v: unknown): v is Record<string, unknown> =>
  typeof v === 'object' && v !== null && Object.getPrototypeOf(v) === Object.prototype

const key = (v: unknown): string => JSON.stringify(v) ?? 'undefined'

/** `live` with every value that still equals `from` taken to `to`; the same object when nothing moved */
export function rethemedValue<T>(live: T, from: unknown, to: unknown): T {
  if (from === to) return live
  if (isPlain(live) && isPlain(from) && isPlain(to)) {
    let next: Record<string, unknown> | null = null
    for (const k of new Set([...Object.keys(from), ...Object.keys(to)])) {
      const value = rethemedValue(live[k], from[k], to[k])
      if (value === live[k]) continue
      next ??= { ...live }
      if (value === undefined) delete next[k]
      else next[k] = value
    }
    return (next ?? live) as T
  }
  if (Array.isArray(live) && Array.isArray(from) && Array.isArray(to)) {
    if (live.length !== from.length || from.length !== to.length) return live
    let next: unknown[] | null = null
    live.forEach((item, i) => {
      const value = rethemedValue(item, from[i], to[i])
      if (value === item) return
      next ??= [...live]
      next[i] = value
    })
    return (next ?? live) as T
  }
  const was = key(from)
  return was !== key(to) && key(live) === was ? (to as T) : live
}

/** notes pair by id: one added or removed in between leaves the others in place */
export function rethemedNotes(live: NoteInfo[], from: NoteInfo[], to: NoteInfo[]): NoteInfo[] {
  const was = new Map(from.map((n) => [n.id, n]))
  const now = new Map(to.map((n) => [n.id, n]))
  let moved = false
  const next = live.map((note) => {
    if (note.created || note.origin || !was.has(note.id) || !now.has(note.id)) return note
    const value = rethemedValue(note, was.get(note.id), now.get(note.id))
    if (value !== note) moved = true
    return value
  })
  return moved ? next : live
}
