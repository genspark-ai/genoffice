/** Shared bits of the Word-style comment card: the Comments panel (React) and the margin bubbles (DOM) read the same identity helpers and class names */
import type { CommentInfo } from '@genoffice/docx-engine'
import { REVISION_INK_COUNT } from './revision-filters'

export const inkVar = (idx: number) =>
  idx === 0 ? 'var(--docs-paper-rev-ink)' : `var(--docs-paper-rev-ink-${idx})`

/** stable avatar colour per author name, drawn from the revision ink cycle */
export function authorInk(name: string): string {
  let h = 0
  for (let i = 0; i < name.length; i++) h = (h * 31 + name.charCodeAt(i)) >>> 0
  return inkVar(h % REVISION_INK_COUNT)
}

export function authorInitials(c: Pick<CommentInfo, 'author' | 'initials'>): string {
  if (c.initials) return c.initials.slice(0, 2)
  const words = (c.author ?? '').trim().split(/\s+/).filter(Boolean)
  if (words.length === 0) return '?'
  const raw = words.length === 1 ? words[0].slice(0, 2) : words[0][0] + words[words.length - 1][0]
  return raw.toUpperCase()
}

export function absoluteDate(iso?: string | null): string {
  if (!iso) return ''
  const d = new Date(iso)
  return Number.isNaN(d.getTime()) ? '' : d.toLocaleString()
}
