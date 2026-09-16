import { basename } from 'node:path'
import type { CommandResult } from './result'

export type IssueLevel = 'error' | 'warning' | 'info'

/** One finding of `sheet check` / `docs check`; `path` is the human locator, the typed fields the machine one. */
export interface CheckIssue {
  id: string
  code: string
  level: IssueLevel
  path: string
  message: string
  context?: string
  sheet?: string
  cell?: string
  range?: string
  blockIndex?: number
  part?: string
  /** an op for `apply` that fixes the finding */
  suggest?: Record<string, unknown>
}

export type IssueDraft = Omit<CheckIssue, 'id'>

/** Text an agent left behind: template keys, TODO markers, dummy text. */
export const PLACEHOLDER = /\{\{[^}]*\}\}|\bXX+%?\b|\blorem ipsum\b|\bTBD\b|\bTODO\b|\[insert\b/i

export const MAX_ISSUES = 200

const LEVEL_ORDER: Record<IssueLevel, number> = { error: 0, warning: 1, info: 2 }

export function checkResult(
  file: string,
  drafts: IssueDraft[],
  checks: readonly string[],
  notes?: string,
): CommandResult {
  const sorted = [...drafts].sort((a, b) => LEVEL_ORDER[a.level] - LEVEL_ORDER[b.level])
  const truncated = sorted.length > MAX_ISSUES
  const issues: CheckIssue[] = sorted
    .slice(0, MAX_ISSUES)
    .map((d, i) => ({ id: `${d.level[0]!.toUpperCase()}${i + 1}`, ...d }))
  const counts = { error: 0, warning: 0, info: 0 }
  for (const d of drafts) counts[d.level]++
  const name = basename(file)
  const summary =
    drafts.length === 0
      ? `${name}: no issues found (${checks.length} checks)`
      : `${name}: ${drafts.length} issue(s): ${counts.error} error(s), ${counts.warning} warning(s), ${counts.info} note(s)`
  const detail: Record<string, unknown> = { issues, counts, checks }
  if (truncated) detail.truncated = `${drafts.length - MAX_ISSUES} more issue(s) not listed`
  if (notes) detail.notes = notes
  return { summary, detail }
}

/** Trim a formula or text for the context field. */
export function excerpt(text: string, max = 120): string {
  const flat = text.replace(/\s+/g, ' ').trim()
  return flat.length > max ? `${flat.slice(0, max - 1)}…` : flat
}

/**
 * Approximate rendered width of grid text in px at `sizePt`, tuned to Calibri (7 px per digit
 * at 11 pt, Excel's column unit). Good to about one character; the checks leave slack for that.
 */
export function textWidthPx(text: string, sizePt = 11): number {
  let units = 0
  for (const ch of text) {
    const code = ch.codePointAt(0)!
    if (code >= 0x2e80) units += 11
    else if (ch >= '0' && ch <= '9') units += 7
    else if (ch >= 'A' && ch <= 'Z') units += 8
    else if (ch >= 'a' && ch <= 'z') units += 6.2
    else if (ch === ' ') units += 3
    else if (".,:;-/'|!".includes(ch)) units += 3.5
    else units += 6.5
  }
  return (units * sizePt) / 11
}
