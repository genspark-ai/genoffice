import { readFileSync, renameSync, writeFileSync } from 'node:fs'
import type { ScriptRunOutcome } from '@genoffice/agent-core'
import type { ScriptApplyResult } from '../shared/feature-channels'
import { readInput } from '@genoffice/cli/headless-docs'
import {
  ScriptRunError,
  TEXT_SNAPSHOT_BYTES,
  type SandboxOutcome,
  type TextSnapshot,
} from './script-shared'

/**
 * The text half of the script kind — markdown and html. The document is its own
 * engine: the script reads the whole text ($.text.get(), $.text.lines()) and
 * returns the new full text, and the host makes one whole-document edit. That
 * keeps the diff reviewable line by line and the write atomic; there is no op
 * vocabulary and no partial patch.
 */

export const textHost = {
  async snapshot(documentPath: string): Promise<TextSnapshot> {
    return textSnapshot(Buffer.from(readInput(documentPath)))
  },
  plan({
    outcome,
    documentPath,
  }: {
    outcome: SandboxOutcome
    documentPath: string
  }): ScriptRunOutcome {
    if (outcome.text === undefined) {
      return { applied: 0, plan: [], failures: [], log: outcome.log, ms: outcome.ms }
    }
    const current = readFileSync(documentPath, 'utf8')
    if (outcome.text === current) {
      return {
        applied: 0,
        plan: ['the script returned the text unchanged'],
        failures: [],
        log: outcome.log,
        ms: outcome.ms,
      }
    }
    return {
      applied: 1,
      plan: textDiffLines(current, outcome.text),
      failures: [],
      log: outcome.log,
      ms: outcome.ms,
    }
  },
  apply({
    outcome,
    documentPath,
  }: {
    outcome: SandboxOutcome
    documentPath: string
  }): ScriptApplyResult {
    if (outcome.text === undefined) {
      throw new ScriptRunError(
        'the script returned no text to write',
        'return the new full document text (a string, or { text }) from the function',
      )
    }
    const current = readFileSync(documentPath, 'utf8')
    if (outcome.text === current) {
      throw new ScriptRunError(
        'the script returned the text unchanged, so there is nothing to apply',
        'return the edited document text',
      )
    }
    const bytes = Buffer.from(outcome.text, 'utf8')
    if (bytes.length > TEXT_SNAPSHOT_BYTES) {
      throw new ScriptRunError(
        'the returned text is larger than the ' +
          Math.round(TEXT_SNAPSHOT_BYTES / 1024 / 1024) +
          'MB cap',
        'return the document, not an expansion of it',
      )
    }
    // replace in one step: a reader never sees a half-written file
    const temp = documentPath + '.genoffice-apply-' + process.pid
    writeFileSync(temp, bytes)
    renameSync(temp, documentPath)
    return {
      applied: 1,
      plan: textDiffLines(current, outcome.text),
      failures: [],
      log: outcome.log,
      ms: outcome.ms,
      written: documentPath,
      bytes: bytes.length,
    }
  },
}

function textSnapshot(bytes: Buffer): TextSnapshot {
  if (bytes.length > TEXT_SNAPSHOT_BYTES) {
    throw new ScriptRunError(
      'this file is ' +
        Math.round(bytes.length / 1024 / 1024) +
        'MB; the text API reads at most ' +
        Math.round(TEXT_SNAPSHOT_BYTES / 1024 / 1024) +
        'MB',
      'run it against a smaller document',
    )
  }
  return { format: 'text', text: bytes.toString('utf8') }
}

const DIFF_PREVIEW_CHARS = 80
const DIFF_PAIRS_SHOWN = 5

/**
 * The diff summary a text run reports: the changed line span and the first few
 * old/new pairs. A common prefix/suffix trim keeps it O(lines) — enough for a
 * reviewable summary, never a full LCS.
 */
export function textDiffLines(before: string, after: string): string[] {
  const a = before.split('\n')
  const b = after.split('\n')
  let start = 0
  while (start < a.length && start < b.length && a[start] === b[start]) start++
  let endA = a.length
  let endB = b.length
  while (endA > start && endB > start && a[endA - 1] === b[endB - 1]) {
    endA--
    endB--
  }
  const removed = endA - start
  const added = endB - start
  const head =
    removed === 0
      ? added + ' line(s) inserted after line ' + start
      : added === 0
        ? removed + ' line(s) deleted at line ' + (start + 1)
        : removed + ' line(s) replaced by ' + added + ' at line ' + (start + 1)
  const lines = [head]
  let shown = 0
  for (let i = start; i < endA && shown < DIFF_PAIRS_SHOWN; i++, shown++) {
    lines.push('- ' + clip(a[i] ?? '', DIFF_PREVIEW_CHARS))
  }
  for (let j = start; j < endB && shown < DIFF_PAIRS_SHOWN * 2; j++, shown++) {
    lines.push('+ ' + clip(b[j] ?? '', DIFF_PREVIEW_CHARS))
  }
  const rest = removed + added - shown
  if (rest > 0) lines.push('…(+' + rest + ' more changed line(s))')
  return lines
}

function clip(text: string, max: number): string {
  return text.length > max ? text.slice(0, max) + '…' : text
}
