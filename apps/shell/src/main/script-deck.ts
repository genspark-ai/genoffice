import { renameSync, writeFileSync } from 'node:fs'
import type { ScriptRunOutcome } from '@genoffice/agent-core'
import type { ScriptApplyResult } from '../shared/feature-channels'
import { readInput } from '@genoffice/cli/headless-docs'
import {
  applyOps,
  auditDeckBytes,
  deckAuditDetail,
  describeDeck,
  openDeck,
  saveDeck,
  type Op,
} from '@genoffice/cli/headless-formats'
import { ScriptRunError, type SandboxOutcome, type DeckSnapshot } from './script-shared'

/**
 * The deck half of the script kind — one dialect, not a second one: the script
 * reads what describeDeck reports (real EMU geometry and text of every element,
 * so no read call first) and proposes the same op vocabulary the slides app and
 * its assistant already use (setText, addElement, addPicture, ...).
 *
 * The engine is runTxn in atomic isolation: a refused op rolls the whole
 * transaction back, so nothing is written; an apply saves and then reports the
 * layout audit (overlap, out-of-bounds, text overflow) with the result, the same
 * audit the slides app shows.
 */

export const deckHost = {
  async snapshot(documentPath: string): Promise<DeckSnapshot> {
    const opened = await openDeck(readInput(documentPath))
    return { format: 'deck', deck: describeDeck(opened) }
  },
  async plan({
    outcome,
    documentPath,
  }: {
    outcome: SandboxOutcome
    documentPath: string
  }): Promise<ScriptRunOutcome> {
    if (!outcome.ops.length) {
      return { applied: 0, plan: [], failures: [], log: outcome.log, ms: outcome.ms }
    }
    const opened = await openDeck(readInput(documentPath))
    // the sandbox already checked every entry is an object with a string op name
    const txn = applyOps(opened, outcome.ops as Op[], { isolation: 'atomic', dryRun: true })
    return {
      applied: txn.plan?.length ?? 0,
      plan: txn.plan ?? [],
      failures: (txn.failures ?? []).map(
        (failure) => failure.index + ' ' + failure.op.op + ': ' + failure.error,
      ),
      log: outcome.log,
      ms: outcome.ms,
    }
  },
  async apply({
    outcome,
    documentPath,
  }: {
    outcome: SandboxOutcome
    documentPath: string
  }): Promise<ScriptApplyResult> {
    if (!outcome.ops.length) {
      throw new ScriptRunError(
        'the script proposed nothing, so there is nothing to apply',
        'check what the script reads: $.deck.slides() and $.deck.propose([...])',
      )
    }
    const opened = await openDeck(readInput(documentPath))
    const txn = applyOps(opened, outcome.ops as Op[], { isolation: 'atomic' })
    if (!txn.applied || txn.failures?.length) {
      const first = txn.failures?.[0]
      throw new ScriptRunError(
        'the engine refused ' +
          (txn.failures?.length ?? 1) +
          ' operation(s), so nothing was written',
        first ? first.index + ' ' + first.op.op + ': ' + first.error : undefined,
      )
    }
    const bytes = await saveDeck(opened)
    // replace in one step: a reader never sees a half-written file
    const temp = documentPath + '.genoffice-apply-' + process.pid
    writeFileSync(temp, bytes)
    renameSync(temp, documentPath)
    return {
      applied: txn.records?.length ?? outcome.ops.length,
      plan: [...planLines(txn), ...(await deckAuditLines(bytes))],
      failures: [],
      log: outcome.log,
      ms: outcome.ms,
      written: documentPath,
      bytes: bytes.length,
    }
  },
}

function planLines(txn: { records?: { op: { op: string } }[] }): string[] {
  return (txn.records ?? []).map((record, index) => index + ': ' + record.op.op)
}

/** The layout audit is advisory: a failure to audit never fails the apply. */
async function deckAuditLines(bytes: Uint8Array): Promise<string[]> {
  try {
    const detail = deckAuditDetail(await auditDeckBytes(bytes))
    return [
      'layout audit: ' +
        detail.counts.error +
        ' error(s), ' +
        detail.counts.warning +
        ' warning(s)',
      ...detail.issues
        .slice(0, 5)
        .map((issue) => issue.id + ' ' + issue.path + ' ' + issue.code + ': ' + issue.message),
    ]
  } catch (err) {
    return ['layout audit unavailable: ' + (err instanceof Error ? err.message : String(err))]
  }
}
