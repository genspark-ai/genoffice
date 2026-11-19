import type { TxnResult } from '@genoffice/pptx-ops'
import { classifyOpError, opSuggestion } from '../op-errors'
import { CliError, EXIT } from '../result'

/** Guided op errors go back verbatim: the agent fixes the op and retries. `offset` maps single-op batches back to the caller's list. */
export function txnFailure(r: TxnResult, offset = 0): CliError {
  const failures = (r.failures ?? []).map((f) =>
    classifyOpError(f.index + offset, f.op.op, f.error),
  )
  const first = failures[0]
  return new CliError(
    EXIT.usage,
    first ? `op ${first.index} (${first.op}) rejected: ${first.error}` : 'no ops were applied',
    { failures },
    first ? { reason: first.reason, suggestion: opSuggestion(first, 'slides') } : {},
  )
}

export function txnDetail(r: TxnResult): Record<string, unknown> {
  const records = (r.records ?? []).map((rec) => ({
    op: rec.op.op,
    ...(rec.slideId ? { slide: rec.slideId } : {}),
    ...(rec.created?.length ? { created: rec.created } : {}),
  }))
  return {
    applied: r.applied,
    ops: records.length,
    records,
    ...(r.plan ? { plan: r.plan } : {}),
    ...(r.failures?.length
      ? { failures: r.failures.map((f) => classifyOpError(f.index, f.op.op, f.error)) }
      : {}),
  }
}
