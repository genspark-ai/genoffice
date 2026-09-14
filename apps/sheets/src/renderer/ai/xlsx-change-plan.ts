/**
 * XLSX → cross-app `ChatChangePlan` adapter.
 *
 * The XLSX skill emits a rich `ChangePlan` (cellChanges / sheetRenames /
 * structuralChanges / formatChanges) that the timeline + summary
 * components from `@genoffice/ui` already render via the unified
 * `ChatChangePlan` model. This module is the single normalisation point
 * so the renderer never has to special-case the XLSX shape.
 */

import {
  normalizeChangePlan,
  summarizeChangePlan,
  type ChatChangePlan,
  type ChatChangePlanOp,
} from '@genoffice/chat-runtime'
import type { ChangePlan } from '@genoffice/xlsx-gateway/domain/workbook.types'
import type { WorkbookOperation } from '@genoffice/xlsx-gateway/domain/workbook-dsl'

let counter = 0
function nextPlanId(transactionId?: string): string {
  counter += 1
  return transactionId
    ? `xlsx-${transactionId}-${counter.toString(36)}`
    : `xlsx-${Date.now().toString(36)}-${counter.toString(36)}`
}

/**
 * Convert a validated XLSX `ChangePlan` (the post-validation shape used by
 * the XLSX gateway) into the cross-app `ChatChangePlan`.
 */
export function toChatChangePlan(plan: ChangePlan, opts: { title?: string; summary?: string } = {}): ChatChangePlan {
  const ops: ChatChangePlanOp[] = []

  if (plan.structuralChanges.length > 0) {
    ops.push({
      kind: 'workbook',
      ops: plan.structuralChanges.map(c => ({ ...c })),
      description: plan.structuralChanges.map(c => c.label).join('; '),
    })
  }

  if (plan.formatChanges.length > 0) {
    ops.push({
      kind: 'workbook',
      ops: plan.formatChanges.map(c => ({ ...c })),
      description: plan.formatChanges.map(c => c.label).join('; '),
    })
  }

  if (plan.cellChanges.length > 0) {
    const shown = plan.cellChanges.slice(0, 20)
    const rest = plan.cellChanges.length - shown.length
    ops.push({
      kind: 'workbook',
      ops: plan.cellChanges.map(c => ({ ...c })),
      description:
        shown
          .map(c => `${c.address}: ${c.before ?? ''} → ${c.after ?? ''}`)
          .join('; ') + (rest > 0 ? `; …${rest} more` : ''),
    })
  }

  if (plan.sheetRenames.length > 0) {
    ops.push({
      kind: 'workbook',
      ops: plan.sheetRenames.map(r => ({ ...r })),
      description: plan.sheetRenames.map(r => `${r.before} → ${r.after}`).join('; '),
    })
  }

  const summary =
    opts.summary ??
    ((ops.map(o => o.description ?? '').filter(Boolean).join(' | ')) || '(no changes)')

  return normalizeChangePlan({
    id: nextPlanId(plan.transactionId),
    app: 'sheets',
    title: opts.title ?? `Workbook change (${plan.transactionId})`,
    summary,
    ops,
    warnings: plan.warnings ? [...plan.warnings] : [],
    requireConfirm: false,
  })
}

/**
 * Convert a raw `WorkbookOperation[]` (the shape the model emits before
 * XLSX-side validation) into a `ChatChangePlan`. Used by tests and by the
 * preview-only fallback path that does not run through the gateway.
 */
export function toChangePlanFromOps(
  operations: readonly WorkbookOperation[],
  opts: { title?: string; summary?: string; warnings?: readonly string[] } = {},
): ChatChangePlan {
  return normalizeChangePlan({
    id: nextPlanId(),
    app: 'sheets',
    title: opts.title ?? 'Workbook operations',
    summary: opts.summary ?? `${operations.length} operation${operations.length === 1 ? '' : 's'}`,
    ops: operations.map(op => ({ ...op })),
    warnings: opts.warnings ? [...opts.warnings] : [],
    requireConfirm: false,
  })
}

/**
 * Default per-op renderer used by `<AiChangeSummary>` when the host
 * doesn't pass a custom `previewRenderer`. Keeps the XLSX preview
 * readable without re-importing the gateway internals.
 */
export function defaultXlsxPreviewRenderer(op: ChatChangePlanOp): string {
  if (op.description) return `[workbook] ${op.description}`
  if (Array.isArray(op.ops)) return `[workbook] ${op.ops.length} op${op.ops.length === 1 ? '' : 's'}`
  return '[workbook]'
}

export { summarizeChangePlan }
