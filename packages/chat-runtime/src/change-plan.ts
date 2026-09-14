/**
 * ChangePlan normalisation.
 *
 * XLSX's `propose_operations` tool emits a `WorkbookOp[]`; docs/pdf/slides
 * can opt in by emitting the same shape (or a free-form shape that we wrap).
 * This module is the single normalisation point so the timeline + summary
 * components can render any plan without per-app branching.
 */

import type { ChatChangePlan, ChatChangePlanOp } from './types.js'

let counter = 0
function nextId(prefix: string): string {
  counter += 1
  return `${prefix}-${Date.now().toString(36)}-${counter.toString(36)}`
}

export interface NormalizeChangePlanInput {
  /** Stable id within the originating tool call (e.g. `plan-123`). */
  id?: string
  app: ChatChangePlan['app']
  title: string
  /** Optional human-friendly summary, used as a fallback when the ops don't describe themselves. */
  summary?: string
  /** Operations: either a flat array (XLSX-style) or pre-shaped `ChatChangePlanOp[]`. */
  ops: ChatChangePlanOp[] | unknown[]
  warnings?: string[]
  requireConfirm?: boolean
}

/**
 * Normalise any change-plan shape into the canonical `ChatChangePlan`.
 *
 * - When `ops` already contains `ChatChangePlanOp` entries (objects with
 *   `kind`), they're used as-is.
 * - When `ops` is a flat array of `WorkbookOp` (no `kind`), it is wrapped
 *   into `{ kind: 'workbook', ops }` so the renderer can pick the right
 *   preview renderer.
 */
export function normalizeChangePlan(input: NormalizeChangePlanInput): ChatChangePlan {
  const ops: ChatChangePlanOp[] = []
  for (const op of input.ops ?? []) {
    if (op && typeof op === 'object' && 'kind' in (op as Record<string, unknown>)) {
      ops.push(op as ChatChangePlanOp)
    } else {
      ops.push({ kind: 'workbook', ops: [op] } as ChatChangePlanOp)
    }
  }
  return {
    id: input.id ?? nextId('plan'),
    app: input.app,
    title: input.title,
    summary: input.summary ?? deriveSummary(ops),
    ops,
    warnings: input.warnings,
    requireConfirm: input.requireConfirm ?? false,
    createdAt: Date.now(),
  }
}

function deriveSummary(ops: ChatChangePlanOp[]): string {
  if (ops.length === 0) return 'No operations'
  const counts = new Map<string, number>()
  for (const op of ops) {
    const kind = op.kind ?? 'freeform'
    counts.set(kind, (counts.get(kind) ?? 0) + (op.ops?.length ?? 1))
  }
  return [...counts.entries()].map(([k, n]) => `${n} ${k}`).join(', ')
}

/**
 * Render a short bullet list (one per `ChatChangePlanOp`) for the default
 * preview slot. Apps that want a richer preview can pass a `previewRenderer`
 * into `summarizeChangePlan` instead.
 */
export function summarizeChangePlan(
  plan: ChatChangePlan,
  previewRenderer?: (op: ChatChangePlanOp) => string,
): string[] {
  const out: string[] = []
  for (const op of plan.ops) {
    if (previewRenderer) {
      out.push(previewRenderer(op))
      continue
    }
    const inner = op.ops?.length ?? 0
    out.push(`[${op.kind}] ${op.description ?? `${inner} op${inner === 1 ? '' : 's'}`}`)
  }
  return out
}
