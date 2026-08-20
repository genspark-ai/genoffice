/**
 * Transaction executor: plan-then-execute with snapshot rollback (semantics
 * modeled on genpptx's production-proven ops executor).
 *
 * - the WHOLE batch is validated first (plan); dryRun returns the plan and
 *   touches nothing;
 * - atomic (default): snapshot → apply all → any failure restores the snapshot,
 *   the deck is untouched and the failure carries a guided error;
 * - per_op: each op applies independently; failures are collected per index and
 *   skipped, successes stay.
 *
 * Snapshots clone the slides deep and the archive entry map shallow — entry
 * buffers are replaced wholesale, never mutated (the same invariant the
 * app's undo history relies on). Electron-free by design: callable from any
 * executor that holds an OpenedPptx, and testable in plain Node.
 */
import {
  ensureCreationId,
  materializeSlide,
  patchSlideXml,
  type OpenedPptx,
  type Slide,
} from '@genoffice/pptx-engine'
import { opUsage } from '../../shared/op-docs'
import {
  GuidedError,
  lookup,
  resolveElement,
  type Op,
  type OpRecord,
  type OpContext,
} from './registry'

export interface TxnRequest {
  /** atomic: all-or-nothing (default). per_op: independent, failures skip. */
  isolation?: 'atomic' | 'per_op'
  /** Validate and return the plan without touching the deck. */
  dryRun?: boolean
  /**
   * Live parsed parts keyed by part path: part-addressed ops mutate THESE
   * objects (parse-time ids stay stable for the caller's render state) and
   * the flush serializes them. Mutations to seeded objects are not
   * snapshot-restored on rollback — seed only single-op transactions whose
   * ops validate before mutating.
   */
  parts?: Map<string, Slide>
  ops: Op[]
}

export interface OpFailure {
  index: number
  op: Op
  /** Guided error: states what failed AND how to fix it. */
  error: string
}

export interface TxnResult {
  applied: boolean
  dryRun?: boolean
  records?: OpRecord[]
  /** dry-run: the validated plan, one line per op */
  plan?: string[]
  failures?: OpFailure[]
}

interface Snapshot {
  slides: OpenedPptx['deck']['slides']
  entries: Map<string, unknown>
  size: OpenedPptx['deck']['size']
}

function takeSnapshot(opened: OpenedPptx): Snapshot {
  return {
    slides: structuredClone(opened.deck.slides),
    entries: new Map(opened.archive.entries as Map<string, unknown>),
    size: { ...opened.deck.size },
  }
}

function restoreSnapshot(opened: OpenedPptx, snap: Snapshot): void {
  opened.deck.slides = snap.slides
  opened.deck.size = snap.size
  const entries = opened.archive.entries as Map<string, unknown>
  entries.clear()
  for (const [k, v] of snap.entries) entries.set(k, v)
}

function planLine(i: number, op: Op): string {
  const t = op.target ? ` s${op.target.slide}${op.target.el ? `/${op.target.el}` : ''}` : ''
  return `[${i}] ${op.op}${t}`
}

function message(e: unknown): string {
  return e instanceof Error ? e.message : String(e)
}

/**
 * Progressive identity hardening for foreign decks: any element an op just
 * rewrote gets an a16:creationId minted into its bytes if it lacks one
 * (idempotent), so its durable id upgrades from the cNvPr fallback to a GUID
 * that survives renumbering. The executor is the single write gateway, which
 * makes this the one choke point; untouched elements keep their exact bytes.
 * Deleted targets and group children (no top-level byte slice) resolve-fail
 * and are skipped. Part-addressed targets harden too — their bytes flush with
 * the part at transaction end.
 */
function hardenTargetIdentity(op: Op, ctx: OpContext): void {
  if (typeof op.target?.el !== 'string') return
  try {
    ensureCreationId(resolveElement(ctx, op, { allowPart: true }).el)
  } catch {
    /* nothing to harden */
  }
}

/** Guided errors teach the exact signature on first contact: a failing op's
    message carries its one-line usage, so the model needs no resident field
    reference — the vocabulary in the tool description plus this is enough. */
function withUsage(op: Op, msg: string): string {
  const usage = op && typeof op.op === 'string' ? opUsage(op.op) : undefined
  return usage ? `${msg}\n${usage}` : msg
}

/** Serialize part-addressed (master/layout) edits back to entries — once per
    transaction — then re-materialize every deck slide: slides resolve
    inherited styles against the chrome, so their models are stale after it
    changes. Rolled-back transactions never reach this (mutations lived only
    in the discarded parse cache). */
function flushTouchedParts(opened: OpenedPptx, ctx: OpContext): void {
  if (ctx.touchedParts.size === 0) return
  const entries = opened.archive.entries as Map<string, unknown>
  for (const part of ctx.touchedParts) {
    const slide = ctx.parts.get(part)
    if (slide) entries.set(part, Buffer.from(patchSlideXml(slide), 'utf8'))
  }
  for (let i = 0; i < opened.deck.slides.length; i++) materializeSlide(opened, i)
}

export function runTxn(opened: OpenedPptx, req: TxnRequest): TxnResult {
  if (!Array.isArray(req.ops) || req.ops.length === 0) {
    return {
      applied: false,
      failures: [{ index: 0, op: { op: '(none)' }, error: 'ops must be a non-empty array.' }],
    }
  }
  const ctx: OpContext = { opened, parts: new Map(req.parts ?? []), touchedParts: new Set() }
  const isolation = req.isolation ?? 'atomic'

  // ── plan: validate everything up front (against the pre-transaction state) ──
  const plan: string[] = []
  const planFailures: OpFailure[] = []
  req.ops.forEach((op, i) => {
    try {
      if (typeof op !== 'object' || op === null || typeof op.op !== 'string') {
        throw new GuidedError('each op must be an object with an "op" name field.')
      }
      lookup(op.op).validate(op, ctx)
      plan.push(planLine(i, op))
    } catch (e) {
      planFailures.push({ index: i, op, error: withUsage(op, message(e)) })
    }
  })

  if (req.dryRun) {
    return {
      applied: false,
      dryRun: true,
      plan,
      ...(planFailures.length ? { failures: planFailures } : {}),
    }
  }
  if (isolation === 'atomic' && planFailures.length) {
    return {
      applied: false,
      failures: planFailures.map((f) => ({
        ...f,
        error: `${f.error} Nothing was applied (atomic) — fix this op and resend the whole transaction.`,
      })),
    }
  }

  // ── execute ──
  // Parts resolved during validation were not mutated; only apply-phase
  // resolutions may dirty them, so the flush set starts clean here.
  ctx.touchedParts.clear()
  if (isolation === 'atomic') {
    const snap = takeSnapshot(opened)
    const records: OpRecord[] = []
    for (const [i, op] of req.ops.entries()) {
      try {
        records.push(lookup(op.op).apply(op, ctx))
        hardenTargetIdentity(op, ctx)
      } catch (e) {
        restoreSnapshot(opened, snap)
        return {
          applied: false,
          failures: [
            {
              index: i,
              op,
              error: withUsage(
                op,
                `${message(e)} Nothing was applied (atomic) — fix this op and resend the whole transaction.`,
              ),
            },
          ],
        }
      }
    }
    flushTouchedParts(opened, ctx)
    return { applied: true, records }
  }

  const records: OpRecord[] = []
  const failures: OpFailure[] = [...planFailures]
  const failedIdx = new Set(planFailures.map((f) => f.index))
  for (const [i, op] of req.ops.entries()) {
    if (failedIdx.has(i)) continue
    try {
      records.push(lookup(op.op).apply(op, ctx))
      hardenTargetIdentity(op, ctx)
    } catch (e) {
      failures.push({ index: i, op, error: withUsage(op, message(e)) })
    }
  }
  if (records.length > 0) flushTouchedParts(opened, ctx)
  failures.sort((a, b) => a.index - b.index)
  return {
    applied: records.length > 0,
    records,
    ...(failures.length ? { failures } : {}),
  }
}
