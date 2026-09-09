/**
 * Canonical edit-op registry: the one write path into the pending-edit state
 * (EditSnapshot), mirroring apps/slides/src/main/ops. Ops are pure reducers that
 * read and return only the buckets in `touches`, so the React host can run each
 * bucket through its own functional setState while tests reduce a whole snapshot.
 * Async display metadata (validated text bounds, ghost PNGs) is not an edit.
 */
import type { EditSnapshot } from '../edit-state'
import { opUsage } from '../../shared/op-docs'

export class GuidedError extends Error {}

export type Bucket = keyof EditSnapshot

export interface Op {
  op: string
  /** Minted at plan time for additive ops when absent */
  id?: string
  [key: string]: unknown
}

export interface OpContext {
  readOnly: boolean
  pageCount: number
  deleted: ReadonlySet<number>
  /** `pageIndex:rectKey` of existing images already claimed by a pending edit */
  claimedImages: ReadonlySet<string>
}

export interface OpRecord {
  op: Op
  created?: string[]
}

export interface OpDef {
  name: string
  touches: Bucket[]
  additive?: boolean
  /** Throws GuidedError: what is wrong and what to do */
  validate(op: Op, ctx: OpContext): void
  /** Context the ops after this one in the batch plan against (a page it deletes, an image it claims) */
  advance?(op: Op, ctx: OpContext): Partial<OpContext>
  apply(op: Op, state: EditSnapshot): Partial<EditSnapshot>
}

const REGISTRY = new Map<string, OpDef>()

export function register(def: OpDef): void {
  if (REGISTRY.has(def.name)) throw new Error(`duplicate op: ${def.name}`)
  REGISTRY.set(def.name, def)
}

export function opNames(): string[] {
  return [...REGISTRY.keys()]
}

export function lookup(name: string): OpDef {
  const def = REGISTRY.get(name)
  if (!def) {
    throw new GuidedError(`Unknown op "${name}". Available ops: ${opNames().join(', ')}`)
  }
  return def
}

export const BUCKETS: Bucket[] = [
  'markups',
  'annotDeletes',
  'noteEdits',
  'drawings',
  'textEdits',
  'textInserts',
  'imageEdits',
  'stampCfg',
  'formEdits',
  'rotations',
  'deleted',
  'order',
  'metadata',
]

export interface OpFailure {
  index: number
  op: Op
  error: string
}

export interface PlanResult {
  /** Stamped ops in input order; empty when the plan failed */
  ops: Op[]
  records: OpRecord[]
  failures: OpFailure[]
  touched: Set<Bucket>
}

/** Atomic plan: validate the whole batch and stamp ids; nothing is applied. */
export function planEditOps(ops: Op[], ctx: OpContext, newId: () => string): PlanResult {
  const stamped: Op[] = []
  const records: OpRecord[] = []
  const failures: OpFailure[] = []
  const touched = new Set<Bucket>()
  let cur = ctx
  ops.forEach((raw, index) => {
    try {
      const def = lookup(raw.op)
      if (cur.readOnly) throw new GuidedError('The document is read-only; no edit can be applied')
      def.validate(raw, cur)
      const adv = def.advance?.(raw, cur)
      if (adv) cur = { ...cur, ...adv }
      const op = def.additive && !raw.id ? { ...raw, id: newId() } : raw
      stamped.push(op)
      records.push(def.additive ? { op, created: [op.id as string] } : { op })
      for (const b of def.touches) touched.add(b)
    } catch (e) {
      const usage = opUsage(raw.op)
      const msg = e instanceof Error ? e.message : String(e)
      failures.push({ index, op: raw, error: usage ? `${msg}\n${usage}` : msg })
    }
  })
  return failures.length > 0
    ? { ops: [], records: [], failures, touched: new Set() }
    : { ops: stamped, records, failures, touched }
}

/** Reduce one bucket through a planned batch. */
export function reduceBucket<K extends Bucket>(
  bucket: K,
  prev: EditSnapshot[K],
  ops: Op[],
  base: EditSnapshot,
): EditSnapshot[K] {
  let cur = prev
  for (const op of ops) {
    const def = lookup(op.op)
    if (!def.touches.includes(bucket)) continue
    const out = def.apply(op, { ...base, [bucket]: cur })
    if (bucket in out) cur = out[bucket] as EditSnapshot[K]
  }
  return cur
}

/** Reduce a whole snapshot through a planned batch. */
export function reduceEditOps(state: EditSnapshot, ops: Op[]): EditSnapshot {
  let cur = state
  for (const op of ops) {
    const out = lookup(op.op).apply(op, cur)
    if (Object.keys(out).length > 0) cur = { ...cur, ...out }
  }
  return cur
}

export function runEditOps(
  state: EditSnapshot,
  ops: Op[],
  ctx: OpContext,
  newId: () => string,
): { state: EditSnapshot; plan: PlanResult } {
  const plan = planEditOps(ops, ctx, newId)
  return { state: plan.failures.length > 0 ? state : reduceEditOps(state, plan.ops), plan }
}
