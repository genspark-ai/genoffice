/**
 * Canonical edit-op registry — the first slice of the unified write contract:
 * the model is the executor's private property, the op is the public contract.
 * Every op supplies plan-time validation and an apply that mutates the engine
 * model and returns a before/after record for the transaction journal.
 *
 * Guided errors are a contract, not a style choice: validation failures must
 * state what is wrong AND what to do (available ids included) — models
 * self-correct from "no element X on slide N. Available: [...]" but
 * blind-retry a bare failure.
 *
 * Addressing is transitional: { slide: index, el: parse-time element id }.
 * Durable identity (a16:creationId) is the separate step-0 line; the envelope
 * shape stays the same when it lands.
 *
 * Master/layout chrome shares this vocabulary: element ops that only mutate
 * the slide model accept target.part instead of target.slide, resolved
 * through a per-transaction parse cache; the executor serializes touched
 * parts back once per transaction (the flush also re-materializes every deck
 * slide, because slides resolve inherited styles against the chrome).
 */
import {
  elementDurableId,
  groupChildDurableId,
  listMasterParts,
  matchesElementRef,
  parseMasterPart,
  slideDurableId,
} from '@genoffice/pptx-engine'
export { elementDurableId, matchesElementRef, slideDurableId }
import type { OpenedPptx, Slide, SlideElement } from '@genoffice/pptx-engine'

export class GuidedError extends Error {}

export interface OpTarget {
  /** 0-based slide index, or the durable slide id ("s_<n>", from the part path — save never renames slide parts) */
  slide?: number | string
  /** master/layout part path ("ppt/slideMasters/slideMaster1.xml") — the op edits deck chrome instead of a slide */
  part?: string
  /** element id within the slide: the parse-time id or the durable id ("e_<guid8>" from a16:creationId, falling back to "e_<cNvPr id>" — both live in the file bytes and survive save→reopen/reparse) */
  el?: string
}

export interface Op {
  op: string
  target?: OpTarget
  [key: string]: unknown
}

/** One executed op with captured before/after — inverse ops derive from this. */
export interface OpRecord {
  op: Op
  before?: unknown
  after?: unknown
  /** ids minted by additive ops */
  created?: string[]
}

export interface OpContext {
  opened: OpenedPptx
  /** Per-transaction parse cache for part-addressed (master/layout) targets. */
  parts: Map<string, Slide>
  /** Parts resolved during apply; the executor flushes these back to entries once per transaction. */
  touchedParts: Set<string>
}

export interface OpDef {
  name: string
  /** Plan-time validation; throws GuidedError. Runs against the pre-transaction state. */
  validate(op: Op, ctx: OpContext): void
  /** Execution; mutates the model and returns the record. Throws GuidedError on failure. */
  apply(op: Op, ctx: OpContext): OpRecord
}

const REGISTRY = new Map<string, OpDef>()

export function register(def: OpDef): void {
  REGISTRY.set(def.name, def)
}

export function lookup(name: string): OpDef {
  const def = REGISTRY.get(name)
  if (!def) {
    throw new GuidedError(
      `unknown op "${name}". Supported ops: [${[...REGISTRY.keys()].join(', ')}].`,
    )
  }
  return def
}

export function opNames(): string[] {
  return [...REGISTRY.keys()]
}

// ── target resolution (guided) ──────────────────────────────────────────

function resolvePart(ctx: OpContext, op: Op, part: string): Slide {
  let slide = ctx.parts.get(part)
  if (!slide) {
    const known = listMasterParts(ctx.opened.archive)
    if (!known.some((p) => p.partPath === part)) {
      throw new GuidedError(
        `op "${op.op}": no master/layout part "${part}". Available: [${known.map((p) => p.partPath).join(', ')}].`,
      )
    }
    const parsed = parseMasterPart(ctx.opened.archive, part)
    if (!parsed) {
      throw new GuidedError(`op "${op.op}": part "${part}" could not be parsed.`)
    }
    slide = parsed
    ctx.parts.set(part, slide)
  }
  ctx.touchedParts.add(part)
  return slide
}

export interface ResolveOpts {
  types?: SlideElement['type'][]
  /** Ops that only mutate the slide model may opt in to master/layout part targets. */
  allowPart?: boolean
}

export function resolveSlide(
  ctx: OpContext,
  op: Op,
  opts: ResolveOpts = {},
): { index: number; slide: Slide } {
  const part = op.target?.part
  if (typeof part === 'string' && part) {
    if (!opts.allowPart) {
      throw new GuidedError(
        `op "${op.op}" cannot target a master/layout part — address a slide via target.slide.`,
      )
    }
    return { index: -1, slide: resolvePart(ctx, op, part) }
  }
  const ref = op.target?.slide
  const slides = ctx.opened.deck.slides
  if (typeof ref === 'string') {
    const index = slides.findIndex((s) => slideDurableId(s) === ref)
    if (index < 0) {
      throw new GuidedError(
        `op "${op.op}": no slide "${ref}". Available: [${slides.map(slideDurableId).join(', ')}].`,
      )
    }
    return { index, slide: slides[index]! }
  }
  if (typeof ref !== 'number') {
    throw new GuidedError(
      `op "${op.op}" needs target.slide (a 0-based index or a durable "s_<n>" id; this deck has ${slides.length} slide(s)).`,
    )
  }
  const slide = slides[ref]
  if (!slide) {
    throw new GuidedError(
      `op "${op.op}": slide index ${ref} is out of range (0-${slides.length - 1}).`,
    )
  }
  return { index: ref, slide }
}

export function resolveElement(
  ctx: OpContext,
  op: Op,
  opts: ResolveOpts = {},
): { index: number; slide: Slide; el: SlideElement } {
  const { index, slide } = resolveSlide(ctx, op, opts)
  const where = op.target?.part ? `part "${op.target.part}"` : `slide ${index}`
  const id = op.target?.el
  if (typeof id !== 'string' || !id) {
    throw new GuidedError(`op "${op.op}" needs target.el (an element id on ${where}).`)
  }
  const el = slide.elements.find((x) => matchesElementRef(x, id))
  if (!el) {
    const available = slide.elements
      .map((x) => {
        const durable = elementDurableId(x)
        return durable ? `${x.id} (${durable})` : x.id
      })
      .join(', ')
    throw new GuidedError(
      `op "${op.op}": no element "${id}" on ${where}. Available: [${available}].`,
    )
  }
  if (opts.types && !opts.types.includes(el.type)) {
    throw new GuidedError(
      `op "${op.op}" targets a ${el.type} element — it only works on ${opts.types.join('/')} elements.`,
    )
  }
  return { index, slide, el }
}

/** Group ops carry the group's top-level id (parse-time or durable); returns the
    parse-time id the engine matches, with a guided listing on a miss. */
export function resolveGroup(op: Op, slideIdx: number, groups: SlideElement[]): string {
  const groupId = op.group
  if (typeof groupId !== 'string' || !groupId) {
    throw new GuidedError(`op "${op.op}": "group" must be the group element's id.`)
  }
  const grp = groups.find((g) => g.type === 'group' && matchesElementRef(g, groupId))
  if (!grp) {
    const available = groups
      .filter((g) => g.type === 'group')
      .map((g) => {
        const durable = elementDurableId(g)
        return durable ? `${g.id} (${durable})` : g.id
      })
    throw new GuidedError(
      `op "${op.op}": no group "${groupId}" on slide ${slideIdx}. Available groups: [${available.join(', ')}].`,
    )
  }
  return grp.id
}

/** Group-child references may be durable too: translate to the child's parse-time
    id (what the engine matches) by walking the group's children. Unknown refs pass
    through unchanged — the engine call then fails with its own guided error. */
export function resolveGroupChildId(slide: Slide, groupId: string, ref: string): string {
  const grp = slide.elements.find((x) => x.id === groupId && x.type === 'group')
  const children = (grp as { children?: SlideElement[] } | undefined)?.children
  const child = children?.find(
    (c) =>
      c.id === ref ||
      (grp && groupChildDurableId(grp, c) === ref) ||
      // nvId form: what children showed before creationId continuity landed
      elementDurableId(c) === ref,
  )
  return child ? child.id : ref
}
