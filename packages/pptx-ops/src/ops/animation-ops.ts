/**
 * Per-item animation ops. The engine models a slide's timeline as an ordered
 * SlideAnimation[] (spid-addressed); these ops address elements by id and edit
 * one item at a time, so an agent can add, drop or reorder a single effect
 * without re-sending the whole list.
 */
import {
  ANIM_DIRECTIONS,
  ANIM_EFFECTS,
  ANIM_TRIGGERS,
  animClassOf,
  elementDurableId,
  elementSpid,
  getSlideAnimations,
  setSlideAnimations,
  type AnimClass,
  type AnimDirection,
  type AnimEffectKind,
  type AnimTrigger,
  type Slide,
  type SlideAnimation,
  type SlideElement,
} from '@genoffice/pptx-engine'
import {
  GuidedError,
  register,
  resolveElement,
  resolveSlide,
  type Op,
  type OpRecord,
} from './registry'

export interface AnimationEntry {
  /** 0-based position in the slide's timeline */
  seq: number
  /** element id (durable when the shape has one), null when the target shape is gone */
  el: string | null
  spid: number
  effect: AnimEffectKind
  kind: AnimClass
  trigger: AnimTrigger
  durationMs: number
  delayMs: number
  direction?: AnimDirection
  paragraph?: number
  motionPath?: string
  /** the file's effect has no modeled writer: `effect` is the closest approximation, the bytes are kept verbatim */
  custom?: true
  /** presetID/presetClass/presetSubtype of such a kept-verbatim effect */
  preset?: { id: number; cls: string; sub: number }
}

function elementBySpid(elements: SlideElement[], spid: number): SlideElement | undefined {
  for (const el of elements) {
    if (elementSpid(el) === spid) return el
    if (el.type === 'group') {
      const hit = elementBySpid(el.children, spid)
      if (hit) return hit
    }
  }
  return undefined
}

/** The slide's timeline flattened for read surfaces (CLI `slides read`, AI context). */
export function listSlideAnimations(slide: Slide): AnimationEntry[] {
  return getSlideAnimations(slide).map((a, seq) => {
    const el = elementBySpid(slide.elements, a.spid)
    return {
      seq,
      el: el ? (elementDurableId(el) ?? el.id) : null,
      spid: a.spid,
      effect: a.effect,
      kind: animClassOf(a.effect),
      trigger: a.trigger,
      durationMs: a.durationMs,
      delayMs: a.delayMs,
      ...(a.direction ? { direction: a.direction } : {}),
      ...(a.paragraph != null ? { paragraph: a.paragraph } : {}),
      ...(a.motionPath ? { motionPath: a.motionPath } : {}),
      ...(a.preset ? { preset: a.preset } : {}),
      ...(a.presetXml ? { custom: true as const } : {}),
    }
  })
}

const EFFECT_ALIASES: Record<string, AnimEffectKind> = {
  fly: 'flyIn',
  flyin: 'flyIn',
  fadein: 'fade',
  wipein: 'wipe',
  split: 'splitIn',
  flip: 'flipIn',
  zoomin: 'zoom',
  float: 'flyIn',
  floatin: 'flyIn',
  fadeout: 'fadeOut',
  flyout: 'flyOut',
  wipeout: 'wipeOut',
  zoomout: 'zoomOut',
  path: 'motionPath',
  motion: 'motionPath',
}

const TRIGGER_ALIASES: Record<string, AnimTrigger> = {
  onclick: 'onClick',
  click: 'onClick',
  withprev: 'withPrev',
  withprevious: 'withPrev',
  with: 'withPrev',
  afterprev: 'afterPrev',
  afterprevious: 'afterPrev',
  after: 'afterPrev',
}

const DIRECTIONAL = new Set<AnimEffectKind>(['flyIn', 'flyOut', 'wipe', 'wipeOut'])
const CLASS_DEFAULT_EFFECT: Record<string, AnimEffectKind> = {
  entr: 'fade',
  emph: 'pulse',
  exit: 'fadeOut',
  path: 'motionPath',
}

function defaultDuration(effect: AnimEffectKind): number {
  if (effect === 'appear' || effect === 'disappear') return 0
  if (effect === 'spin' || effect === 'grow' || effect === 'bounce' || effect === 'motionPath')
    return 2000
  if (effect === 'pulse' || effect === 'teeter') return 1000
  return 500
}

function parseEffect(op: Op): AnimEffectKind | undefined {
  if (op.effect === undefined) return undefined
  if (typeof op.effect !== 'string') {
    throw new GuidedError(`op "${op.op}": "effect" must be one of [${ANIM_EFFECTS.join(', ')}].`)
  }
  const exact = ANIM_EFFECTS.find((e) => e === op.effect)
  if (exact) return exact
  const folded = op.effect.toLowerCase().replace(/[^a-z]/g, '')
  const loose = ANIM_EFFECTS.find((e) => e.toLowerCase() === folded) ?? EFFECT_ALIASES[folded]
  if (loose) return loose
  throw new GuidedError(
    `op "${op.op}": unknown effect ${JSON.stringify(op.effect)}. Use one of [${ANIM_EFFECTS.join(', ')}].`,
  )
}

function parseTrigger(op: Op): AnimTrigger {
  if (op.trigger === undefined) return 'onClick'
  if (typeof op.trigger === 'string') {
    const exact = ANIM_TRIGGERS.find((t) => t === op.trigger)
    if (exact) return exact
    const loose = TRIGGER_ALIASES[op.trigger.toLowerCase().replace(/[^a-z]/g, '')]
    if (loose) return loose
  }
  throw new GuidedError(
    `op "${op.op}": "trigger" must be one of [${ANIM_TRIGGERS.join(', ')}] (default onClick).`,
  )
}

function parseMs(op: Op, field: 'duration' | 'delay', fallback: number): number {
  const v = op[`${field}Ms`] ?? op[field]
  if (v === undefined) return fallback
  if (typeof v !== 'number' || !Number.isFinite(v) || v < 0) {
    throw new GuidedError(`op "${op.op}": "${field}" must be >= 0 milliseconds.`)
  }
  return Math.round(v)
}

function parseSeq(op: Op, field: string, required: boolean): number | undefined {
  const v = op[field]
  if (v === undefined) {
    if (required)
      throw new GuidedError(`op "${op.op}" needs "${field}": a 0-based timeline position.`)
    return undefined
  }
  if (typeof v !== 'number' || !Number.isInteger(v) || v < 0) {
    throw new GuidedError(`op "${op.op}": "${field}" must be a 0-based integer timeline position.`)
  }
  return v
}

function timelineRange(anims: SlideAnimation[]): string {
  return anims.length ? `0-${anims.length - 1}` : 'empty (no animations on this slide)'
}

interface PresetXmlInfo {
  xml: string
  cls: string
}

/** A verbatim effect <p:par> (PowerPoint-exported) must be one balanced block targeting this shape. */
function parsePresetXml(op: Op, spid: number): PresetXmlInfo | undefined {
  if (op.presetXml === undefined) return undefined
  const xml = typeof op.presetXml === 'string' ? op.presetXml.trim() : ''
  if (!/^<p:par\b[\s\S]*<\/p:par>$/.test(xml)) {
    throw new GuidedError(
      `op "${op.op}": "presetXml" must be one effect <p:par>…</p:par> block as PowerPoint writes it inside the main sequence.`,
    )
  }
  const cls = /<p:cTn\b[^>]*\bpresetClass="([^"]*)"/.exec(xml)?.[1]
  if (!cls || !/<p:cTn\b[^>]*\bpresetID="\d+"/.test(xml)) {
    throw new GuidedError(
      `op "${op.op}": "presetXml" needs a <p:cTn> with presetID and presetClass attributes.`,
    )
  }
  const spids = [...xml.matchAll(/<p:spTgt\s[^>]*\bspid="(\d+)"/g)].map((m) => Number(m[1]))
  if (!spids.length || spids.some((s) => s !== spid)) {
    throw new GuidedError(
      `op "${op.op}": every <p:spTgt spid> in "presetXml" must be ${spid} (the target element's shape id); found [${spids.join(', ')}].`,
    )
  }
  if ((xml.match(/<p:par\b/g) ?? []).length !== (xml.match(/<\/p:par>/g) ?? []).length) {
    throw new GuidedError(`op "${op.op}": "presetXml" is not balanced.`)
  }
  return { xml, cls }
}

function buildAnimation(op: Op, el: SlideElement): SlideAnimation {
  const spid = elementSpid(el)
  if (spid == null) {
    throw new GuidedError(`op "${op.op}": element "${el.id}" has no shape id to animate.`)
  }
  const preset = parsePresetXml(op, spid)
  let effect = parseEffect(op)
  if (!effect) {
    if (!preset) {
      throw new GuidedError(
        `op "${op.op}" needs "effect": one of [${ANIM_EFFECTS.join(', ')}] (or "presetXml").`,
      )
    }
    effect = CLASS_DEFAULT_EFFECT[preset.cls] ?? 'fade'
  }
  if (op.kind !== undefined) {
    const want = String(op.kind)
    const have = animClassOf(effect)
    if (want !== have && !(want === 'motion' && have === 'path')) {
      throw new GuidedError(
        `op "${op.op}": "${effect}" is an ${have} effect, not ${want}; pick an effect of that class (${ANIM_EFFECTS.filter((e) => animClassOf(e) === want).join(', ') || 'none'}).`,
      )
    }
  }
  const anim: SlideAnimation = {
    spid,
    effect,
    trigger: parseTrigger(op),
    durationMs: parseMs(op, 'duration', defaultDuration(effect)),
    delayMs: parseMs(op, 'delay', 0),
  }
  if (op.direction !== undefined) {
    const d = ANIM_DIRECTIONS.find((x) => x === op.direction)
    if (!d) {
      throw new GuidedError(
        `op "${op.op}": "direction" must be one of [${ANIM_DIRECTIONS.join(', ')}].`,
      )
    }
    if (!DIRECTIONAL.has(effect)) {
      throw new GuidedError(
        `op "${op.op}": "direction" only applies to ${[...DIRECTIONAL].join('/')}, not ${effect}.`,
      )
    }
    anim.direction = d
  }
  if (op.paragraph !== undefined) {
    if (typeof op.paragraph !== 'number' || !Number.isInteger(op.paragraph) || op.paragraph < 0) {
      throw new GuidedError(`op "${op.op}": "paragraph" must be a 0-based paragraph index.`)
    }
    anim.paragraph = op.paragraph
  }
  if (op.motionPath !== undefined) {
    if (typeof op.motionPath !== 'string' || !op.motionPath.trim()) {
      throw new GuidedError(
        `op "${op.op}": "motionPath" must be an SVG-like path (M/L/C/Z, coordinates 0..1 of the slide).`,
      )
    }
    if (effect !== 'motionPath') {
      throw new GuidedError(`op "${op.op}": "motionPath" only applies to effect "motionPath".`)
    }
    anim.motionPath = op.motionPath.trim()
  }
  if (preset) anim.presetXml = preset.xml
  return anim
}

register({
  name: 'addAnimation',
  validate(op, ctx) {
    const { el } = resolveElement(ctx, op)
    buildAnimation(op, el)
    parseSeq(op, 'after', false)
  },
  apply(op, ctx): OpRecord {
    const { slide, el } = resolveElement(ctx, op)
    const anim = buildAnimation(op, el)
    const anims = getSlideAnimations(slide)
    const after = parseSeq(op, 'after', false)
    if (after !== undefined && after >= anims.length) {
      throw new GuidedError(
        `op "addAnimation": "after" ${after} is past the timeline (${timelineRange(anims)}); omit it to append.`,
      )
    }
    const seq = after === undefined ? anims.length : after + 1
    anims.splice(seq, 0, anim)
    setSlideAnimations(slide, anims)
    return { op, after: { seq, count: anims.length } }
  },
})

register({
  name: 'removeAnimation',
  validate(op, ctx) {
    const seq = parseSeq(op, 'seq', false)
    if (seq === undefined) {
      resolveElement(ctx, op)
    } else {
      resolveSlide(ctx, op)
    }
  },
  apply(op, ctx): OpRecord {
    const seq = parseSeq(op, 'seq', false)
    if (seq !== undefined) {
      const { slide } = resolveSlide(ctx, op)
      const anims = getSlideAnimations(slide)
      if (seq >= anims.length) {
        throw new GuidedError(
          `op "removeAnimation": no animation at seq ${seq}; the timeline is ${timelineRange(anims)}.`,
        )
      }
      const [removed] = anims.splice(seq, 1)
      setSlideAnimations(slide, anims)
      return { op, before: removed, after: { removed: 1, count: anims.length } }
    }
    const { slide, el } = resolveElement(ctx, op)
    const spid = elementSpid(el)
    const anims = getSlideAnimations(slide)
    const kept = anims.filter((a) => a.spid !== spid)
    if (kept.length !== anims.length) setSlideAnimations(slide, kept)
    return { op, after: { removed: anims.length - kept.length, count: kept.length } }
  },
})

register({
  name: 'reorderAnimation',
  validate(op, ctx) {
    resolveSlide(ctx, op)
    parseSeq(op, 'seq', true)
    parseSeq(op, 'to', true)
  },
  apply(op, ctx): OpRecord {
    const { slide } = resolveSlide(ctx, op)
    const seq = parseSeq(op, 'seq', true)!
    const to = parseSeq(op, 'to', true)!
    const anims = getSlideAnimations(slide)
    if (seq >= anims.length || to >= anims.length) {
      throw new GuidedError(
        `op "reorderAnimation": seq/to must be inside the timeline (${timelineRange(anims)}).`,
      )
    }
    if (seq !== to) {
      const [item] = anims.splice(seq, 1)
      anims.splice(to, 0, item!)
      setSlideAnimations(slide, anims)
    }
    return { op, after: { seq: to, count: anims.length } }
  },
})
