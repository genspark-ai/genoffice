import { describe, it, expect } from 'vitest'
import {
  animClassOf,
  animStateKey,
  buildSteps,
  computeMediaCommands,
  computeNodeStates,
  parseParaStateKey,
  pointAtPath,
  samplePathPoints,
} from '../src/renderer/animation-play'
import type { AnimationItem } from '../src/shared/ipc'

const item = (
  o: Partial<AnimationItem> & Pick<AnimationItem, 'sourceId' | 'effect'>,
): AnimationItem => ({
  targetName: 'Shape',
  trigger: 'onClick',
  durationMs: 500,
  delayMs: 0,
  ...o,
})

describe('buildSteps', () => {
  it('groups onClick into new steps; withPrev/afterPrev join the current one', () => {
    const steps = buildSteps([
      item({ sourceId: 'a', effect: 'fade' }),
      item({ sourceId: 'b', effect: 'flyIn', trigger: 'withPrev', delayMs: 100 }),
      item({ sourceId: 'c', effect: 'wipe', trigger: 'afterPrev', delayMs: 50 }),
      item({ sourceId: 'a', effect: 'fadeOut' }),
    ])
    expect(steps).toHaveLength(2)
    expect(steps[0]!.auto).toBe(false)
    expect(steps[0]!.items.map((t) => [t.startMs, t.endMs])).toEqual([
      [0, 500], // a: onClick
      [100, 600], // b: withPrev = previous item's start 0 + 100
      [650, 1150], // c: afterPrev = previous item's end 600 + 50
    ])
    expect(steps[0]!.totalMs).toBe(1150)
    expect(steps[1]!.items).toHaveLength(1)
  })

  it('marks step 0 auto when the list starts with a non-click trigger', () => {
    const steps = buildSteps([item({ sourceId: 'a', effect: 'fade', trigger: 'withPrev' })])
    expect(steps[0]!.auto).toBe(true)
  })
})

describe('computeNodeStates', () => {
  const H = 540
  const anims = [
    item({ sourceId: 'a', effect: 'fade' }), // step 0
    item({ sourceId: 'b', effect: 'grow' }), // step 1
    item({ sourceId: 'a', effect: 'flyOut' }), // step 2
  ]
  const steps = buildSteps(anims)

  it('hides entrance targets before they play; leaves others visible', () => {
    const st = computeNodeStates(steps, 0, null, H)
    expect(st.get('a')!.hidden).toBe(true) // first animation is entrance -> hidden initially
    expect(st.get('b')!.hidden).toBe(false) // first is emphasis -> visible initially
  })

  it('interpolates the active step and finalizes played steps', () => {
    // Step 0 half-played: a is fading in
    const mid = computeNodeStates(steps, 0, 250, H)
    expect(mid.get('a')!.hidden).toBe(false)
    expect(mid.get('a')!.opacity).toBeGreaterThan(0)
    expect(mid.get('a')!.opacity).toBeLessThan(1)

    // Steps 0/1 finished: a normal, b stays scaled up at 1.5
    const after = computeNodeStates(steps, 2, null, H)
    expect(after.get('a')).toMatchObject({ hidden: false, opacity: 1 })
    expect(after.get('b')!.scale).toBeCloseTo(1.5)

    // All finished: a has flown out and is hidden
    const done = computeNodeStates(steps, 3, null, H)
    expect(done.get('a')!.hidden).toBe(true)
  })

  it('keeps pending items of the active step untouched until their start time', () => {
    const two = buildSteps([
      item({ sourceId: 'a', effect: 'fade' }),
      item({ sourceId: 'b', effect: 'fade', trigger: 'afterPrev' }),
    ])
    const st = computeNodeStates(two, 0, 100, H) // b doesn't start until 500ms
    expect(st.get('b')!.hidden).toBe(true)
  })
})

describe('new effects (P2 additions)', () => {
  const H = 540
  const W = 960

  it('classifies the new effects', () => {
    expect(animClassOf('bounce')).toBe('entrance')
    expect(animClassOf('splitIn')).toBe('entrance')
    expect(animClassOf('teeter')).toBe('emphasis')
    expect(animClassOf('wipeOut')).toBe('exit')
    expect(animClassOf('motionPath')).toBe('path')
  })

  it('wipe variants clip from the right edge (btm/top/mid)', () => {
    const steps = buildSteps([item({ sourceId: 'a', effect: 'wipeDown' })])
    const mid = computeNodeStates(steps, 0, 250, H, W).get('a')!
    expect(mid.hidden).toBe(false)
    expect(mid.clip!.mode).toBe('top')
    expect(mid.clip!.t).toBeGreaterThan(0)
    expect(mid.clip!.t).toBeLessThan(1)
    // No more clipping once finished
    expect(computeNodeStates(steps, 1, null, H, W).get('a')!.clip).toBeNull()

    const split = buildSteps([item({ sourceId: 'a', effect: 'splitIn' })])
    expect(computeNodeStates(split, 0, 250, H, W).get('a')!.clip!.mode).toBe('mid')
  })

  it('bounce drops in from above and settles at 0', () => {
    const steps = buildSteps([item({ sourceId: 'a', effect: 'bounce', durationMs: 1000 })])
    const early = computeNodeStates(steps, 0, 50, H, W).get('a')!
    expect(early.hidden).toBe(false)
    expect(early.dy).toBeLessThan(0)
    expect(computeNodeStates(steps, 1, null, H, W).get('a')!.dy).toBe(0)
  })

  it('flipIn expands horizontally (scaleX 0→1)', () => {
    const steps = buildSteps([item({ sourceId: 'a', effect: 'flipIn' })])
    const mid = computeNodeStates(steps, 0, 250, H, W).get('a')!
    expect(mid.scaleX).toBeGreaterThan(0)
    expect(mid.scaleX).toBeLessThan(1)
    expect(computeNodeStates(steps, 1, null, H, W).get('a')!.scaleX).toBe(1)
  })

  it('teeter rocks around 0 and returns to rest; shrink/wipeOut hide at the end', () => {
    const teeter = buildSteps([item({ sourceId: 'a', effect: 'teeter', durationMs: 1000 })])
    expect(computeNodeStates(teeter, 0, 170, H, W).get('a')!.rotationDeg).not.toBe(0)
    expect(computeNodeStates(teeter, 1, null, H, W).get('a')!.rotationDeg).toBe(0)

    const shrink = buildSteps([item({ sourceId: 'a', effect: 'shrink' })])
    expect(computeNodeStates(shrink, 1, null, H, W).get('a')!.hidden).toBe(true)
    const wipeOut = buildSteps([item({ sourceId: 'a', effect: 'wipeOut' })])
    expect(computeNodeStates(wipeOut, 1, null, H, W).get('a')!.hidden).toBe(true)
  })
})

describe('paragraph animations', () => {
  const H = 540

  it('animStateKey separates paragraph targets and round-trips through parse', () => {
    expect(animStateKey('el-1')).toBe('el-1')
    expect(animStateKey('el-1', 0)).not.toBe('el-1')
    expect(parseParaStateKey(animStateKey('el-1', 2))).toEqual({ sourceId: 'el-1', para: 2 })
    expect(parseParaStateKey('el-1')).toBeNull()
  })

  it('same-shape paragraphs build/hide independently, one bullet per click', () => {
    const steps = buildSteps([
      item({ sourceId: 'a', effect: 'fade', paragraph: 0 }),
      item({ sourceId: 'a', effect: 'fade', paragraph: 1 }),
    ])
    expect(steps).toHaveLength(2)

    const k0 = animStateKey('a', 0)
    const k1 = animStateKey('a', 1)
    // Not played: both paragraphs hidden; the whole shape has no state
    const fresh = computeNodeStates(steps, 0, null, H)
    expect(fresh.get(k0)!.hidden).toBe(true)
    expect(fresh.get(k1)!.hidden).toBe(true)
    expect(fresh.has('a')).toBe(false)
    // After the first click: paragraph 0 visible, paragraph 1 still hidden
    const one = computeNodeStates(steps, 1, null, H)
    expect(one.get(k0)!.hidden).toBe(false)
    expect(one.get(k1)!.hidden).toBe(true)
    expect(computeNodeStates(steps, 2, null, H).get(k1)!.hidden).toBe(false)
  })

  it('mixes paragraph and whole-shape animations without cross-talk', () => {
    const steps = buildSteps([
      item({ sourceId: 'a', effect: 'fade', paragraph: 0 }),
      item({ sourceId: 'a', effect: 'grow' }), // whole-shape emphasis
    ])
    const after = computeNodeStates(steps, 2, null, H)
    expect(after.get(animStateKey('a', 0))!.hidden).toBe(false)
    expect(after.get('a')!.scale).toBeCloseTo(1.5)
  })
})

describe('motion paths', () => {
  const H = 540
  const W = 960

  it('samples M/L polylines and closes Z back to the start', () => {
    const pts = samplePathPoints('M 0 0 L 0.5 0 L 0.5 0.5 Z')
    expect(pts[0]).toEqual({ x: 0, y: 0 })
    expect(pts.at(-1)).toEqual({ x: 0, y: 0 })
    // The arc-length midpoint lies on the polyline
    const mid = pointAtPath(pts, 0.5)
    expect(mid.x).toBeCloseTo(0.5, 5)
  })

  it('subdivides cubic beziers', () => {
    const pts = samplePathPoints('M 0 0 C 0.069 0 0.125 0.056 0.125 0.125')
    expect(pts.length).toBeGreaterThan(10)
    expect(pts.at(-1)!.x).toBeCloseTo(0.125, 5)
    expect(pts.at(-1)!.y).toBeCloseTo(0.125, 5)
  })

  it('moves the node along the path and keeps it at the end point', () => {
    const steps = buildSteps([
      item({ sourceId: 'a', effect: 'motionPath', motionPath: 'M 0 0 L 0.5 0', durationMs: 1000 }),
    ])
    // Not played: the motion path target stays visible (not an entrance type)
    expect(computeNodeStates(steps, 0, null, H, W).get('a')!.hidden).toBe(false)
    const mid = computeNodeStates(steps, 0, 500, H, W).get('a')!
    expect(mid.dx).toBeGreaterThan(0)
    expect(mid.dx).toBeLessThan(0.5 * W)
    expect(mid.dy).toBe(0)
    // When finished, it rests at the path end (0.5 × canvas width)
    const done = computeNodeStates(steps, 1, null, H, W).get('a')!
    expect(done.dx).toBeCloseTo(0.5 * W, 5)
  })

  it('legacy items without motionPath fall back to the default path (no crash)', () => {
    const steps = buildSteps([item({ sourceId: 'a', effect: 'motionPath', durationMs: 1000 })])
    const done = computeNodeStates(steps, 1, null, H, W).get('a')!
    expect(done.dx).toBeCloseTo(0.25 * W, 5)
  })
})

describe('media commands', () => {
  const play = item({ sourceId: 'v', effect: 'mediaPlay', durationMs: 0 })
  const fade = item({ sourceId: 't', effect: 'fade' })

  it('media items take a step but leave no visual state', () => {
    const steps = buildSteps([play, fade])
    expect(steps).toHaveLength(2)
    const states = computeNodeStates(steps, 0, null, 900)
    expect(states.has('v')).toBe(false)
    expect(states.get('t')!.hidden).toBe(true)
    expect(animClassOf('mediaPause')).toBe('media')
  })

  it('fires commands of played steps in order and the current step by time', () => {
    const stop = item({ sourceId: 'v', effect: 'mediaStop', trigger: 'withPrev', delayMs: 300 })
    const steps = buildSteps([play, fade, stop])
    expect(computeMediaCommands(steps, 0, null)).toEqual([])
    expect(computeMediaCommands(steps, 0, 1)).toEqual([{ sourceId: 'v', effect: 'mediaPlay' }])
    expect(computeMediaCommands(steps, 1, null)).toEqual([{ sourceId: 'v', effect: 'mediaPlay' }])
    // Second step: the fade is running but the stop (delay 300) has not started
    expect(computeMediaCommands(steps, 1, 100)).toHaveLength(1)
    expect(computeMediaCommands(steps, 1, 301).map((c) => c.effect)).toEqual([
      'mediaPlay',
      'mediaStop',
    ])
    expect(computeMediaCommands(steps, 2, null)).toHaveLength(2)
  })

  it('an auto first step fires its media command on page entry', () => {
    const auto = item({ sourceId: 'v', effect: 'mediaPlay', trigger: 'withPrev', durationMs: 0 })
    const steps = buildSteps([auto])
    expect(steps[0]!.auto).toBe(true)
    expect(computeMediaCommands(steps, 1, null)).toEqual([{ sourceId: 'v', effect: 'mediaPlay' }])
  })
})
