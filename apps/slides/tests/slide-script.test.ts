/**
 * execute_slide_script (unified editing script) tests:
 *  - runLayoutScript new primitives: moveBy/resizeBy/setText/setStyle/setFill/setStroke collect the right ops
 *  - inGroup/locked elements are rejected by every write primitive
 *  - Tool chain: script -> batchEditTransform + editText/editFill/editStroke dispatched in order -> audit report back
 */
import { describe, it, expect, vi, beforeEach } from 'vitest'
import type { RenderSlide, RenderNode, ShapeRenderNode, PlacedBox } from '@genoffice/pptx-render'
import { runLayoutScript, type LayoutScriptElement } from '../src/renderer/ai/layout-script'
import { createSlidesSkill, type DeckAccess } from '../src/renderer/ai/slides-skill'

const box = (x: number, y: number, w: number, h: number, rot = 0): PlacedBox => ({
  x,
  y,
  w,
  h,
  rotationDeg: rot,
  flipH: false,
  flipV: false,
  centerX: x + w / 2,
  centerY: y + h / 2,
})

/** Build a shape node with text layout (same as layout-tools.test.ts). */
const textNode = (id: string, b: PlacedBox, text: string): ShapeRenderNode => ({
  id,
  sourceId: id,
  type: 'shape',
  box: b,
  fill: { kind: 'none' },
  text: {
    lines: [
      {
        runs: [
          {
            text,
            x: 8,
            baselineY: 20,
            fontFamily: 'Arial',
            fontSizePx: 24,
            color: '#000000',
            bold: false,
            italic: false,
            underline: false,
            widthPx: text.length * 12,
          },
        ],
        top: 0,
        height: 28,
      },
    ],
    insets: { l: 8, t: 4, r: 8, b: 4 },
    anchor: 'top',
    fontScale: 1,
    wrap: true,
    contentHeight: 28,
  },
})

const groupNode = (id: string, b: PlacedBox, children: RenderNode[]): RenderNode =>
  ({ id, sourceId: id, type: 'group', box: b, children }) as unknown as RenderNode

const slideOf = (nodes: RenderNode[], w = 1280, h = 720): RenderSlide => ({
  widthPx: w,
  heightPx: h,
  scale: 1,
  background: { kind: 'solid', color: '#FFFFFF' },
  nodes,
})

// ── New sandbox primitives ─────────────────────────────────────

describe('runLayoutScript editing primitives', () => {
  const els: LayoutScriptElement[] = [
    { id: 'a', type: 'shape', text: 'Title', x: 100, y: 200, w: 300, h: 100, rotation: 0 },
    { id: 'b', type: 'shape', text: 'Body', x: 480, y: 210, w: 310, h: 300, rotation: 0 },
    {
      id: 'child',
      type: 'text',
      text: 'In group',
      x: 10,
      y: 10,
      w: 50,
      h: 20,
      rotation: 0,
      inGroup: true,
    },
    {
      id: 'deco',
      type: 'shape',
      text: 'Decoration',
      x: 0,
      y: 0,
      w: 1280,
      h: 40,
      rotation: 0,
      locked: true,
    },
  ]
  const canvas = { w: 1280, h: 720 }

  it('moveBy accumulates on the current value (multiple calls mixed with setBox)', () => {
    const r = runLayoutScript(
      `moveBy('a', -30, 0); moveBy('a', 10, 5); setBox('a', { w: 320 }); moveBy('a', 0, -5);`,
      els,
      canvas,
    )
    expect(r.error).toBeUndefined()
    expect(r.ops).toHaveLength(1)
    expect(r.ops[0]).toMatchObject({ id: 'a', x: 80, y: 200, w: 320, h: 100 })
  })

  it('resizeBy resizes relatively, result never below 1px', () => {
    const r = runLayoutScript(`resizeBy('b', -10, 20); resizeBy('a', -9999, -9999);`, els, canvas)
    expect(r.error).toBeUndefined()
    expect(r.ops.find((o) => o.id === 'b')).toMatchObject({ w: 300, h: 320 })
    expect(r.ops.find((o) => o.id === 'a')).toMatchObject({ w: 1, h: 1 })
  })

  it('setText string collects a text op (split into paragraphs by \\n, one run each)', () => {
    const r = runLayoutScript(`setText('a', 'New title\\nSecond line');`, els, canvas)
    expect(r.error).toBeUndefined()
    expect(r.edits).toEqual([
      {
        kind: 'text',
        id: 'a',
        paragraphs: [{ runs: [{ text: 'New title' }] }, { runs: [{ text: 'Second line' }] }],
      },
    ])
  })

  it('setText passes paragraph arrays through (same flat format as apply_ops setText)', () => {
    const r = runLayoutScript(
      `setText('a', [{ text: 'Big title', bold: true, fontSize: 40, align: 'center' }]);`,
      els,
      canvas,
    )
    expect(r.error).toBeUndefined()
    expect(r.edits).toEqual([
      {
        kind: 'text',
        id: 'a',
        paragraphs: [{ runs: [{ text: 'Big title', bold: true, fontSize: 40 }], align: 'center' }],
      },
    ])
  })

  it('setStyle collects a style patch (colors normalized to # prefix)', () => {
    const r = runLayoutScript(
      `setStyle('a', { color: '1a73e8', bold: true, align: 'right' });`,
      els,
      canvas,
    )
    expect(r.error).toBeUndefined()
    expect(r.edits).toEqual([
      { kind: 'style', id: 'a', style: { color: '#1a73e8', bold: true, align: 'right' } },
    ])
  })

  it('setStyle errors on invalid color / empty patch', () => {
    expect(runLayoutScript(`setStyle('a', { color: 'blue' });`, els, canvas).error).toContain(
      '#RRGGBB',
    )
    expect(runLayoutScript(`setStyle('a', {});`, els, canvas).error).toContain(
      'at least one style field',
    )
  })

  it('setFill collects fill (none/#RRGGBB), errors on invalid values', () => {
    const r = runLayoutScript(`setFill('a', 'none'); setFill('b', 'FF0000');`, els, canvas)
    expect(r.error).toBeUndefined()
    expect(r.edits).toEqual([
      { kind: 'fill', id: 'a', fill: 'none' },
      { kind: 'fill', id: 'b', fill: '#FF0000' },
    ])
    expect(runLayoutScript(`setFill('a', 'red');`, els, canvas).error).toContain('#RRGGBB')
  })

  it('setStroke collects stroke (null = remove; color/width have defaults)', () => {
    const r = runLayoutScript(
      `setStroke('a', { color: '#ff0000', widthPt: 2 }); setStroke('b', null); setStroke('b', {});`,
      els,
      canvas,
    )
    expect(r.error).toBeUndefined()
    expect(r.edits).toEqual([
      { kind: 'stroke', id: 'a', stroke: { color: '#ff0000', widthPt: 2 } },
      { kind: 'stroke', id: 'b', stroke: null },
      { kind: 'stroke', id: 'b', stroke: { color: '#000000', widthPt: 1 } },
    ])
  })

  it('one script mixing multiple operations keeps edits in call order', () => {
    const r = runLayoutScript(
      `
        moveBy('a', -30, 0);
        setText('a', 'Main title');
        setStyle('a', { color: '#1a73e8', bold: true });
        setFill('b', '#f8fafc');
        setStroke('b', { color: '#e2e8f0', widthPt: 1 });
        return 'done';
      `,
      els,
      canvas,
    )
    expect(r.error).toBeUndefined()
    expect(r.returned).toBe('done')
    expect(r.ops).toHaveLength(1)
    expect(r.edits.map((e) => e.kind)).toEqual(['text', 'style', 'fill', 'stroke'])
  })

  it('children nested in a sub-group (inGroup without groupId) are rejected by all write primitives', () => {
    for (const call of [
      `moveBy('child', 1, 1)`,
      `resizeBy('child', 1, 1)`,
      `setText('child', 'x')`,
      `setStyle('child', { bold: true })`,
      `setFill('child', 'none')`,
      `setStroke('child', null)`,
    ]) {
      const r = runLayoutScript(call, els, canvas)
      expect(r.error).toContain('sub-group')
      expect(r.ops).toHaveLength(0)
      expect(r.edits).toHaveLength(0)
    }
  })

  it('direct group children (inGroup+groupId) accept all write primitives; ops/edits carry groupId', () => {
    const els2: LayoutScriptElement[] = [
      ...els,
      {
        id: 'gc',
        type: 'shape',
        text: 'Member',
        x: 210,
        y: 120,
        w: 80,
        h: 40,
        rotation: 0,
        inGroup: true,
        groupId: 'grp1',
      },
    ]
    const r = runLayoutScript(
      `moveBy('gc', 10, 0); setText('gc', 'hi'); setStyle('gc', { bold: true }); setFill('gc', '#112233'); setStroke('gc', null);`,
      els2,
      canvas,
    )
    expect(r.error).toBeUndefined()
    expect(r.ops).toEqual([
      { id: 'gc', x: 220, y: 120, w: 80, h: 40, rotation: 0, groupId: 'grp1' },
    ])
    expect(r.edits).toEqual([
      { kind: 'text', id: 'gc', paragraphs: [{ runs: [{ text: 'hi' }] }], groupId: 'grp1' },
      { kind: 'style', id: 'gc', style: { bold: true }, groupId: 'grp1' },
      { kind: 'fill', id: 'gc', fill: '#112233', groupId: 'grp1' },
      { kind: 'stroke', id: 'gc', stroke: null, groupId: 'grp1' },
    ])
  })

  it('locked elements are rejected by all write primitives (including setBox)', () => {
    for (const call of [
      `setBox('deco', { x: 0 })`,
      `moveBy('deco', 1, 1)`,
      `setText('deco', 'x')`,
      `setStyle('deco', { bold: true })`,
      `setFill('deco', 'none')`,
      `setStroke('deco', null)`,
    ]) {
      const r = runLayoutScript(call, els, canvas)
      expect(r.error).toContain('read-only')
    }
  })

  it('unknown id errors', () => {
    expect(runLayoutScript(`setText('nope', 'x')`, els, canvas).error).toContain('nope')
  })

  it('allows legitimate layout computation: filtering, regex, callbacks, loops, Math, computed properties', () => {
    const r = runLayoutScript(
      `
        const cards = els.filter(e => /^(a|b)$/.test(e.id));
        const key = 'x';
        const left = Math.min(...cards.map(e => e[key]));
        for (let i = 0; i < cards.length; i++) {
          setBox(cards[i].id, { x: left + i * 340, w: 320 });
        }
        log('cards', cards.length);
        return { count: cards.length, left };
      `,
      els,
      canvas,
    )
    expect(r.error).toBeUndefined()
    expect(r.ops).toEqual([
      { id: 'a', x: 100, y: 200, w: 320, h: 100, rotation: 0 },
      { id: 'b', x: 440, y: 210, w: 320, h: 300, rotation: 0 },
    ])
    expect(r.logs).toEqual(['cards 2'])
    expect(r.returned).toBe('{"count":2,"left":100}')
  })
})

describe('runLayoutScript security boundary', () => {
  const els: LayoutScriptElement[] = [
    { id: 'a', type: 'shape', text: 'title', x: 100, y: 200, w: 300, h: 100, rotation: 0 },
  ]
  const canvas = { w: 1280, h: 720 }

  it.each([
    ['dynamic constructor spelling', `const p = ['con', 'structor'].join(''); return [][p];`],
    [
      'element constructor chain',
      `return els[0]['con' + 'structor']['constructor']('return globalThis')();`,
    ],
    ['prototype chain', `return ({})['__proto__']['constructor'];`],
    ['globalThis', `return globalThis;`],
    ['fetch', `return fetch('https://example.com');`],
    ['process', `return process.env;`],
    ['dynamic import', `return import('node:fs');`],
    ['Function constructor', `return new Function('return 1')();`],
  ])('rejects %s and collects no operations', (_name, code) => {
    const r = runLayoutScript(`setBox('a', { x: 999 }); ${code}`, els, canvas)
    expect(r.error).toBeTruthy()
    expect(r.ops).toEqual([])
    expect(r.edits).toEqual([])
  })

  it('computed properties may only read own JSON data or explicitly allowed methods', () => {
    const ok = runLayoutScript(`const p = 'x'; setBox('a', { x: els[0][p] + 10 });`, els, canvas)
    expect(ok.error).toBeUndefined()
    expect(ok.ops[0].x).toBe(110)

    const absentOwnProperty = runLayoutScript(
      `return els[0][['con', 'structor'].join('')];`,
      els,
      canvas,
    )
    expect(absentOwnProperty.error).toBeUndefined()
    expect(absentOwnProperty.returned).toBeUndefined()

    for (const code of [
      `return [][['__', 'proto__'].join('')];`,
      `return /x/[['con', 'structor'].join('')];`,
    ])
      expect(runLayoutScript(code, els, canvas).error).toBeTruthy()
  })

  it('bounds catastrophic-backtracking regex instead of freezing', () => {
    const r = runLayoutScript(
      `
        let long = '';
        for (let i = 0; i < 200; i++) long += 'aaaaaaaaaa';
        return /(a+)+$/.test(long + 'b');
      `,
      els,
      canvas,
    )
    expect(r.error).toMatch(/execution budget/)
    expect(r.ops).toEqual([])
  })

  it('supports the common regex subset with bounded execution', () => {
    const cases: Array<[string, string]> = [
      [`return /^(shape|text)$/.test(els[0].type);`, 'true'],
      [`return /TITLE/i.test(els[0].text);`, 'true'],
      [`return /\\d{2,4}/.test('year 2026');`, 'true'],
      [`return /[a-z]+_[0-9]+/.test('shape_12');`, 'true'],
      [`return /^\\s*$/.test('   ');`, 'true'],
      [`return /b.t/s.test('b\\nt') && /^t/m.test('a\\nt');`, 'true'],
      [`return /(a?)*$/.test('b');`, 'true'],
      [`return /xyz/.test(els[0].text);`, 'false'],
    ]
    for (const [code, expected] of cases) {
      const r = runLayoutScript(code, els, canvas)
      expect(r.error).toBeUndefined()
      expect(r.returned).toBe(expected)
    }
  })

  it('scans long text and honors start anchors without exhausting the budget', () => {
    const r = runLayoutScript(
      `
        let long = '';
        for (let i = 0; i < 500; i++) long += 'lorem ipsum ';
        return /needle|dolor$/.test(long + 'needle') && /^lorem/.test(long) && !/^needle/.test(long);
      `,
      els,
      canvas,
    )
    expect(r.error).toBeUndefined()
    expect(r.returned).toBe('true')
  })

  it('rejects regex features that cannot be bounded', () => {
    for (const code of [
      `return /(a)\\1/.test('aa');`,
      `return /a(?=b)/.test('ab');`,
      `return /(?<=a)b/.test('ab');`,
      `return /\\p{L}/u.test('a');`,
      `return /a/y.test('a');`,
    ])
      expect(runLayoutScript(code, els, canvas).error).toBeTruthy()
  })
})

// ── Review round 2: per-run regex budget, bounded serializer, property counter ──

describe('runLayoutScript regex budget (per-run total owned by the interpreter)', () => {
  const canvas = { w: 1280, h: 720 }
  // 230 chars with no 'z': every start position of /zzzz\d+q/ fails after ~3 steps
  const LOREM_230 = (
    'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor ' +
    'incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud ' +
    'exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat duiaut'
  ).slice(0, 230)
  const bigDeck = (count: number): LayoutScriptElement[] =>
    Array.from({ length: count }, (_, i) => ({
      id: `e${i}`,
      type: 'shape',
      text: LOREM_230,
      x: i,
      y: 0,
      w: 10,
      h: 10,
      rotation: 0,
    }))

  it('hoisted regex over a 2000-element × 230-char deck still completes (maintainer probe)', () => {
    // 2000 test() calls cost ≈1.39M matcher steps in total: re-arming a budget per
    // call (main) or per compiled pattern aborted this mid-loop; the per-run pool
    // keeps the normal "hoisted pattern over a large deck" script working.
    const r = runLayoutScript(
      `const re = /zzzz\\d+q/; let hits = 0; for (const e of els) { if (re.test(e.text)) hits += 1; } return hits;`,
      bigDeck(2000),
      canvas,
    )
    expect(r.error).toBeUndefined()
    expect(r.returned).toBe('0')
  })

  it('regex literals re-created every iteration draw from the same pool (stacking aborts)', () => {
    // A literal inside the loop mints a fresh RegexValue per iteration; each
    // /^(a|aa)+$/ call on 22 a's + 'b' burns ≈0.5M steps. Per-call (main) or
    // per-literal budgets let all 12 calls run; the per-run pool aborts the run
    // (9th call crosses the 4M-step total).
    const nearBudget = `${'a'.repeat(22)}b`
    const r = runLayoutScript(
      `let hits = 0; for (let i = 0; i < 12; i++) { if (/^(a|aa)+$/.test('${nearBudget}')) hits += 1; } return hits;`,
      [{ id: 'a', type: 'shape', text: 'title', x: 0, y: 0, w: 10, h: 10, rotation: 0 }],
      canvas,
    )
    expect(r.error).toMatch(/execution budget/)
    expect(r.ops).toEqual([])
  })

  it('repeated calls of one hoisted near-budget pattern share the pool too', () => {
    const nearBudget = `${'a'.repeat(22)}b`
    const r = runLayoutScript(
      `const re = /^(a|aa)+$/; let hits = 0; for (let i = 0; i < 12; i++) { if (re.test('${nearBudget}')) hits += 1; } return hits;`,
      [{ id: 'a', type: 'shape', text: 'title', x: 0, y: 0, w: 10, h: 10, rotation: 0 }],
      canvas,
    )
    expect(r.error).toMatch(/execution budget/)
  })

  it('a budget-exhausting pattern still aborts the whole script', () => {
    const r = runLayoutScript(
      `let long = ''; for (let i = 0; i < 200; i++) long += 'aaaaaaaaaa'; return /(a+)+$/.test(long + 'b');`,
      [{ id: 'a', type: 'shape', text: 'title', x: 0, y: 0, w: 10, h: 10, rotation: 0 }],
      canvas,
    )
    expect(r.error).toMatch(/execution budget/)
    expect(r.ops).toEqual([])
  })
})

describe('runLayoutScript output caps (size-bounded serializer)', () => {
  const els: LayoutScriptElement[] = [
    { id: 'a', type: 'shape', text: 'title', x: 100, y: 200, w: 300, h: 100, rotation: 0 },
  ]
  const canvas = { w: 1280, h: 720 }
  const ENTRY_CAP = 4_000
  const TRUNCATED = '…(truncated)'

  it('log() of a {x:o, y:o} doubling chain truncates without materializing it', () => {
    const r = runLayoutScript(
      `let o = {}; for (let i = 0; i < 20; i++) o = { x: o, y: o }; log(o); return 'ok';`,
      els,
      canvas,
    )
    expect(r.error).toBeUndefined()
    expect(r.logs).toHaveLength(1)
    expect(r.logs[0]!.length).toBeLessThanOrEqual(ENTRY_CAP + TRUNCATED.length)
    expect(r.logs[0]).toContain(TRUNCATED)
    expect(r.returned).toBe('ok')
  })

  it('returning a doubling chain is bounded at the entry cap (no 25 MB stringify)', () => {
    const r = runLayoutScript(
      `let o = {}; for (let i = 0; i < 20; i++) o = { x: o, y: o }; return o;`,
      els,
      canvas,
    )
    expect(r.error).toBeUndefined()
    expect(r.returned!.length).toBeLessThanOrEqual(ENTRY_CAP + TRUNCATED.length)
    expect(r.returned).toContain(TRUNCATED)
  })

  it('a cyclic return value depth-caps instead of hanging or erroring', () => {
    const r = runLayoutScript(`const o = {}; o.self = o; return o;`, els, canvas)
    expect(r.error).toBeUndefined()
    expect(r.returned!.length).toBeLessThanOrEqual(ENTRY_CAP + TRUNCATED.length)
    expect(r.returned).toContain(TRUNCATED)
  })

  it('per-entry (4k) and total (20k) log caps hold', () => {
    const longEntry = JSON.stringify('x'.repeat(5000))
    const r = runLayoutScript(
      `for (let i = 0; i < 60; i++) log(${longEntry}); return 'ok';`,
      els,
      canvas,
    )
    expect(r.error).toBeUndefined()
    expect(r.logs.length).toBeGreaterThan(0)
    expect(r.logs.length).toBeLessThanOrEqual(50)
    for (const entry of r.logs)
      expect(entry.length).toBeLessThanOrEqual(ENTRY_CAP + TRUNCATED.length)
    expect(r.logs[0]).toContain(TRUNCATED)
    const total = r.logs.reduce((sum, entry) => sum + entry.length, 0)
    expect(total).toBeLessThanOrEqual(20_000 + ENTRY_CAP + TRUNCATED.length)

    const countCapped = runLayoutScript(
      `for (let i = 0; i < 60; i++) log('entry-' + i); return 'ok';`,
      els,
      canvas,
    )
    expect(countCapped.logs).toHaveLength(50)
  })

  it('a long string return value is truncated at the entry cap', () => {
    const r = runLayoutScript(`return ${JSON.stringify('y'.repeat(100000))};`, els, canvas)
    expect(r.error).toBeUndefined()
    expect(r.returned!.length).toBe(ENTRY_CAP + TRUNCATED.length)
    expect(r.returned!.endsWith(TRUNCATED)).toBe(true)
  })

  it('the bounded serializer matches JSON.stringify for ordinary data', () => {
    const cases: Array<[string, string]> = [
      [`return { count: 2, left: 100 };`, '{"count":2,"left":100}'],
      [`return { s: 'a"b\\\\c\\nd\\té' };`, '{"s":"a\\"b\\\\c\\nd\\té"}'],
      [`return { a: 1, b: els[0].missing };`, '{"a":1}'],
      [`return [1, 'two', null, true];`, '[1,"two",null,true]'],
      [`return { n: 0.5, big: 1e21 };`, '{"n":0.5,"big":1e+21}'],
    ]
    for (const [code, expected] of cases) {
      const r = runLayoutScript(code, els, canvas)
      expect(r.error).toBeUndefined()
      expect(r.returned).toBe(expected)
    }
  })

  it('error messages are capped at the entry cap too', () => {
    const r = runLayoutScript(`throw ${JSON.stringify('e'.repeat(100000))};`, els, canvas)
    expect(r.error!.length).toBe(ENTRY_CAP + TRUNCATED.length)
    expect(r.error!.endsWith(TRUNCATED)).toBe(true)
  })
})

describe('runLayoutScript object property cap (O(1) counter)', () => {
  const els: LayoutScriptElement[] = [
    { id: 'a', type: 'shape', text: 'title', x: 100, y: 200, w: 300, h: 100, rotation: 0 },
  ]
  const canvas = { w: 1280, h: 720 }
  const objectLiteral = (props: number): string =>
    'return {' + Array.from({ length: props }, (_, i) => `p${i}: ${i}`).join(',') + '}'

  it('object literals above 10,000 properties are rejected', () => {
    const r = runLayoutScript(objectLiteral(10_001), els, canvas)
    expect(r.error).toContain('too large')
    expect(r.ops).toEqual([])
    expect(r.edits).toEqual([])
  })

  it('a 10,000-property literal still builds (counter stays O(1) per property)', () => {
    const r = runLayoutScript(objectLiteral(10_000), els, canvas)
    expect(r.error).toBeUndefined()
    expect(r.returned).toBeTruthy()
  })

  it('spreading an object with more than 10,000 properties is rejected', () => {
    const bigJson =
      '{"p0":0' + Array.from({ length: 10_000 }, (_, i) => `,"p${i + 1}":${i + 1}`).join('') + '}'
    const r = runLayoutScript(
      `const o = JSON.parse(${JSON.stringify(bigJson)}); return { ...o };`,
      els,
      canvas,
    )
    expect(r.error).toContain('too large')
  })

  it('duplicate keys do not trip the property counter', () => {
    const r = runLayoutScript(`return { a: 1, a: 2, a: 3 };`, els, canvas)
    expect(r.error).toBeUndefined()
    expect(r.returned).toBe('{"a":3}')
  })
})

// ── execute_slide_script tool chain ────────────────────────────

describe('execute_slide_script tool', () => {
  let slide: RenderSlide
  let applied: RenderSlide | null

  const access = (): DeckAccess => ({
    getSlides: () => [slide],
    getCurrent: () => 0,
    getSelectedIds: () => [],
    applySlide: (_i, updated) => {
      applied = updated
      slide = updated
    },
    applyDeck: (all) => {
      slide = all[0]!
    },
    fitWidthPx: 1280,
  })

  beforeEach(() => {
    applied = null
    slide = slideOf([
      textNode('t1', box(100, 100, 400, 100), 'Title'),
      textNode('t2', box(120, 300, 400, 100), 'Subtitle'),
    ])
    // Simulate the main process: applyEditScript applies boxes/text onto the render
    // tree and returns the rebuilt slide (fidelity mapping is covered by the
    // main-side script-map tests against a real deck)
    ;(globalThis as any).window = {
      slidesApi: {
        applyEditScript: vi.fn(async (op: any) => {
          let nodes = slide.nodes.map((n) => {
            const item = op.boxes.find((b: any) => b.id === n.sourceId)
            if (!item) return n
            return { ...n, box: box(item.x, item.y, item.w, item.h, item.rotation) }
          })
          for (const e of op.edits) {
            if (e.kind !== 'text') continue
            nodes = nodes.map((n) => {
              if (n.sourceId !== e.id) return n
              const sn = n as ShapeRenderNode
              const lines = e.paragraphs.map((p: any) => ({
                runs: p.runs.map((r: any) => ({
                  text: r.text,
                  x: 8,
                  baselineY: 20,
                  fontFamily: r.fontFamily ?? 'Arial',
                  fontSizePx: r.fontSize ? Math.round((r.fontSize * 96) / 72) : 24,
                  color: r.color ?? '#000000',
                  bold: !!r.bold,
                  italic: !!r.italic,
                  underline: !!r.underline,
                  widthPx: String(r.text).length * 12,
                })),
                top: 0,
                height: 28,
              }))
              return { ...sn, text: { ...sn.text!, lines } }
            })
          }
          return { slide: { ...slide, nodes } }
        }),
        editText: vi.fn(async (op: any) => {
          const nodes = slide.nodes.map((n) => {
            if (n.sourceId !== op.sourceId) return n
            return n
          })
          return { ...slide, nodes }
        }),
      },
    }
  })

  it('mixed script: one IPC carrying merged geometry + ordered edits', async () => {
    const skill = createSlidesSkill(access())
    const r = await skill.executeTool({
      id: '1',
      name: 'execute_slide_script',
      input: {
        slideIndex: 0,
        code: `
          moveBy('t1', -20, 0);
          setText('t2', 'New subtitle');
          setStyle('t1', { color: '#1a73e8', bold: true });
          setFill('t2', '#f8fafc');
          setStroke('t2', { color: '#e2e8f0', widthPt: 1 });
          return 'ok';
        `,
      },
    } as any)
    expect((r as any).isError).toBeFalsy()
    expect(r.mutated).toBe(true)
    const api = (globalThis as any).window.slidesApi
    // The whole script is exactly one IPC / one transaction
    expect(api.applyEditScript).toHaveBeenCalledTimes(1)
    const payload = api.applyEditScript.mock.calls[0][0]
    expect(payload.slideIndex).toBe(0)
    expect(payload.fitWidthPx).toBe(1280)
    // moveBy is based on real coordinates (100-20=80); geometry merged per element
    expect(payload.boxes).toEqual([{ id: 't1', x: 80, y: 100, w: 400, h: 100, rotation: 0 }])
    // Edits keep script call order; the style patch passes through raw (the main
    // process maps it onto setFont/setParagraphFormat ops)
    expect(payload.edits).toEqual([
      { kind: 'text', id: 't2', paragraphs: [{ runs: [{ text: 'New subtitle' }] }] },
      { kind: 'style', id: 't1', style: { color: '#1a73e8', bold: true } },
      { kind: 'fill', id: 't2', fill: '#f8fafc' },
      { kind: 'stroke', id: 't2', stroke: { color: '#e2e8f0', widthPt: 1 } },
    ])
    expect(r.output).toContain('layout 1 element(s)')
    expect(r.output).toContain('text 1 item(s)')
    expect(r.output).toContain('style 1 item(s)')
    expect(r.output).toContain('fill 1 item(s)')
    expect(r.output).toContain('stroke 1 item(s)')
    expect(r.output).toContain('<layout-audit>')
    expect(applied).not.toBeNull()
  })

  it('setStyle after setText in the same script rides the same transaction in order', async () => {
    const skill = createSlidesSkill(access())
    await skill.executeTool({
      id: '2',
      name: 'execute_slide_script',
      input: {
        slideIndex: 0,
        code: `setText('t1', 'Edited title'); setStyle('t1', { fontSize: 32 });`,
      },
    } as any)
    const api = (globalThis as any).window.slidesApi
    const payload = api.applyEditScript.mock.calls[0][0]
    expect(payload.edits.map((e: any) => e.kind)).toEqual(['text', 'style'])
    // Ops apply sequentially on the model, so the style lands on the new text
    expect(payload.edits[1]).toMatchObject({ id: 't1', style: { fontSize: 32 } })
  })

  it('a failing op aborts the whole script: guided error, renderer state untouched', async () => {
    const api = (globalThis as any).window.slidesApi
    api.applyEditScript.mockResolvedValueOnce({
      error:
        'op "setFill": element "t1" does not support fill. Nothing was applied (atomic) — fix this op and resend the whole transaction.',
    })
    const skill = createSlidesSkill(access())
    const r = await skill.executeTool({
      id: '3',
      name: 'execute_slide_script',
      input: {
        slideIndex: 0,
        code: `moveBy('t1', -20, 0); setFill('t1', '#ff0000'); setText('t2', 'never applied');`,
      },
    } as any)
    expect((r as any).isError).toBe(true)
    expect(r.mutated).toBe(false)
    expect(r.output).toContain('setFill')
    expect(r.output).toContain('Nothing was applied')
    expect(r.output).toContain('unchanged')
    // Guided: the failure lists the ids that do exist on the page
    expect(r.output).toContain('t1')
    expect(r.output).toContain('t2')
    // The executor rolled back main-side; the renderer was never touched mid-dispatch
    expect(applied).toBeNull()
    expect(slide.nodes[0]!.box.x).toBe(100)
  })

  it('script sandbox error (e.g. locked/in-group) → isError and nothing is applied', async () => {
    const skill = createSlidesSkill(access())
    const r = await skill.executeTool({
      id: '5',
      name: 'execute_slide_script',
      input: { slideIndex: 0, code: `setText('missing', 'x');` },
    } as any)
    expect((r as any).isError).toBe(true)
    expect(applied).toBeNull()
    expect((globalThis as any).window.slidesApi.applyEditScript).not.toHaveBeenCalled()
  })

  it('group child: setBox stays absolute px in the payload (apply-time conversion) and setText carries groupId', async () => {
    slide = slideOf([
      groupNode('g1', box(200, 100, 400, 300), [textNode('c1', box(10, 20, 50, 30), 'Member')]),
    ])
    const skill = createSlidesSkill(access())
    const r = await skill.executeTool({
      id: '7',
      name: 'execute_slide_script',
      input: { slideIndex: 0, code: `setBox('c1', { x: 240 }); setText('c1', 'hi');` },
    } as any)
    expect((r as any).isError).toBeFalsy()
    expect(r.mutated).toBe(true)
    const payload = (globalThis as any).window.slidesApi.applyEditScript.mock.calls[0][0]
    // els shows c1 at absolute (210,120); the box travels absolute — the setTransform
    // op converts against the group's live state at apply time (a group moved earlier
    // in the same script must not double-shift its children)
    expect(payload.boxes).toEqual([
      { id: 'c1', groupId: 'g1', x: 240, y: 120, w: 50, h: 30, rotation: 0 },
    ])
    expect(payload.edits[0]).toMatchObject({ kind: 'text', id: 'c1', groupId: 'g1' })
  })

  it('legacy name execute_layout_script still works (alias), new primitives also take effect', async () => {
    const skill = createSlidesSkill(access())
    const r = await skill.executeTool({
      id: '6',
      name: 'execute_layout_script',
      input: { slideIndex: 0, code: `moveBy('t1', 0, 50); setStyle('t1', { bold: true });` },
    } as any)
    expect((r as any).isError).toBeFalsy()
    expect(r.mutated).toBe(true)
    expect(r.output).toContain('layout 1 element(s)')
    expect(r.output).toContain('style 1 item(s)')
  })
})
