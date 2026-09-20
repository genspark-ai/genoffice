import type { DiagramResult } from './diagrams'

export const WAVEDROM_LANGUAGE = 'wavedrom'

export const WAVEDROM_TEMPLATE = [
  '```wavedrom',
  '{ signal: [',
  "  { name: 'clk',  wave: 'p.....' },",
  "  { name: 'data', wave: 'x.345x', data: ['head', 'body', 'tail'] },",
  "  { name: 'req',  wave: '0.1..0' },",
  ']}',
  '```',
].join('\n')

type WaveDromModule = typeof import('wavedrom')
type Json5Module = typeof import('json5')

let loading: Promise<[WaveDromModule, Json5Module]> | null = null

/** Loaded on first use so documents without timing diagrams never pay for the bundle */
function loadWavedrom() {
  loading ??= Promise.all([import('wavedrom'), import('json5')]).then(([wd, json5]) => [
    wd,
    json5.default ?? json5,
  ])
  return loading
}

let renderSeq = 0

export async function renderWavedrom(
  source: string,
  deps?: { load?: () => Promise<[WaveDromModule, Json5Module]> },
): Promise<DiagramResult> {
  const preError = prevalidateWaveSource(source)
  if (preError) return { ok: false, error: preError }
  let wavedrom: WaveDromModule
  let json5: Json5Module
  try {
    ;[wavedrom, json5] = await (deps?.load ?? loadWavedrom)()
  } catch (err) {
    return {
      ok: false,
      error: `failed to load the wavedrom renderer: ${err instanceof Error ? err.message : String(err)}`,
    }
  }
  let parsed: unknown
  try {
    // WaveJSON is JS-object-literal flavoured (unquoted keys, single quotes, comments)
    parsed = json5.parse(source)
  } catch (err) {
    return { ok: false, error: err instanceof Error ? err.message : String(err) }
  }
  if (!isWaveJson(parsed)) {
    return { ok: false, error: 'expected an object with a "signal", "assign" or "reg" array' }
  }
  try {
    const tree = wavedrom.renderAny(++renderSeq, parsed, wavedrom.waveSkin)
    if (tree[0] !== 'svg') return { ok: false, error: 'nothing to draw' }
    return { ok: true, svg: scopeSkinStyles(wavedrom.onml.stringify(tree)) }
  } catch (err) {
    return { ok: false, error: err instanceof Error ? err.message : String(err) }
  }
}

/** Actionable pre-check before paying for the bundle; returns an error message or null */
export function prevalidateWaveSource(source: unknown): string | null {
  if (typeof source !== 'string') {
    return `expected WaveJSON text, got ${typeof source}: provide an object like { signal: [...] }`
  }
  if (!source.trim()) {
    return 'wavedrom source is empty: expected WaveJSON like { signal: [...] }'
  }
  const trimmed = source.trimStart()
  if (!trimmed.startsWith('{')) {
    const preview = trimmed.slice(0, 24)
    return `expected a WaveJSON object starting with "{", got "${preview}": provide an object like { signal: [...] }`
  }
  return null
}

export function isWaveJson(value: unknown): value is Record<string, unknown> {
  if (typeof value !== 'object' || value === null || Array.isArray(value)) return false
  const v = value as Record<string, unknown>
  let found = false
  for (const key of ['signal', 'assign', 'reg'] as const) {
    const arr = v[key]
    if (arr === undefined) continue
    if (!Array.isArray(arr)) return false
    if (arr.length === 0) continue
    if (
      !arr.every((entry) => typeof entry === 'object' && entry !== null && !Array.isArray(entry))
    ) {
      return false
    }
    found = true
  }
  return found
}

/**
 * The skin's inline stylesheet uses bare selectors (`text`, `.h1`, `.error`)
 * that an inline <svg> would apply to the whole page. Shapes are drawn via
 * <use>, whose shadow tree cannot see outer ancestors, so an ancestor prefix
 * would not reach them: namespace the class names instead.
 */
export function scopeSkinStyles(svg: string): string {
  return svg
    .replace(/<style([^>]*)>([\s\S]*?)<\/style>/g, (_m, attrs: string, css: string) => {
      const scoped = css.replace(/(^|\})\s*([^{}]+)\{/g, (_r, sep: string, selectors: string) => {
        const list = selectors
          .split(',')
          .map((s) => s.trim())
          .map((s) =>
            s.startsWith('.')
              ? s.replace(/\.([\w-]+)/g, (m, name: string) =>
                  name.startsWith('wd-') ? m : `.wd-${name}`,
                )
              : `svg.WaveDrom ${s}`,
          )
          .join(',')
        return `${sep}${list}{`
      })
      return `<style${attrs}>${scoped}</style>`
    })
    .replace(/\bclass="([^"]*)"/g, (m, names: string) =>
      names === 'WaveDrom'
        ? m
        : `class="${names
            .split(/\s+/)
            .filter(Boolean)
            .map((n) => (n.startsWith('wd-') ? n : `wd-${n}`))
            .join(' ')}"`,
    )
}
