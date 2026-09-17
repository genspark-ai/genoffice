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

export async function renderWavedrom(source: string): Promise<DiagramResult> {
  const [wavedrom, json5] = await loadWavedrom()
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

function isWaveJson(value: unknown): value is object {
  if (typeof value !== 'object' || value === null || Array.isArray(value)) return false
  const v = value as Record<string, unknown>
  return Array.isArray(v.signal) || Array.isArray(v.assign) || Array.isArray(v.reg)
}

/**
 * The skin's inline stylesheet uses bare selectors (`text`, `.h1`, `.error`)
 * that an inline <svg> would apply to the whole page. Shapes are drawn via
 * <use>, whose shadow tree cannot see outer ancestors, so an ancestor prefix
 * would not reach them: namespace the class names instead.
 */
function scopeSkinStyles(svg: string): string {
  return svg
    .replace(/<style([^>]*)>([\s\S]*?)<\/style>/, (_m, attrs: string, css: string) => {
      const scoped = css.replace(/(^|\})\s*([^{}]+)\{/g, (_r, sep: string, selectors: string) => {
        const list = selectors
          .split(',')
          .map((s) => s.trim())
          .map((s) =>
            s.startsWith('.') ? s.replace(/\.([\w-]+)/g, '.wd-$1') : `svg.WaveDrom ${s}`,
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
            .map((n) => `wd-${n}`)
            .join(' ')}"`,
    )
}
