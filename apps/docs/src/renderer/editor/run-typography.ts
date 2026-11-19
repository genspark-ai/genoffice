import type { RunOpenType } from '@genoffice/docx-engine'

/** the underline styles the Font dialog offers; a file's other w:u values still draw and save */
export const UNDERLINE_STYLES = ['single', 'double', 'thick', 'dotted', 'dash', 'wave'] as const

const LINE_CSS: Array<[RegExp, string]> = [
  [/^wav/, 'wavy'],
  [/double/i, 'double'],
  [/^dotted/, 'dotted'],
  // CSS has one dash pattern: long dashes and dot-dash mixes draw as dashes
  [/^(dash|dot)/, 'dashed'],
]

/** text-decoration declarations that draw a w:u style and colour; none for a plain single line */
export function underlineDecls(
  style: string | null | undefined,
  color: string | null | undefined,
): string[] {
  const decls: string[] = []
  const line = style ? LINE_CSS.find(([re]) => re.test(style))?.[1] : undefined
  if (line) decls.push(`text-decoration-style:${line}`)
  if (style && /thick|heavy/i.test(style)) decls.push('text-decoration-thickness:0.12em')
  if (color && /^[0-9A-Fa-f]{6}$/.test(color)) decls.push(`text-decoration-color:#${color}`)
  return decls
}

export function parseOtfAttr(value: unknown): RunOpenType | undefined {
  if (typeof value !== 'string' || !value) return undefined
  try {
    const parsed: unknown = JSON.parse(value)
    return parsed && typeof parsed === 'object' ? (parsed as RunOpenType) : undefined
  } catch {
    return undefined
  }
}

const LIGATURE_FEATURES: Array<[string, string]> = [
  ['standard', 'liga'],
  ['contextual', 'clig'],
  ['historical', 'hlig'],
  ['discretional', 'dlig'],
]

/** font-feature-settings for a run's w14 OpenType features; null when it sets none */
export function otfFeatureSettings(otf: RunOpenType | undefined): string | null {
  if (!otf) return null
  const features: string[] = []
  if (otf.ligatures !== undefined) {
    const val = otf.ligatures.toLowerCase()
    for (const [name, tag] of LIGATURE_FEATURES)
      features.push(`"${tag}" ${val === 'all' || val.includes(name) ? 1 : 0}`)
  }
  if (otf.numForm === 'lining') features.push('"lnum" 1')
  else if (otf.numForm === 'oldStyle') features.push('"onum" 1')
  if (otf.numSpacing === 'proportional') features.push('"pnum" 1')
  else if (otf.numSpacing === 'tabular') features.push('"tnum" 1')
  for (const id of otf.stylisticSets ?? [])
    features.push(`"ss${String(Math.round(id)).padStart(2, '0')}" 1`)
  if (otf.cntxtAlts !== undefined) features.push(`"calt" ${otf.cntxtAlts ? 1 : 0}`)
  return features.length > 0 ? features.join(',') : null
}

/** `existing` (a mark's otf JSON) with `patch` over it, keys in the order parse writes them */
export function mergeOtfAttr(existing: unknown, patch: RunOpenType): string {
  const all = { ...parseOtfAttr(existing), ...patch }
  const out: RunOpenType = {}
  if (all.ligatures !== undefined) out.ligatures = all.ligatures
  if (all.numForm !== undefined) out.numForm = all.numForm
  if (all.numSpacing !== undefined) out.numSpacing = all.numSpacing
  if (all.stylisticSets !== undefined) out.stylisticSets = all.stylisticSets
  if (all.cntxtAlts !== undefined) out.cntxtAlts = all.cntxtAlts
  return JSON.stringify(out)
}

/** text-emphasis declarations for a w:em value; none for 'none' */
export function emphasisDecls(em: string | null | undefined): string[] {
  if (!em || em === 'none') return []
  const shape = em === 'circle' ? 'open circle' : em === 'comma' ? 'filled sesame' : 'filled dot'
  // Word draws dot / underDot below the text; comma and circle go above
  const pos = em === 'comma' || em === 'circle' ? 'over' : 'under'
  return [`text-emphasis:${shape}`, `text-emphasis-position:${pos} right`]
}

/** `prop:value` declarations as a React style object */
export function declsToStyle(decls: string[]): Record<string, string> {
  return Object.fromEntries(
    decls.map((d) => {
      const at = d.indexOf(':')
      return [d.slice(0, at).replace(/-(\w)/g, (_, c: string) => c.toUpperCase()), d.slice(at + 1)]
    }),
  )
}
