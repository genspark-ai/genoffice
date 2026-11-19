/** w:shd hatch patterns in the order of Word's Style list (after the percentages) */
export const SHADING_HATCHES = [
  'horzStripe',
  'vertStripe',
  'reverseDiagStripe',
  'diagStripe',
  'horzCross',
  'diagCross',
  'thinHorzStripe',
  'thinVertStripe',
  'thinReverseDiagStripe',
  'thinDiagStripe',
  'thinHorzCross',
  'thinDiagCross',
] as const

const ANGLES: Record<string, number[]> = {
  HorzStripe: [0],
  VertStripe: [90],
  ReverseDiagStripe: [45],
  DiagStripe: [135],
  HorzCross: [0, 90],
  DiagCross: [45, 135],
}

/** a hatch as background-image layers over the fill (document colours); null for any other w:shd value */
export function shadingHatchCss(val: string, color?: string | null): string | null {
  const thin = val.startsWith('thin')
  const base = thin ? val.slice(4) : val.replace(/^./, (c) => c.toUpperCase())
  const angles = ANGLES[base]
  if (!angles) return null
  const ink = `#${color ?? '000000'}`
  const [on, period] = thin ? [1, 4] : [3, 6]
  return angles
    .map(
      (deg) =>
        `repeating-linear-gradient(${deg}deg,${ink} 0 ${on}px,transparent ${on}px ${period}px)`,
    )
    .join(',')
}
