// Which face of a CSS chain the canvas really measures while web faces are still loading.
import { GENERIC_FAMILIES, isBundledFont, isFontAvailable } from './availability'

/**
 * Bundled web face that leads a chain but has not loaded yet: canvas measures
 * the next installed family instead (an unloaded font-display:block face is
 * skipped), so the strut would carry e.g. Apple SD Gothic Neo's 0.90/0.30
 * under GenOffice Sans KR spans (1.16/0.29) and every hangul line unions
 * ~1.5px past its line-height (+1pt per Korean table row vs Word).
 */
export function unloadedBundledPrimary(chain: string): string | null {
  const fonts = typeof document === 'undefined' ? undefined : document.fonts
  if (!fonts?.check) return null
  // the chain's primary is its first face covering U+0020; canvas skips any
  // web face still loading, so every loaded one is walked past
  for (const raw of chain.split(',')) {
    const family = raw.trim().replace(/^["']|["']$/g, '')
    if (!family || GENERIC_FAMILIES.has(family.toLowerCase())) return null
    if (isWebFace(family)) {
      try {
        if (!fonts.check(`16px "${family}"`)) return family
      } catch {
        return null
      }
    } else if (isFontAvailable(family)) return null
  }
  return null
}

/**
 * Every fonts.css face: bundled stand-ins plus the metric/range aliases. An
 * unloaded 'MS Mincho GO' leaves canvas measuring the Noto tail (1.15/0.29em)
 * while layout later resolves the loaded alias (0.69/0.25em) as primary.
 */
export function isWebFace(family: string): boolean {
  return isBundledFont(family) || / GO$/.test(family) || /^GenOffice /.test(family)
}

/** bundled web face a chain resolves to before any installed family or generic */
export function bundledPrimary(chain: string): string | null {
  for (const raw of chain.split(',')) {
    const family = raw.trim().replace(/^["']|["']$/g, '')
    if (!family || GENERIC_FAMILIES.has(family.toLowerCase())) return null
    if (isBundledFont(family)) return family
    if (isFontAvailable(family)) return null
  }
  return null
}
