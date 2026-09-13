/// Document theme writer: rewrites xl/theme/theme1.xml's <a:clrScheme> and/or
/// <a:fontScheme> in place. Styles referencing theme slots (color theme=,
/// font scheme=) follow the new theme in Excel; explicit rgb stays verbatim.

export class ThemeStateError extends Error {}

export interface WorkbookThemeState {
  /// #RRGGBB values in theme index order [lt1, dk1, lt2, dk2, accent1-6,
  /// hlink, folHlink].
  readonly colors?: { readonly name: string; readonly values: readonly string[] } | undefined
  readonly fonts?:
    { readonly name: string; readonly major: string; readonly minor: string } | undefined
}

/// clrScheme document order, paired with the slot's position in the theme
/// index order the renderer uses (light/dark pairs are swapped).
const SCHEME_SLOTS: readonly (readonly [string, number])[] = [
  ['dk1', 1],
  ['lt1', 0],
  ['dk2', 3],
  ['lt2', 2],
  ['accent1', 4],
  ['accent2', 5],
  ['accent3', 6],
  ['accent4', 7],
  ['accent5', 8],
  ['accent6', 9],
  ['hlink', 10],
  ['folHlink', 11],
]

function escapeAttr(value: string): string {
  return value
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
}

export function applyThemeState(themeXml: string, state: WorkbookThemeState): string {
  let xml = themeXml
  if (state.colors !== undefined) {
    const { name, values } = state.colors
    if (values.length !== 12) throw new ThemeStateError('A theme palette needs 12 colors.')
    for (const [slot, index] of SCHEME_SLOTS) {
      const hex = (values[index] ?? '').replace(/^#/, '').toUpperCase()
      if (!/^[0-9A-F]{6}$/.test(hex)) {
        throw new ThemeStateError(`Theme color ${slot} is not a valid #RRGGBB value.`)
      }
      const pattern = new RegExp(`<a:${slot}>[\\s\\S]*?</a:${slot}>`)
      if (!pattern.test(xml)) {
        throw new ThemeStateError(`The theme has no ${slot} color scheme slot.`)
      }
      xml = xml.replace(pattern, `<a:${slot}><a:srgbClr val="${hex}"/></a:${slot}>`)
    }
    xml = xml.replace(/(<a:clrScheme name=")[^"]*(")/, `$1${escapeAttr(name)}$2`)
  }
  if (state.fonts !== undefined) {
    const { name, major, minor } = state.fonts
    const majorPattern = /(<a:majorFont>[\s\S]*?<a:latin[^>]*typeface=")[^"]*(")/
    const minorPattern = /(<a:minorFont>[\s\S]*?<a:latin[^>]*typeface=")[^"]*(")/
    if (!majorPattern.test(xml) || !minorPattern.test(xml)) {
      throw new ThemeStateError('The theme has no font scheme to rewrite.')
    }
    xml = xml.replace(majorPattern, `$1${escapeAttr(major)}$2`)
    xml = xml.replace(minorPattern, `$1${escapeAttr(minor)}$2`)
    xml = xml.replace(/(<a:fontScheme name=")[^"]*(")/, `$1${escapeAttr(name)}$2`)
  }
  return xml
}
