/**
 * Which faces a run lays out with, answered once for the live marks, the
 * static DOM and the inline rules: the faces a run without its own w:rFonts
 * inherits from its paragraph style chain, and the CSS declarations a run's
 * declared slots turn into.
 */
import type { StyleDisplay } from '@genoffice/docx-engine'
import type { ListNumberingStorage } from '../editor/extensions'
import {
  cssCsFontFamily,
  cssFontFamily,
  cssRunFontFamily,
  cutAdjustedAlias,
  isEaLayoutFace,
  lineHeightFactor,
  runSlotFaces,
  wordSlotFaces,
} from '../line-metrics'

/** paragraph styleId -> inherited face (null = none declared on the chain) */
export type FaceOf = (styleId: unknown) => string | null

/** EA face a run without its own font inherits: paragraph style chain, then
 *  Normal, then docDefaults (an empty-theme-slot backfill is not a choice) */
export function inheritedEaFamily(
  styleId: unknown,
  storage: ListNumberingStorage | undefined,
): string | null {
  if (!storage) return null
  const ownEa = (d: StyleDisplay | undefined): string | null =>
    d?.font && !d.eaSlotEmpty && (d.font !== d.fontAscii || isEaLayoutFace(d.font)) ? d.font : null
  let id = typeof styleId === 'string' ? styleId : null
  for (let hops = 0; id && hops < 32; hops++) {
    const info = storage.styles?.get(id)
    const fam = ownEa(info?.display)
    if (fam) return fam
    id = info?.basedOn ?? null
  }
  for (const info of storage.styles?.values() ?? []) {
    if (info.isDefault && info.type === 'paragraph') {
      const fam = ownEa(info.display)
      if (fam) return fam
      break
    }
  }
  const dd = storage.docDefaults
  return dd?.eastAsiaFont && !dd.eaSlotEmpty ? dd.eastAsiaFont : null
}

/** ascii face a run without its own font inherits: paragraph style chain, Normal, docDefaults */
export function inheritedAsciiFamily(
  styleId: unknown,
  storage: ListNumberingStorage | undefined,
): string | null {
  if (!storage) return null
  const own = (d: StyleDisplay | undefined): string | null => d?.fontAscii ?? d?.font ?? null
  let id = typeof styleId === 'string' ? styleId : null
  for (let hops = 0; id && hops < 32; hops++) {
    const info = storage.styles?.get(id)
    const fam = own(info?.display)
    if (fam) return fam
    id = info?.basedOn ?? null
  }
  for (const info of storage.styles?.values() ?? []) {
    if (info.isDefault && info.type === 'paragraph')
      return own(info.display) ?? storage.docDefaults?.asciiFont ?? null
  }
  return storage.docDefaults?.asciiFont ?? null
}

/** memoised resolver over the live style table (one lookup per style per table) */
function memoised(
  getStorage: () => ListNumberingStorage | undefined,
  resolve: (styleId: unknown, storage: ListNumberingStorage | undefined) => string | null,
): FaceOf {
  const memo = new WeakMap<object, Map<string, string | null>>()
  return (styleId) => {
    const storage = getStorage()
    if (!storage?.styles) return resolve(styleId, storage)
    let m = memo.get(storage.styles)
    if (!m) memo.set(storage.styles, (m = new Map()))
    const k = typeof styleId === 'string' ? styleId : ''
    let fam = m.get(k)
    if (fam === undefined) m.set(k, (fam = resolve(styleId, storage)))
    return fam
  }
}

export function asciiFamilyResolver(getStorage: () => ListNumberingStorage | undefined): FaceOf {
  return memoised(getStorage, inheritedAsciiFamily)
}

export function eaFamilyResolver(getStorage: () => ListNumberingStorage | undefined): FaceOf {
  return memoised(getStorage, inheritedEaFamily)
}

export interface RunFaceAttrs {
  font?: string | null
  fontAscii?: string | null
  eastAsiaFont?: string | null
  eaSlotEmpty?: boolean | null
  rawRPr?: string | null
}

/**
 * The face declarations of a run with declared slots: the explicit East Asian
 * face (--doc-east-asian-font), the chain (font-family and its --doc-run-ff
 * twin, re-headed by the strut face in .doc-grid-strut paragraphs) and the
 * face's own line factor (--doc-run-lh, .doc-mixed-face). `cs` is the complex
 * script face when it applies to this run (the callers decide when).
 */
export function runFaceDecls(
  attrs: RunFaceAttrs,
  cs: string | null,
): { decls: string[]; family: string | null; cutAdjusted: boolean } {
  if (!attrs.font && !attrs.fontAscii && !cs) return { decls: [], family: null, cutAdjusted: false }
  const { ascii, eastAsia: ea, declaredEa } = runSlotFaces(attrs)
  const decls: string[] = []
  const explicitEa = declaredEa ? wordSlotFaces(null, declaredEa).eastAsia : null
  if (explicitEa && !attrs.eaSlotEmpty)
    decls.push(`--doc-east-asian-font:${cssFontFamily(explicitEa)}`)
  const family = cs
    ? cssCsFontFamily(cs, ascii ?? undefined, ea ?? undefined)
    : cssRunFontFamily(ascii, ea)
  decls.push(`font-family:${family}`, `--doc-run-ff:${family}`)
  const lh = Math.max(...[cs, ascii ?? ea].filter(Boolean).map((f) => lineHeightFactor(f!)))
  decls.push(`--doc-run-lh:${lh}`)
  return { decls, family, cutAdjusted: cutAdjustedAlias(family) }
}
