// Face classes the line metrics branch on, and the stand-in faces the emitters name:
// one place for the name tests that used to sit next to their callers in
// line-metrics.ts, doc-style-css.ts and the editor emitters.
import { resolveDeclaredFace } from '../../font-resolve'

/** MS Gothic / MS Mincho as a chain head: a fixed 0.5 em Latin advance */
export const isMsFixedPitchHead = (head: string): boolean =>
  /^'?(ms (gothic|mincho)|ＭＳ (ゴシック|\u660e\u671d))'?$/i.test(head.trim())

/** Arial Unicode MS (lower-cased key): its list markers are unprobed */
export const isArialUnicodeFace = (key: string): boolean => /^arial unicode/.test(key)

export const isSegoeUiEmojiFace = (family: string): boolean => /segoe ui emoji/i.test(family)

/** Calibri and its metric twin: they lack a few WGL4 shapes */
export const isCalibriClassFace = (face: string): boolean => /calibri|carlito/i.test(face)

/** Latin text faces whose lead shift is modelled; other heads keep the CSS centre until a corpus run adopts them */
export const hasModelledLineBox = (face: string): boolean =>
  /aptos|calibri|carlito|cambria|caladea|times|liberation serif|\barial\b(?! unicode)|liberation sans|georgia|verdana|reference sans|tahoma|courier|trebuchet/.test(
    face,
  )

/** Times New Roman and its twins: a scaled run takes the line factor without the line gap */
export const isTimesClassFace = (face: string): boolean =>
  /^times(\s|$)|times new roman|liberation serif/i.test(face)

/** the bundled Segoe UI Symbol stand-in, as a CSS family */
export const SEGOE_SYMBOL_CSS = "'Segoe UI Symbol GO'"

/** the face whose only glyph is the balanced-width space, as a CSS family */
export const BALANCED_SPACE_CSS = "'GenOffice Balanced Space'"

/** family name of the blank PUA face in fonts.css */
export const PUA_BLANK_FAMILY = 'GenOffice PUA Blank'

/** metrics of the face a declared name resolves to */
export const declaredFaceMetrics = (font: string) => resolveDeclaredFace(font)?.metrics
