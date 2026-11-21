import type { Editor } from '@tiptap/core'
import type {
  TextFxOwn,
  TextGlow,
  TextOutline,
  TextReflection,
  TextShadow,
} from '@genoffice/docx-engine'
import {
  glowAndShadowDecl,
  parseFxOwnAttr,
  parseGlowAttr,
  parseReflectionAttr,
  parseTextShadowAttr,
  reflectionDecl,
} from './text-effects'
import { parseTextOutlineAttr, textOutlineDecl } from './text-outline'

/** the effects of the Text Effects dialog; null is "none" */
export interface TextFx {
  outline: TextOutline | null
  shadow: TextShadow | null
  glow: TextGlow | null
  reflection: TextReflection | null
}

/** document colours offered for an outline or a glow (hex without '#') */
export const FX_COLORS = [
  '000000',
  'FFFFFF',
  '4472C4',
  'ED7D31',
  'A5A5A5',
  'FFC000',
  '5B9BD5',
  '70AD47',
  'FF0000',
] as const

export const OUTLINE_WIDTHS_PT = [0.75, 1, 1.5, 2.25] as const
export const GLOW_SIZES_PT = [5, 8, 11, 18] as const

/** Word's outer shadows thrown below the text: soft, 60% transparent black */
const shadowTowards = (dirDeg: number): TextShadow => ({
  color: '000000',
  alpha: 0.6,
  blurPt: 4,
  distPt: 3,
  dirDeg,
})

export const SHADOW_PRESETS = [
  { id: 'right', shadow: shadowTowards(45) },
  { id: 'below', shadow: shadowTowards(90) },
  { id: 'left', shadow: shadowTowards(135) },
] as const

/** the glow of a size in a colour, as transparent as Word's glow presets */
export const glowOf = (color: string, radiusPt: number): TextGlow => ({
  color,
  radiusPt,
  alpha: 0.6,
})

const reflectionOf = (startAlpha: number, endPos: number): TextReflection => ({
  startAlpha,
  endPos,
  blurPt: 0.5,
  distPt: 0,
})

/** Word's reflections that touch the text: how much of it they show */
export const REFLECTION_PRESETS = [
  { id: 'tight', reflection: reflectionOf(0.52, 0.35) },
  { id: 'half', reflection: reflectionOf(0.5, 0.55) },
  { id: 'full', reflection: reflectionOf(0.5, 0.9) },
] as const

const NO_FX: TextFx = { outline: null, shadow: null, glow: null, reflection: null }

/** ready-made looks: a text colour with the effects that go with it */
export const TEXT_FX_PRESETS: readonly { id: string; color: string; fx: TextFx }[] = [
  { id: 'black-shadow', color: '000000', fx: { ...NO_FX, shadow: shadowTowards(45) } },
  { id: 'blue-shadow', color: '4472C4', fx: { ...NO_FX, shadow: shadowTowards(45) } },
  {
    id: 'orange-outline',
    color: 'ED7D31',
    fx: { ...NO_FX, outline: { color: 'ED7D31', widthPt: 0.75 } },
  },
  {
    id: 'white-outline-glow',
    color: 'FFFFFF',
    fx: { ...NO_FX, outline: { color: '4472C4', widthPt: 0.75 }, glow: glowOf('4472C4', 5) },
  },
  {
    id: 'white-outline-shadow',
    color: 'FFFFFF',
    fx: { ...NO_FX, outline: { color: '5B9BD5', widthPt: 0.75 }, shadow: shadowTowards(45) },
  },
  {
    id: 'blue-reflection',
    color: '5B9BD5',
    fx: { ...NO_FX, reflection: REFLECTION_PRESETS[1].reflection },
  },
]

/** the effects the text at the selection shows */
export function textFxOf(editor: Editor): TextFx {
  const attrs = editor.getAttributes('docTextStyle')
  return {
    outline: attrs.textOutline ? (parseTextOutlineAttr(String(attrs.textOutline)) ?? null) : null,
    shadow: attrs.textShadow ? (parseTextShadowAttr(String(attrs.textShadow)) ?? null) : null,
    glow: attrs.glow ? (parseGlowAttr(String(attrs.glow)) ?? null) : null,
    reflection: attrs.textReflection
      ? (parseReflectionAttr(String(attrs.textReflection)) ?? null)
      : null,
  }
}

/** the colour the text at the selection is filled with, hex without '#'; null for the automatic colour */
export function textFillOf(editor: Editor): string | null {
  const color = editor.getAttributes('docTextStyle').color as unknown
  const hex = typeof color === 'string' ? color.replace(/^#/, '').toUpperCase() : ''
  return /^[0-9A-F]{6}$/.test(hex) ? hex : null
}

export const sameTextFx = (a: TextFx, b: TextFx): boolean =>
  JSON.stringify([a.outline, a.shadow, a.glow, a.reflection]) ===
  JSON.stringify([b.outline, b.shadow, b.glow, b.reflection])

/** CSS declarations that draw the effects, for the dialog's preview */
export function textFxDecls(fx: TextFx): string[] {
  const shadows = glowAndShadowDecl(fx.glow ?? undefined, fx.shadow ?? undefined)
  return [
    ...(fx.outline ? [textOutlineDecl(fx.outline)] : []),
    ...(shadows ? [shadows] : []),
    ...(fx.reflection ? [reflectionDecl(fx.reflection)] : []),
  ]
}

/**
 * Puts the effects on the selected text. All are written as the text's own:
 * one that is off is cleared on it, whatever its style says. `color` (hex
 * without '#') is a solid text fill: the text's colour, saved as Word saves
 * one. Without it the text keeps the fill it has.
 */
export function applyTextFx(editor: Editor, fx: TextFx, color?: string): boolean {
  const fill = color ?? parseFxOwnAttr(editor.getAttributes('docTextStyle').fxOwn)?.fill
  const own: TextFxOwn = {
    outline: fx.outline,
    shadow: fx.shadow,
    glow: fx.glow,
    reflection: fx.reflection,
    ...(fill !== undefined ? { fill } : {}),
  }
  return editor
    .chain()
    .focus()
    .setMark('docTextStyle', {
      textOutline: fx.outline ? JSON.stringify(fx.outline) : null,
      textShadow: fx.shadow ? JSON.stringify(fx.shadow) : null,
      glow: fx.glow ? JSON.stringify(fx.glow) : null,
      textReflection: fx.reflection ? JSON.stringify(fx.reflection) : null,
      fxOwn: JSON.stringify(own),
      ...(color ? { color } : {}),
    })
    .run()
}

/** puts one of the ready-made looks on the selected text; false when there is no such look */
export function applyTextFxPreset(editor: Editor, id: string | undefined): boolean {
  const preset = TEXT_FX_PRESETS.find((item) => item.id === id)
  return preset ? applyTextFx(editor, preset.fx, preset.color) : false
}
