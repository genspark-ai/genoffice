/**
 * The overlay slot's shapes (K4): a screen-only layer the host draws over the
 * content area.
 *
 * Two rules shape it. Coordinates are fractions of the layer's box, so a mod
 * never measures pixels. Tones are names the host maps to theme tokens, so a mod
 * never picks a colour. The layer takes no pointer events: a press goes to the
 * app underneath, which is what keeps it a layer over the screen rather than
 * another place to click. It is chrome — a document, its rendering and its
 * print/export paths never see it.
 */

export const OVERLAY_TONES = ['accent', 'muted', 'success', 'danger'] as const
export type OverlayTone = (typeof OVERLAY_TONES)[number]

export const OVERLAY_LIMITS = {
  /** shapes in one layer */
  maxShapes: 64,
  /** characters in one label */
  maxLabelChars: 80,
} as const

interface Placed {
  /** fraction of the layer's width, 0 at the left edge */
  x: number
  /** fraction of the layer's height, 0 at the top edge */
  y: number
  tone?: OverlayTone
}

export type OverlayShape =
  /** a rectangle around a region, optionally named */
  | (Placed & {
      kind: 'frame'
      width: number
      height: number
      label?: string
      dashed?: boolean
    })
  /** a straight guide */
  | {
      kind: 'line'
      x1: number
      y1: number
      x2: number
      y2: number
      tone?: OverlayTone
      dashed?: boolean
    }
  /** a marker, optionally named */
  | (Placed & { kind: 'dot'; label?: string })
  /** free text on the layer */
  | (Placed & { kind: 'label'; value: string })

export interface OverlayError {
  field: string
  message: string
}

function fraction(value: unknown, field: string, errors: OverlayError[]): void {
  if (typeof value !== 'number' || !Number.isFinite(value) || value < 0 || value > 1) {
    errors.push({ field, message: 'must be a fraction between 0 and 1' })
  }
}

function tone(value: unknown, field: string, errors: OverlayError[]): void {
  if (value !== undefined && !(OVERLAY_TONES as readonly string[]).includes(String(value))) {
    errors.push({ field, message: 'must be one of ' + OVERLAY_TONES.join(', ') })
  }
}

function optionalText(value: unknown, field: string, errors: OverlayError[]): void {
  if (value === undefined) return
  if (typeof value !== 'string' || value.trim().length === 0) {
    errors.push({ field, message: 'needs a non-empty string' })
    return
  }
  if (value.length > OVERLAY_LIMITS.maxLabelChars) {
    errors.push({
      field,
      message: 'is longer than ' + OVERLAY_LIMITS.maxLabelChars + ' characters',
    })
  }
}

/** Validate one overlay's shapes; the caller owns the error list. */
export function validateOverlayShapes(
  value: unknown,
  field: string,
  errors: OverlayError[],
): number {
  if (!Array.isArray(value) || value.length === 0) {
    errors.push({ field, message: 'needs at least one shape' })
    return 0
  }
  if (value.length > OVERLAY_LIMITS.maxShapes) {
    errors.push({ field, message: 'has more than ' + OVERLAY_LIMITS.maxShapes + ' shapes' })
  }
  value.forEach((shape, index) => {
    const at = field + '[' + index + ']'
    if (!shape || typeof shape !== 'object' || Array.isArray(shape)) {
      errors.push({ field: at, message: 'needs a shape object' })
      return
    }
    const node = shape as Record<string, unknown>
    switch (node.kind) {
      case 'frame':
        fraction(node.x, at + '.x', errors)
        fraction(node.y, at + '.y', errors)
        fraction(node.width, at + '.width', errors)
        fraction(node.height, at + '.height', errors)
        tone(node.tone, at + '.tone', errors)
        optionalText(node.label, at + '.label', errors)
        break
      case 'line':
        fraction(node.x1, at + '.x1', errors)
        fraction(node.y1, at + '.y1', errors)
        fraction(node.x2, at + '.x2', errors)
        fraction(node.y2, at + '.y2', errors)
        tone(node.tone, at + '.tone', errors)
        break
      case 'dot':
        fraction(node.x, at + '.x', errors)
        fraction(node.y, at + '.y', errors)
        tone(node.tone, at + '.tone', errors)
        optionalText(node.label, at + '.label', errors)
        break
      case 'label':
        fraction(node.x, at + '.x', errors)
        fraction(node.y, at + '.y', errors)
        tone(node.tone, at + '.tone', errors)
        if (typeof node.value !== 'string' || node.value.trim().length === 0) {
          errors.push({ field: at + '.value', message: 'needs the text to draw' })
        } else if (node.value.length > OVERLAY_LIMITS.maxLabelChars) {
          errors.push({
            field: at + '.value',
            message: 'is longer than ' + OVERLAY_LIMITS.maxLabelChars + ' characters',
          })
        }
        break
      default:
        errors.push({
          field: at + '.kind',
          message: 'must be frame, line, dot or label',
        })
    }
  })
  return value.length
}
