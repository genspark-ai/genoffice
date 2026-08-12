import type { DrawingInput } from '../shared/ipc'
import type { SignatureData } from './SignatureDialog'
import type { FormWidget } from './form-catalog'

/** Fit an existing visual-signature payload into an AcroForm /Sig widget rectangle. */
export function signatureDrawingForField(
  sig: SignatureData,
  target: FormWidget,
  color: [number, number, number],
): DrawingInput {
  const [x1, y1, x2, y2] = target.rect
  const fieldW = Math.abs(x2 - x1)
  const fieldH = Math.abs(y2 - y1)
  const inset = Math.min(fieldW, fieldH) * 0.08
  const availW = Math.max(1, fieldW - inset * 2)
  const availH = Math.max(1, fieldH - inset * 2)
  const k = Math.min(availW / Math.max(sig.width, 1), availH / Math.max(sig.height, 1))
  const width = sig.width * k
  const height = sig.height * k
  const left = Math.min(x1, x2) + (fieldW - width) / 2
  const bottom = Math.min(y1, y2) + (fieldH - height) / 2

  if (sig.kind === 'image') {
    return {
      kind: 'image',
      pageIndex: target.pageIndex,
      image: sig.image,
      rect: [left, bottom, left + width, bottom + height],
      formFieldName: target.fieldName,
    }
  }

  const top = bottom + height
  return {
    kind: 'ink',
    pageIndex: target.pageIndex,
    color,
    width: 1.6,
    formFieldName: target.fieldName,
    paths: sig.paths.map((path) => {
      const out: number[] = []
      for (let index = 0; index < path.length; index += 2) {
        out.push(left + path[index]! * k, top - path[index + 1]! * k)
      }
      return out
    }),
  }
}
