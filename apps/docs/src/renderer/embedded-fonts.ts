import type { EmbeddedFont, EmbeddedFontLineMetrics } from '@genoffice/docx-engine'
import { noteEmbeddedFontsChanged, setEmbeddedLineMetrics } from './line-metrics'

let active: FontFace[] = []
let generation = 0

function sfntBuffer(data: Uint8Array): ArrayBuffer {
  const whole = data.byteOffset === 0 && data.byteLength === data.buffer.byteLength
  return (whole ? data.buffer : data.slice().buffer) as ArrayBuffer
}

/**
 * Register the document's embedded faces (word/fonts) under their fontTable
 * names so they win over the substitution chains, like Word's private font
 * install for the open document. Faces of the previous document are revoked.
 * Resolves false when a newer adoption started meanwhile; the caller's document
 * is then stale and must not be applied.
 */
export async function adoptEmbeddedFonts(
  fonts: readonly EmbeddedFont[] | undefined,
): Promise<boolean> {
  if (typeof FontFace === 'undefined' || typeof document === 'undefined') return true
  const my = ++generation
  const loaded = await Promise.all(
    (fonts ?? []).map(async (f) => {
      try {
        const face = new FontFace(f.family, sfntBuffer(f.data), {
          weight: f.bold ? '700' : '400',
          style: f.italic ? 'italic' : 'normal',
        })
        await face.load()
        return face
      } catch {
        return null
      }
    }),
  )
  if (my !== generation) return false
  const revoked = active
  for (const face of revoked) document.fonts.delete(face)
  active = []
  const boxes: Array<{ styled: boolean; family: string } & EmbeddedFontLineMetrics> = []
  loaded.forEach((face, i) => {
    if (!face) return
    document.fonts.add(face)
    active.push(face)
    const f = (fonts ?? [])[i]
    if (f?.lineMetrics)
      boxes.push({ styled: f.bold || f.italic, family: f.family, ...f.lineMetrics })
  })
  // the regular cut's box stands for the family
  setEmbeddedLineMetrics(boxes.sort((a, b) => Number(a.styled) - Number(b.styled)))
  if (revoked.length === 0 && active.length === 0) return true
  noteEmbeddedFontsChanged([...revoked, ...active].map((f) => f.family))
  // buffer-backed faces parse synchronously: the set never enters 'loading', so
  // the measurement caches keyed on 'loadingdone' must be told by hand
  document.fonts.dispatchEvent(new Event('loadingdone'))
  return true
}
