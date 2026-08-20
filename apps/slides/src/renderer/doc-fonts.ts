/**
 * Office-private font registration. Layout (main process) may resolve run fonts to files
 * Chromium cannot see (PowerPoint's bundled DFonts, Office cloud fonts); glyph positions are
 * baked with those metrics, so the canvas must draw with the same faces. Main exposes the
 * referenced faces + extracted sfnt bytes; here they become document FontFaces. Canvas layers
 * redraw via the FontFaceSet 'loadingdone' listener in SlideCanvas.
 */

const requested = new Set<string>()

declare global {
  interface Window {
    /** Set after each private-font sync completes; the fidelity harness waits on it before screenshots. */
    __genofficeDocFontsSynced?: boolean
  }
}

/** Fetch and register any private faces not yet requested. Safe to call often (idempotent). */
export async function syncPrivateFonts(): Promise<void> {
  window.__genofficeDocFontsSynced = false
  let faces: Array<{ id: string; family: string; bold: boolean; italic: boolean }>
  try {
    faces = await window.slidesApi.privateFontFaces()
  } catch {
    window.__genofficeDocFontsSynced = true
    return
  }
  // Test harnesses stub the bridge with null results
  if (!Array.isArray(faces)) {
    window.__genofficeDocFontsSynced = true
    return
  }
  const added = await Promise.all(
    faces
      .filter((f) => !requested.has(f.id))
      .map(async (f) => {
        requested.add(f.id)
        try {
          const data = await window.slidesApi.privateFontData(f.id)
          if (!data) return false
          const face = new FontFace(f.family, data, {
            weight: f.bold ? '700' : '400',
            style: f.italic ? 'italic' : 'normal',
          })
          document.fonts.add(face)
          await face.load()
          return true
        } catch {
          /* unparseable face: keep the CSS fallback chain */
          return false
        }
      }),
  )
  // ArrayBuffer-backed FontFaces parse synchronously in the constructor, so the set never
  // enters the loading state and never emits a real 'loadingdone' — canvases already
  // rastered with a fallback face would keep the stale pixels. Fire it by hand.
  if (added.some(Boolean)) document.fonts.dispatchEvent(new Event('loadingdone'))
  window.__genofficeDocFontsSynced = true
}
