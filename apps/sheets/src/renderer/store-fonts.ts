/**
 * Store-font FontFace registration: font files downloaded or imported into
 * <userData>/fonts reach the grid and dialogs as FontFaces, served per-face
 * over IPC (ttc collections arrive as standalone sfnt extracts). Canvas cell
 * and shape drawing consult document.fonts, so registered families render
 * everywhere a family name is used.
 */
const registered = new Set<string>()

/** Register every store face not yet registered; returns the new family names. */
export async function registerStoreFonts(): Promise<string[]> {
  let faces
  try {
    faces = await window.desktopApi.fontStoreFaces()
  } catch {
    return []
  }
  const added: string[] = []
  for (const face of faces) {
    const key = `${face.file}#${face.faceOffset}`
    if (registered.has(key)) continue
    try {
      const data = await window.desktopApi.fontData(face.file, face.faceOffset)
      if (!data) continue
      const ff = new FontFace(face.family, data, {
        weight: String(face.weight),
        style: face.italic ? 'italic' : 'normal',
      })
      await ff.load()
      document.fonts.add(ff)
      registered.add(key)
      added.push(face.family)
    } catch {
      /* unreadable face: skip */
    }
  }
  return added
}
