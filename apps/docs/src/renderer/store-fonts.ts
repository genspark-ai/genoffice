/**
 * Store-font FontFace registration: font files downloaded or imported into
 * <userData>/fonts reach the DOM editor as FontFaces, served per-face over IPC
 * (ttc collections arrive as standalone sfnt extracts). Registered families are
 * immediately selectable and count as available for the open-time missing-font
 * check, which measures via canvas.
 */
const registered = new Set<string>()

/** Register every store face not yet registered; returns the new family names. */
export async function registerStoreFonts(): Promise<string[]> {
  let faces
  try {
    faces = await window.desktop.fontStoreFaces()
  } catch {
    return []
  }
  const added: string[] = []
  for (const face of faces) {
    const key = `${face.file}#${face.faceOffset}`
    if (registered.has(key)) continue
    try {
      const data = await window.desktop.fontData(face.file, face.faceOffset)
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
