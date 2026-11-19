import { noteEmbeddedFontsChanged } from './line-metrics'
import { announceFontChange } from './system-fonts'

/**
 * The fonts the user picked in Format > Font Substitution to stand in for a
 * font this computer lacks. A pick is kept for the computer, as in Word, and
 * is put to work as local font faces under the missing font's own name: every
 * place that names the font (runs, styles, headers, text boxes, the canvas
 * that measures them) then draws and measures the stand-in.
 */
const STORE_KEY = 'genoffice.docs.fontSubstitutes'

interface SubstituteFace {
  src: string
  weight: 'normal' | 'bold'
  style: 'normal' | 'italic'
}

interface StoredSubstitute {
  /** the missing font, as documents name it */
  name: string
  substitute: string
  faces: SubstituteFace[]
}

const chosen = new Map<string, StoredSubstitute>()
const registered = new Map<string, FontFace[]>()

const keyOf = (name: string): string => name.trim().toLowerCase()
const local = (name: string): string => `local("${name.replace(/["\\]/g, '')}")`

interface LocalFontData {
  family: string
  postscriptName: string
  style: string
}

const CUTS: Record<string, Pick<SubstituteFace, 'weight' | 'style'>> = {
  regular: { weight: 'normal', style: 'normal' },
  bold: { weight: 'bold', style: 'normal' },
  italic: { weight: 'normal', style: 'italic' },
  oblique: { weight: 'normal', style: 'italic' },
  'bold italic': { weight: 'bold', style: 'italic' },
  'bold oblique': { weight: 'bold', style: 'italic' },
}

/** the installed regular, bold and italic cuts of a family; the family by name when they cannot be listed */
async function facesOf(family: string): Promise<SubstituteFace[]> {
  const whole: SubstituteFace[] = [{ src: local(family), weight: 'normal', style: 'normal' }]
  const query = (window as { queryLocalFonts?: () => Promise<readonly LocalFontData[]> })
    .queryLocalFonts
  if (!query) return whole
  try {
    const faces: SubstituteFace[] = []
    for (const font of await query.call(window)) {
      const cut = font.family === family ? CUTS[font.style.trim().toLowerCase()] : undefined
      if (!cut || faces.some((f) => f.weight === cut.weight && f.style === cut.style)) continue
      faces.push({ src: local(font.postscriptName), ...cut })
    }
    return faces.some((f) => f.weight === 'normal' && f.style === 'normal') ? faces : whole
  } catch {
    return whole
  }
}

function unregister(key: string): void {
  for (const face of registered.get(key) ?? []) document.fonts.delete(face)
  registered.delete(key)
}

/** the cuts that load, as faces under the missing font's name; null when the regular cut does not */
async function loadFaces(entry: StoredSubstitute): Promise<FontFace[] | null> {
  const loaded: FontFace[] = []
  let regular = false
  for (const cut of entry.faces) {
    const face = new FontFace(entry.name, cut.src, { weight: cut.weight, style: cut.style })
    try {
      await face.load()
    } catch {
      continue
    }
    loaded.push(face)
    if (cut.weight === 'normal' && cut.style === 'normal') regular = true
  }
  return regular ? loaded : null
}

/**
 * Puts the stand-in to work. False when its regular cut does not load, by its
 * listed cuts or by its family name: what stood in before is then left as it was.
 */
async function register(entry: StoredSubstitute): Promise<boolean> {
  if (typeof FontFace === 'undefined' || typeof document === 'undefined') return false
  let faces = await loadFaces(entry)
  if (!faces) {
    entry.faces = [{ src: local(entry.substitute), weight: 'normal', style: 'normal' }]
    faces = await loadFaces(entry)
  }
  if (!faces) return false
  const key = keyOf(entry.name)
  unregister(key)
  for (const face of faces) document.fonts.add(face)
  registered.set(key, faces)
  return true
}

function persist(): void {
  try {
    globalThis.localStorage?.setItem(STORE_KEY, JSON.stringify([...chosen.values()]))
  } catch {
    /* the pick still holds for this session */
  }
}

function changed(name: string): void {
  noteEmbeddedFontsChanged([name])
  announceFontChange([name])
}

/** the font picked to stand in for a missing one; null when the app decides */
export function fontSubstituteOf(name: string): string | null {
  return chosen.get(keyOf(name))?.substitute ?? null
}

/**
 * Makes `substitute` stand in for the missing font `name`, or with null hands
 * the choice back to the app. Resolves false when the stand-in cannot be loaded.
 */
export async function setFontSubstitute(name: string, substitute: string | null): Promise<boolean> {
  const key = keyOf(name)
  if (!substitute) {
    if (!chosen.delete(key)) return true
    unregister(key)
    persist()
    changed(name)
    return true
  }
  const entry: StoredSubstitute = { name, substitute, faces: await facesOf(substitute) }
  if (!(await register(entry))) return false
  chosen.set(key, entry)
  persist()
  changed(name)
  return true
}

function restore(): void {
  let stored: StoredSubstitute[]
  try {
    stored = JSON.parse(globalThis.localStorage?.getItem(STORE_KEY) ?? '[]') as StoredSubstitute[]
  } catch {
    return
  }
  if (!Array.isArray(stored)) return
  for (const entry of stored) {
    if (!entry?.name || !entry.substitute || !Array.isArray(entry.faces)) continue
    chosen.set(keyOf(entry.name), entry)
    void register(entry).then((ok) => {
      if (ok) changed(entry.name)
      else chosen.delete(keyOf(entry.name))
    })
  }
}

restore()
