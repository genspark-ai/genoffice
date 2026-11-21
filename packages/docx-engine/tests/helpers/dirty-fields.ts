import JSZip from 'jszip'

const DIRTY_FIELD = /<w:(?:fldChar|fldSimple)\b[^>]*\bw:dirty="(?:true|1|on)"[^>]*>/g

/** the parts of a saved package that hold a field marked `w:dirty`, as `path: count` */
export async function dirtyFieldParts(bytes: Uint8Array | ArrayBuffer): Promise<string[]> {
  const zip = await JSZip.loadAsync(new Uint8Array(bytes as ArrayBuffer))
  const out: string[] = []
  for (const [path, file] of Object.entries(zip.files)) {
    if (file.dir || !/^word\/[^/]+\.xml$/.test(path)) continue
    const count = (await file.async('string')).match(DIRTY_FIELD)?.length ?? 0
    if (count > 0) out.push(`${path}: ${count}`)
  }
  return out
}

interface DirtyFieldSwitch {
  __allowDirtyFields?: () => void
}

/**
 * For a test whose saved file is meant to carry `w:dirty`: a field that came
 * with the mark and was not touched, or one whose result is not computed here.
 * Any other save that writes the mark fails (see setup-strict-xml).
 */
export function allowDirtyFields(): void {
  ;(globalThis as DirtyFieldSwitch).__allowDirtyFields?.()
}
