/**
 * The documents the Docs app opens. `.odt`/`.ott` arrive through the import
 * filter, which converts a copy — the same route the shell takes — so the
 * standalone argv route, the external-open guard and the open dialog's picker
 * filter all read this one list instead of drifting apart.
 */

export const OPENABLE_DOC_EXTENSIONS = ['docx', 'odt', 'ott'] as const

const OPENABLE_DOC_PATTERN = new RegExp(`\\.(${OPENABLE_DOC_EXTENSIONS.join('|')})$`, 'i')

export function isOpenableDocPath(path: string): boolean {
  return OPENABLE_DOC_PATTERN.test(path.trim())
}

export function findOpenableDocPath(argv: readonly string[]): string | null {
  return (
    argv.find((arg) => {
      const value = arg.trim()
      return value.length > 0 && !value.startsWith('-') && isOpenableDocPath(value)
    }) ?? null
  )
}
