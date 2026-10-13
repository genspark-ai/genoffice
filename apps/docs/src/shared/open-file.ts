export function findOpenableDocPath(argv: readonly string[]): string | null {
  return (
    argv.find((arg) => {
      const value = arg.trim()
      return value.length > 0 && !value.startsWith('-') && /\.(docx|odt)$/i.test(value)
    }) ?? null
  )
}
