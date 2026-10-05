import { isAbsolute, join, relative, resolve, sep } from 'node:path'

export function isValidExportPageNumber(value: unknown): value is number {
  return Number.isSafeInteger(value) && (value as number) > 0
}

export function hasValidExportPageNumbers(
  value: unknown,
  expectedLength: number,
): value is number[] {
  return (
    Array.isArray(value) &&
    value.length === expectedLength &&
    value.every((pageNumber) => isValidExportPageNumber(pageNumber)) &&
    // duplicate pages would write two exports over one target file
    new Set(value).size === value.length
  )
}

export function buildExportImagePaths(
  dir: string,
  baseName: unknown,
  pageNumbers: readonly unknown[],
): string[] {
  const root = resolve(dir)
  const safeBase = String(baseName || 'page').replace(/[/\\:*?"<>|]/g, '_')
  return pageNumbers.map((pageNumber) => {
    if (!isValidExportPageNumber(pageNumber)) throw new Error('pdf: invalid page numbers')
    const target = join(root, `${safeBase}-p${pageNumber}.png`)
    const fromRoot = relative(root, target)
    if (!fromRoot || fromRoot === '..' || fromRoot.startsWith(`..${sep}`) || isAbsolute(fromRoot)) {
      throw new Error('pdf: invalid export path')
    }
    return target
  })
}
