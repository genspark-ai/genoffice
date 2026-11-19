/**
 * Type shim: @genoffice/file-parse ships as TS source, and its dependency
 * @genoffice/docx-engine fails under sheets' stricter compiler options
 * (exactOptionalPropertyTypes/noUncheckedIndexedAccess). tsconfig paths point
 * type resolution here; runtime bundling (electron-vite) still uses the real
 * source. Keep the export signatures in sync with packages/file-parse/src/parse.ts.
 */
export type ParsedFileKind = 'text' | 'image' | 'unsupported'

export interface ParsedFile {
  ok: boolean
  text?: string
  kind: ParsedFileKind
  mime?: string
  error?: string
}

/** parse an attachment into plain text (or flag it as image / unsupported) */
export function parseFileToText(filePath: string): Promise<ParsedFile>
