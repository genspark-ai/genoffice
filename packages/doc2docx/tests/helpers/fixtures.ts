import { readFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'

const FIXTURE_DIR = fileURLToPath(new URL('../fixtures', import.meta.url))

/** the real documents the expected .txt files were exported from (LibreOffice) */
export const FIXTURE_NAMES = [
  'text-basic',
  'text-long',
  'word97-sample',
  'formatting',
  'styles',
] as const
export type FixtureName = (typeof FIXTURE_NAMES)[number]

export function fixtureBytes(name: FixtureName): Uint8Array {
  return new Uint8Array(readFileSync(`${FIXTURE_DIR}/${name}.doc`))
}

/** the expected body text: LibreOffice export with the BOM stripped and CRLF normalized */
export function expectedText(name: FixtureName): string {
  return readFileSync(`${FIXTURE_DIR}/${name}.txt`, 'utf8')
    .replace(/^\uFEFF/, '')
    .replace(/\r\n/g, '\n')
}

/** expected paragraphs; a trailing newline leaves the final empty paragraph */
export function expectedParagraphs(name: FixtureName): string[] {
  return expectedText(name).split('\n')
}
