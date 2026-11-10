import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'

/**
 * The dialog locale table (src/main/markdown-main.ts) is an inline object handed
 * to createI18n, so the values are read out of the source rather than imported,
 * the same way the shell's main-locale-tables.test.ts does.
 *
 * The pt block shipped with the Spanish "Archivos de texto" instead of the
 * Portuguese "Arquivos de texto" — one letter apart, so a key-set check never
 * saw it, and a Portuguese user saw a Spanish save-dialog filter.
 */
const source = readFileSync(join(__dirname, '../src/main/markdown-main.ts'), 'utf8')

/** One locale's slice of the table, from its `xx: {` header to the next one. */
function localeBlock(locale: string): string {
  const start = source.search(new RegExp(`\\n {2}${locale}: \\{`))
  if (start < 0) throw new Error(`no ${locale} block in the dialog locale table`)
  const rest = source.slice(start + 1)
  const end = rest.search(/\n {2}\S+: \{/)
  return end < 0 ? rest : rest.slice(0, end)
}

function filterMarkdownIn(locale: string): string {
  const value = localeBlock(locale).match(/filterMarkdown: '([^']*)'/)?.[1]
  if (value === undefined) throw new Error(`no filterMarkdown in the ${locale} block`)
  return value
}

describe('the text-file save-dialog filter', () => {
  it('is Portuguese in the pt block', () => {
    expect(filterMarkdownIn('pt')).toBe('Arquivos de texto (Markdown, TXT, JSON)')
  })

  it('is still Spanish in the es block', () => {
    // the es string is correct there; fixing pt must not "fix" es as well
    expect(filterMarkdownIn('es')).toBe('Archivos de texto (Markdown, TXT, JSON)')
  })
})
