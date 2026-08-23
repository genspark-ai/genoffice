import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'

describe('PRIVACY.md analytics disclosure', () => {
  it('lists every event emitted by the shell main process', () => {
    const source = readFileSync(join(__dirname, '../src/main/index.ts'), 'utf8')
    const privacy = readFileSync(join(__dirname, '../../../PRIVACY.md'), 'utf8')
    const emitted = new Set(
      [...source.matchAll(/\banalytics\.track\(\s*['"]([a-z][a-z0-9_]*)['"]/g)].map(
        (match) => match[1],
      ),
    )
    const documented = new Set(
      [...privacy.matchAll(/^- `([a-z][a-z0-9_]*)` —/gm)].map((match) => match[1]),
    )

    expect([...documented].sort()).toEqual([...emitted].sort())
  })
})
