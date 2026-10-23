import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'

import { describe, expect, it } from 'vitest'

import { shellWindowBackground } from '../src/main/window-background'

const tokens = readFileSync(resolve(__dirname, '../../../packages/ui/src/tokens.css'), 'utf8')

function surfaceIn(selector: string): string | undefined {
  return new RegExp(`${selector}\\s*\\{[^}]*?--surface:\\s*(#[0-9a-fA-F]{3,8})\\s*;`).exec(
    tokens,
  )?.[1]
}

describe('shellWindowBackground', () => {
  it('picks the dark or light surface', () => {
    expect(shellWindowBackground(true)).toBe('#1e1e1e')
    expect(shellWindowBackground(false)).toBe('#ffffff')
  })

  it('matches --surface in tokens.css for both themes', () => {
    expect(surfaceIn(':root')?.toLowerCase()).toBe(shellWindowBackground(false))
    expect(surfaceIn("\\[data-theme='dark'\\]")?.toLowerCase()).toBe(shellWindowBackground(true))
  })
})
