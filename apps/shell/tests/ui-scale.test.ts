import { mkdtempSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import { describe, expect, it, vi } from 'vitest'

// The module reaches for `webContents` to stamp the value into live renderers.
// Nothing here tests that (it needs a running Electron), but the import has to
// resolve — same shape as tests/tab-manager.test.ts.
vi.mock('electron', () => ({ webContents: { getAllWebContents: () => [] } }))

const { UI_SCALE_MAX, UI_SCALE_MIN, UI_SCALE_STEPS, clampUiScale, readUiScale, stepUiScale } =
  await import('../src/main/ui-scale')

/**
 * The interface scale (#1913): the user's "make the UI bigger" preference, and
 * the arithmetic behind the View menu.
 *
 * The setting lives in a plain JSON file a user can hand-edit, so reading it
 * has to survive anything that ends up in there — that is most of these tests.
 */

function settingsFile(contents: unknown): string {
  const dir = mkdtempSync(join(tmpdir(), 'ui-scale-'))
  const path = join(dir, 'app-settings.json')
  writeFileSync(path, typeof contents === 'string' ? contents : JSON.stringify(contents))
  return path
}

describe('clampUiScale', () => {
  it('keeps a value inside the preset range', () => {
    expect(clampUiScale(1.25)).toBe(1.25)
    expect(clampUiScale(0.4)).toBe(UI_SCALE_MIN)
    expect(clampUiScale(8)).toBe(UI_SCALE_MAX)
  })

  it('offers the whole range #1913 asked for', () => {
    expect(UI_SCALE_MAX).toBe(2)
    expect(UI_SCALE_STEPS).toEqual([1, 1.1, 1.25, 1.5, 1.75, 2])
  })
})

describe('stepUiScale', () => {
  it('walks up and down through the presets', () => {
    expect(stepUiScale(1, 1)).toBe(1.1)
    expect(stepUiScale(1.1, 1)).toBe(1.25)
    expect(stepUiScale(1.25, -1)).toBe(1.1)
  })

  it('stops at both ends instead of running off', () => {
    // Stepping past the last preset has to stay there, not wrap or clamp to a
    // value the menu never offers.
    expect(stepUiScale(UI_SCALE_MAX, 1)).toBe(UI_SCALE_MAX)
    expect(stepUiScale(UI_SCALE_MIN, -1)).toBe(UI_SCALE_MIN)
  })

  it('starts from the next preset when the stored value is not one', () => {
    // `readUiScale` clamps the range but does not snap to a preset, so a user
    // who typed `1.05` into the JSON really does get 1.05 here. Stepping has to
    // find the next preset *above* it, not fall off the end: 1.05 + a step is
    // 1.25, and 1.05 - a step is 1.
    expect(stepUiScale(1.05, 1)).toBe(1.25)
    expect(stepUiScale(1.05, -1)).toBe(1)
    // The top half, where an exact-match implementation would coincidentally
    // agree — pinned so both halves are stated, not just the one that bit.
    expect(stepUiScale(1.3, 1)).toBe(1.75)
    expect(stepUiScale(1.3, -1)).toBe(1.25)
  })
})

describe('readUiScale', () => {
  it('reads a stored value', () => {
    expect(readUiScale(settingsFile({ uiScale: 1.25 }))).toBe(1.25)
  })

  it('treats a workbook that never set it as no scaling', () => {
    expect(readUiScale(settingsFile({ onboardingSeen: true }))).toBe(1)
  })

  it('treats a missing settings file as no scaling', () => {
    expect(readUiScale(join(tmpdir(), 'genoffice-does-not-exist', 'app-settings.json'))).toBe(1)
  })

  const rejected: readonly (readonly [string, unknown])[] = [
    ['a string', { uiScale: '150' }],
    ['null', { uiScale: null }],
    ['NaN written as a bare token', '{ "uiScale": NaN }'],
    ['a corrupt file', '{ not json'],
    ['an array', [1, 2, 3]],
  ]
  for (const [what, contents] of rejected) {
    it(`falls back to 1 for ${what} rather than putting it in the layout`, () => {
      // Anything that reached `zoom` unvetted would be `zoom: NaN` or
      // `zoom: "150"`, and the chrome would collapse or vanish.
      expect(readUiScale(settingsFile(contents))).toBe(1)
    })
  }

  it('clamps a stored value outside the range instead of honouring it', () => {
    expect(readUiScale(settingsFile({ uiScale: 99 }))).toBe(UI_SCALE_MAX)
    expect(readUiScale(settingsFile({ uiScale: -3 }))).toBe(UI_SCALE_MIN)
  })
})
