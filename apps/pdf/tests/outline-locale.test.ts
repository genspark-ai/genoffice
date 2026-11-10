import { describe, expect, it } from 'vitest'
import { strings } from '../src/renderer/i18n/strings'

/**
 * The outline strings have to be in their own language's block.
 *
 * A key-parity check cannot see this. `strings` is typed
 * `keyof typeof strings.zh`, so every block already carries every key: a
 * translation that lands in the wrong one satisfies the type and every
 * locale-parity gate, and only the script gives it away. These strings were
 * once swapped between the Russian and the Italian block — a reviewer caught
 * it by reading, not by a failing test — so the script is pinned here.
 */
const CYRILLIC = /[Ѐ-ӿ]/

const OUTLINE_KEYS = [
  'outline',
  'outlineAdd',
  'outlineRename',
  'outlineDelete',
  'outlineNewTitle',
  'outlineEdit',
  'outlineGenerated',
] as const

describe('the outline strings', () => {
  it('are Russian in the Russian block', () => {
    for (const key of OUTLINE_KEYS) {
      expect(strings.ru[key], `ru.${key}`).toMatch(CYRILLIC)
    }
  })

  it('are not Russian in the Italian block', () => {
    for (const key of OUTLINE_KEYS) {
      expect(strings.it[key], `it.${key}`).not.toMatch(CYRILLIC)
    }
  })
})
