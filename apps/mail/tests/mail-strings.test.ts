import { describe, expect, it } from 'vitest'
import { strings } from '../src/renderer/i18n/strings'

/**
 * Mail locale tables: zh defines the key set; every other locale covers
 * exactly the same keys. Non-zh/en locales share the en table until Lane R
 * (mail.accounts/folders/list/reader/search/ai.*) and Lane C
 * (mail.compose.*) fill their prefixes.
 */
const locales = Object.keys(strings) as Array<keyof typeof strings>
const referenceKeys = Object.keys(strings.zh).sort()

describe('mail locale tables', () => {
  it('covers the frozen key inventory in zh and en', () => {
    for (const key of [
      'mail.tabTitle',
      'mail.closeUnsavedDetail',
      'mail.noKeyringBanner',
      'mail.accounts.sessionOnly',
      'mail.folders.unreadCount',
      'mail.list.virtualizedCount',
      'mail.search.noResults',
      'mail.reader.blockedImagesNote',
      'mail.compose.outboxFailed',
      'mail.ai.citationsOpen',
      'mail.errors.conflictNotice',
    ]) {
      expect(referenceKeys).toContain(key)
    }
    expect((strings.en as Record<string, string>)['mail.tabTitle']).toBe('Mail')
  })

  it.each(locales)('locale %s has exactly the zh key set', (locale) => {
    expect(Object.keys(strings[locale]).sort()).toEqual(referenceKeys)
  })

  it.each(locales)('locale %s has no empty values', (locale) => {
    const empty = Object.entries(strings[locale])
      .filter(([, value]) => typeof value !== 'string' || value.trim().length === 0)
      .map(([key]) => key)
    expect(empty).toEqual([])
  })
})
