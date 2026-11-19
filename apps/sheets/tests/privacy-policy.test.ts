import { describe, expect, it } from 'vitest'

import { applyPrivacyPolicy } from '../src/ai/privacy-policy'

describe('applyPrivacyPolicy', () => {
  it('removes denied columns and redacts sensitive values locally', () => {
    const context = applyPrivacyPolicy([
      { name: 'secret', policy: 'deny', values: ['token'] },
      {
        name: 'contact',
        policy: 'redact',
        values: ['alice@example.com', '+1 (415) 555-0100'],
      },
    ])

    expect(context).toEqual([
      {
        name: 'contact',
        policy: 'redact',
        values: ['[EMAIL]', '[PHONE]'],
      },
    ])
  })

  it('uploads only counts for statistics-only columns', () => {
    const context = applyPrivacyPolicy([
      { name: 'salary', policy: 'statistics-only', values: [100, null, 200] },
    ])

    expect(context).toEqual([
      {
        name: 'salary',
        policy: 'statistics-only',
        statistics: { count: 3, nonEmptyCount: 2 },
      },
    ])
  })
})
