import { describe, expect, it } from 'vitest'

import {
  externalLinkRevealRequestSchema,
  externalLinksRequestSchema,
} from '../src/shared/desktop-api'

const SESSION = '11111111-1111-4111-8111-111111111111'

describe('external-link IPC requests are session-bound', () => {
  it('list/update name a workbook session, never a path', () => {
    expect(externalLinksRequestSchema.parse({ sessionId: SESSION })).toEqual({ sessionId: SESSION })
    // the old path-shaped request must not be accepted any more
    expect(() => externalLinksRequestSchema.parse({ path: '/tmp/book.xlsx' })).toThrow()
    expect(() => externalLinksRequestSchema.parse({ sessionId: 'not-a-uuid' })).toThrow()
  })

  it('locate names the session plus a positive link index', () => {
    expect(externalLinkRevealRequestSchema.parse({ sessionId: SESSION, index: 2 })).toEqual({
      sessionId: SESSION,
      index: 2,
    })
    // a path is not a link index: the main process resolves the source itself
    expect(() =>
      externalLinkRevealRequestSchema.parse({ sessionId: SESSION, path: '/tmp/source.xlsx' }),
    ).toThrow()
    expect(() => externalLinkRevealRequestSchema.parse({ sessionId: SESSION, index: 0 })).toThrow()
  })
})
