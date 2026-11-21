import { describe, expect, it } from 'vitest'
import { commentBalloonLabel } from '../src/renderer/components/PaginationPreview'

type Identity = { author: string; initials?: string }

describe('commentBalloonLabel', () => {
  it("keeps the comment author's stored initials", () => {
    expect(commentBalloonLabel({ author: 'Alex Chen', initials: 'KICC' }, 3, 'Commented')).toBe(
      'Commented [KICC3]: ',
    )
  })

  it('derives initials from a single-word author when none were stored', () => {
    expect(commentBalloonLabel({ author: 'Alex' }, 1, 'Commented')).toBe('Commented [A1]: ')
  })

  it('derives initials from the first letters of a multi-word author', () => {
    expect(commentBalloonLabel({ author: '  alex   chen  ' }, 2, 'Commented')).toBe(
      'Commented [AC2]: ',
    )
  })

  it('leaves the number alone when there is no comment', () => {
    expect(commentBalloonLabel(undefined, 4, 'Commented')).toBe('Commented [4]: ')
  })

  it('keeps the translated prefix verbatim', () => {
    const comment: Identity = { author: 'Alex' }
    expect(commentBalloonLabel(comment, 1, '批注')).toBe('批注 [A1]: ')
  })
})
