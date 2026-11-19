import { describe, expect, it } from 'vitest'
import { resolveRelationshipTargetPath } from '../src/parse-package'

describe('resolveRelationshipTargetPath', () => {
  it('normalizes encoded, relative and rooted targets', () => {
    const rel = (target: string) => resolveRelationshipTargetPath('word/document.xml', target)
    expect(rel('media/image%201.png')).toBe('word/media/image 1.png')
    expect(rel('..\\customXml\\item1.xml')).toBe('customXml/item1.xml')
    expect(rel('/word/styles.xml')).toBe('word/styles.xml')
    expect(rel('\\word\\styles.xml')).toBe('word/styles.xml')
    expect(rel('../../word/styles.xml')).toBeNull()
    expect(rel('http://example.com/a.png')).toBeNull()
  })
})
