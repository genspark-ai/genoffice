import { describe, expect, it } from 'vitest'
import { launchPathsWithLastSession, readLastSessionDocuments } from '../src/main/last-session'

const all = () => true

describe('readLastSessionDocuments', () => {
  it('keeps absolute string paths once, in order', () => {
    expect(
      readLastSessionDocuments({
        lastSessionDocuments: ['/a.docx', 'relative.docx', 7, null, '/b.xlsx', '/a.docx'],
      }),
    ).toEqual(['/a.docx', '/b.xlsx'])
  })

  it('treats a missing or malformed value as no documents', () => {
    expect(readLastSessionDocuments({})).toEqual([])
    expect(readLastSessionDocuments({ lastSessionDocuments: '/a.docx' })).toEqual([])
  })

  it('caps a hand-edited list', () => {
    const many = Array.from({ length: 500 }, (_, i) => `/doc-${i}.md`)
    expect(readLastSessionDocuments({ lastSessionDocuments: many })).toHaveLength(50)
  })
})

describe('launchPathsWithLastSession', () => {
  const saved = { lastSessionDocuments: ['/a.docx', '/gone.pptx', '/b.xlsx'] }

  it('opens only the launch paths while the setting is off', () => {
    expect(launchPathsWithLastSession(saved, ['/c.pdf'], all)).toEqual(['/c.pdf'])
  })

  it('restores existing documents first, then the launch paths, without duplicates', () => {
    const settings = { ...saved, restoreLastSession: true }
    const exists = (p: string) => p !== '/gone.pptx'
    expect(launchPathsWithLastSession(settings, ['/b.xlsx', '/c.pdf'], exists)).toEqual([
      '/a.docx',
      '/b.xlsx',
      '/c.pdf',
    ])
  })
})
