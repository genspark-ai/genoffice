import { describe, expect, it } from 'vitest'
import { detectSlashQuery, filterSlashSkills } from '../src/ai-slash'

const SKILLS = [
  { name: 'tables', description: 'Header rows and totals' },
  { name: 'tone', description: 'Write plainly' },
]

describe('detectSlashQuery', () => {
  it('finds a token typed at the very start of the draft', () => {
    expect(detectSlashQuery('/sum', 4)).toEqual({ start: 0, end: 4, query: 'sum' })
    expect(detectSlashQuery('/', 1)).toEqual({ start: 0, end: 1, query: '' })
  })

  it('finds a token after whitespace, with the caret right after the slash', () => {
    expect(detectSlashQuery('run /sum', 8)).toEqual({ start: 4, end: 8, query: 'sum' })
    expect(detectSlashQuery('run /sum', 5)).toEqual({ start: 4, end: 5, query: '' })
    expect(detectSlashQuery('hi\n/x', 5)).toEqual({ start: 3, end: 5, query: 'x' })
    expect(detectSlashQuery('a\t/x', 4)).toEqual({ start: 2, end: 4, query: 'x' })
  })

  it('ignores a slash in the middle of a word', () => {
    expect(detectSlashQuery('plain/text', 10)).toBeNull()
    expect(detectSlashQuery('a//b', 4)).toBeNull()
  })

  it('returns null while the caret sits before the slash', () => {
    expect(detectSlashQuery('/sum', 0)).toBeNull()
    expect(detectSlashQuery('run /sum', 4)).toBeNull()
    expect(detectSlashQuery('plain/text', 5)).toBeNull()
  })

  it('returns null once whitespace separates the slash from the caret', () => {
    // the space ends the token; what follows is the message the skill runs on
    expect(detectSlashQuery('/sum the rows', 13)).toBeNull()
    expect(detectSlashQuery('hi /sum now', 11)).toBeNull()
  })

  it('lets the slash nearest the caret win', () => {
    expect(detectSlashQuery('/a /b', 5)).toEqual({ start: 3, end: 5, query: 'b' })
    expect(detectSlashQuery('hi /a /bc', 9)).toEqual({ start: 6, end: 9, query: 'bc' })
    // the nearest slash is mid-word there, so it triggers nothing
    expect(detectSlashQuery('/a/b', 4)).toBeNull()
  })

  it('clamps a caret that runs past the text', () => {
    expect(detectSlashQuery('/ab', 99)).toEqual({ start: 0, end: 3, query: 'ab' })
    expect(detectSlashQuery('hi /ab', -1)).toBeNull()
  })

  it('hands back the range a pick must strike: [start, end) is the whole `/query`', () => {
    const text = 'please /tone it up'
    const token = detectSlashQuery(text, 12)
    expect(token).toEqual({ start: 7, end: 12, query: 'tone' })
    expect(text.slice(0, token!.start) + text.slice(token!.end)).toBe('please  it up')
  })
})

describe('filterSlashSkills', () => {
  it('returns every skill for an empty query', () => {
    expect(filterSlashSkills(SKILLS, '')).toBe(SKILLS)
    expect(filterSlashSkills(SKILLS, '  ')).toBe(SKILLS)
  })

  it('matches case-insensitively over name and description', () => {
    expect(filterSlashSkills(SKILLS, 'TAB')).toEqual([SKILLS[0]])
    expect(filterSlashSkills(SKILLS, 'Plainly')).toEqual([SKILLS[1]])
  })

  it('returns nothing when no skill matches', () => {
    expect(filterSlashSkills(SKILLS, 'zebra')).toEqual([])
  })
})
