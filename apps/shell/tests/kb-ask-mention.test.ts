import { describe, expect, it } from 'vitest'
import type { KbFileInfo } from '@genoffice/agent-core'
import {
  matchMentionFiles,
  mentionToken,
  mentionKeyAction,
  nextMentionIndex,
  removeMentionToken,
} from '../src/renderer/src/KbAskPanel'

/**
 * The @ mention in the Ask composer: the token the reader is typing, the
 * corpus files it matches, and the draft it leaves behind once a file is
 * picked. The panel uses these helpers directly, so they are tested here
 * without mounting it.
 */

const files: KbFileInfo[] = [
  {
    path: '/corpus/lease.docx',
    folder: 'contracts',
    name: 'lease.docx',
    ext: 'docx',
    sizeBytes: 1,
  },
  { path: '/corpus/q3.xlsx', folder: 'budgets', name: 'q3.xlsx', ext: 'xlsx', sizeBytes: 2 },
  { path: '/corpus/notes.docx', folder: 'notes', name: 'notes.docx', ext: 'docx', sizeBytes: 3 },
]

describe('the @ token being typed', () => {
  it('is open after a bare @', () => {
    expect(mentionToken('@')).toEqual({ start: 0, query: '' })
  })

  it('starts at the @ after whitespace', () => {
    expect(mentionToken('summarize @lea')).toEqual({ start: 10, query: 'lea' })
  })

  it('opens in the middle of a sentence', () => {
    expect(mentionToken('compare @q3 and')).toBeNull()
  })

  it('does not fire inside an email address', () => {
    expect(mentionToken('mail me at a@b')).toBeNull()
  })

  it('is closed once the token is followed by a space', () => {
    expect(mentionToken('@lease now')).toBeNull()
  })

  it('keeps the query after the @ only', () => {
    expect(mentionToken('ask @q3')).toEqual({ start: 4, query: 'q3' })
  })
})

describe('matching corpus files to the mention query', () => {
  it('lists the whole corpus for an empty query', () => {
    expect(matchMentionFiles(files, '')).toHaveLength(3)
  })

  it('matches the file name case-insensitively', () => {
    expect(matchMentionFiles(files, 'LEASE').map((file) => file.name)).toEqual(['lease.docx'])
  })

  it('matches the folder', () => {
    expect(matchMentionFiles(files, 'budget').map((file) => file.name)).toEqual(['q3.xlsx'])
  })

  it('matches the extension', () => {
    expect(matchMentionFiles(files, 'docx').map((file) => file.name)).toEqual([
      'lease.docx',
      'notes.docx',
    ])
  })

  it('is empty when nothing matches', () => {
    expect(matchMentionFiles(files, 'invoice')).toEqual([])
  })
})

describe('picking a mention', () => {
  it('removes the token and keeps the text before it', () => {
    expect(removeMentionToken('summarize @lea', 10)).toBe('summarize ')
  })

  it('leaves an empty draft when the whole draft is the token', () => {
    expect(removeMentionToken('@lea', 0)).toBe('')
  })

  it('wraps the highlight at both ends', () => {
    expect(nextMentionIndex(0, -1, 3)).toBe(2)
    expect(nextMentionIndex(2, 1, 3)).toBe(0)
    expect(nextMentionIndex(1, 1, 0)).toBe(0)
  })
})

describe('deciding a keydown against the open @ list', () => {
  const base = {
    key: 'Enter',
    shiftKey: false,
    isImeConfirm: false,
    listOpen: true,
    listLength: 3,
    navigated: false,
  }

  it('leaves an IME confirmation alone, for Enter and for Escape', () => {
    // the candidate window's keys are not list keys even when the key name matches
    expect(mentionKeyAction({ ...base, isImeConfirm: true }).action).toBe('none')
    expect(mentionKeyAction({ ...base, key: 'Escape', isImeConfirm: true }).action).toBe('none')
    expect(mentionKeyAction({ ...base, key: 'a', isImeConfirm: true }).action).toBe('none')
  })

  it('dismisses on Escape, which keeps the typed token', () => {
    expect(mentionKeyAction({ ...base, key: 'Escape' })).toEqual({ action: 'dismiss' })
    // a closed list is the composer's: its own Escape stops the run
    expect(mentionKeyAction({ ...base, key: 'Escape', listOpen: false }).action).toBe('none')
  })

  it('picks on Tab anywhere, and on Enter only after the reader moved into the list', () => {
    // a trailing @word takes over Enter only once the arrows were used —
    // "what did we send to @acme" must still send with a folder named acme
    expect(mentionKeyAction(base)).toEqual({ action: 'none' })
    expect(mentionKeyAction({ ...base, navigated: true })).toEqual({ action: 'pick' })
    expect(mentionKeyAction({ ...base, key: 'Tab' })).toEqual({ action: 'pick' })
  })

  it('Shift+Enter always passes through for a newline', () => {
    expect(mentionKeyAction({ ...base, shiftKey: true }).action).toBe('none')
    expect(mentionKeyAction({ ...base, shiftKey: true, navigated: true }).action).toBe('none')
  })

  it('moves with the arrows and stops at an empty list', () => {
    expect(mentionKeyAction({ ...base, key: 'ArrowDown' })).toEqual({ action: 'move', step: 1 })
    expect(mentionKeyAction({ ...base, key: 'ArrowUp' })).toEqual({ action: 'move', step: -1 })
    expect(mentionKeyAction({ ...base, key: 'ArrowDown', listLength: 0 }).action).toBe('none')
    expect(mentionKeyAction({ ...base, key: 'Enter', listLength: 0 }).action).toBe('none')
  })
})
