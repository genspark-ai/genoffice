/**
 * The writing view's font picker plumbing (#2654 next cut): names are CSS
 * string values so quotes and rule characters never ride through, the picked
 * face rides in front of the view's own stack (a missing font just falls
 * through), and the choice persists across sessions.
 */
import { beforeEach, describe, expect, it } from 'vitest'

import {
  STORY_FACE_CHOICES,
  faceInstalled,
  loadStoryFace,
  saveStoryFace,
  sanitizeFaceName,
  storyFaceRule,
} from '../src/renderer/story-font'

describe('face names are CSS string values', () => {
  it('keeps plain names and drops rule characters', () => {
    expect(sanitizeFaceName('iA Writer Mono S')).toBe('iA Writer Mono S')
    expect(sanitizeFaceName("Neon'; --story-size: 99px; '")).toBe('Neon --story-size: 99px')
    expect(sanitizeFaceName('a\\b"c;d{e}')).toBe('abcde')
    expect(sanitizeFaceName('  IBM Plex Mono  ')).toBe('IBM Plex Mono')
  })

  it('refuses empty or overlong names', () => {
    expect(sanitizeFaceName("''")).toBe('')
    expect(sanitizeFaceName('   ')).toBe('')
    expect(sanitizeFaceName('x'.repeat(65))).toBe('')
    expect(sanitizeFaceName('x'.repeat(64))).toHaveLength(64)
  })
})

describe('the picked face rides in front of the view stack', () => {
  it('injects the rule with the fallback tail behind the face', () => {
    const rule = storyFaceRule('Iosevka Etoile')
    expect(rule).toBe(
      ".view-story{--story-face:'Iosevka Etoile', 'Monaspace Neon', 'iA Writer Mono S', " +
        "'iA Writer Quattro S', 'IBM Plex Mono', 'Iosevka Etoile', 'Recursive', system-ui, " +
        "-apple-system, 'Segoe UI', 'PingFang SC', 'Microsoft YaHei', sans-serif;}",
    )
  })

  it('does not repeat Monaspace when it is the pick', () => {
    const rule = storyFaceRule('Monaspace Neon')
    expect(rule).toContain(".view-story{--story-face:'Monaspace Neon', 'iA Writer Mono S'")
    expect((rule.match(/Monaspace Neon/g) ?? []).length).toBe(1)
  })

  it('a custom name is sanitized into the rule, and empty stays silent', () => {
    expect(storyFaceRule("Face'; }")).toContain("'Face'")
    expect(storyFaceRule('')).toBe('')
    expect(storyFaceRule("';")).toBe('')
  })
})

describe('the choice persists', () => {
  beforeEach(() => {
    localStorage.clear()
  })

  it('round-trips and defaults to the view stack', () => {
    expect(loadStoryFace()).toBe('')
    saveStoryFace('IBM Plex Mono')
    expect(loadStoryFace()).toBe('IBM Plex Mono')
    saveStoryFace('')
    expect(loadStoryFace()).toBe('')
  })
})

describe('the curated list', () => {
  it('carries the bundled faces first, never a duplicate of the tail', () => {
    expect(STORY_FACE_CHOICES[0]).toEqual({ family: 'Monaspace Neon', bundled: true })
    expect(STORY_FACE_CHOICES[1]).toEqual({ family: 'IBM Plex Mono', bundled: true })
    const families = STORY_FACE_CHOICES.map((c) => c.family)
    expect(new Set(families).size).toBe(families.length)
  })

  it('faceInstalled answers false where the check is unavailable', () => {
    // jsdom has no working font check: unknown must not throw
    expect(faceInstalled('Monaspace Neon')).toBe(false)
  })
})
