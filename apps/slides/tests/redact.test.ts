import { describe, expect, it } from 'vitest'
import {
  MAX_LABEL_LENGTH,
  checkPlaceholders,
  collectPlaceholders,
  isWholePlaceholder,
  isWithheldMedia,
  mediaLabelOf,
  mediaModelView,
  placeholderInstruction,
  placeholderSource,
  readPlaceholderLabel,
  sanitizeLabel,
  type MediaNodeLike,
} from '../src/renderer/ai/redact'

/**
 * The pure-function layer for withheld content, ported from the docs app.
 *
 * Everything here is a function of its arguments: no editor, no schema, no
 * engine. The media cases are slides' addition to that layer — a deck can
 * withhold a picture, a video or an audio clip, and none of them have words to
 * redact, so what must not reach the model is the *reference* to the content.
 */

/** A signed link: a replayable credential, not a path inside the deck. */
const SIGNED_URL = 'https://cdn.example.com/clip.mp4?sig=abc123&expires=1893456000'

describe('the write guard', () => {
  const before = 'call {{client phone}} to confirm the order'

  it('accepts a reply that keeps every marker', () => {
    expect(
      checkPlaceholders(before, 'please call {{client phone}} before the end of today.'),
    ).toEqual([])
  })

  it('rejects a split across a line break', () => {
    expect(checkPlaceholders(before, 'call {{client\nphone}} to confirm').length).toBeGreaterThan(0)
  })

  it('rejects an interleaving, which an atom could never have prevented', () => {
    expect(checkPlaceholders('{{a}}{{b}}', '{{{{a}}b}}').length).toBeGreaterThan(0)
  })

  it('rejects a drop, a rename, an invention and a duplicate', () => {
    expect(checkPlaceholders('{{a}} {{b}}', '{{a}}').length).toBeGreaterThan(0)
    expect(checkPlaceholders('{{a}}', '{{a2}}').length).toBeGreaterThan(0)
    expect(checkPlaceholders('plain', 'plain {{new}}').length).toBeGreaterThan(0)
    expect(checkPlaceholders('{{a}}', '{{a}} {{a}}').length).toBeGreaterThan(0)
  })

  it('reports a half-typed marker as a split, not a mere absence', () => {
    expect(checkPlaceholders('{{a}}', '{{a}').some((i) => i.reason === 'split')).toBe(true)
    expect(checkPlaceholders('{{a}}', '{a}').some((i) => i.reason === 'split')).toBe(true)
  })

  it('does not report leftovers for a well-formed marker', () => {
    expect(checkPlaceholders('{{a}}', 'Call {{a}} now').some((i) => i.reason === 'split')).toBe(
      false,
    )
  })

  it('never throws on anything a model might produce', () => {
    for (const after of ['', '{{', '}}', '{{{{', '}}}}', '😀{{a}}😀', '{{a}}\n\n{{b}}']) {
      expect(() => checkPlaceholders('{{a}}', after)).not.toThrow()
    }
  })
})

describe('labels', () => {
  it('strips characters that would make the marker ambiguous', () => {
    expect(sanitizeLabel('a{b}<c>"d"')).toBe('abcd')
  })

  it('falls back to a neutral word', () => {
    expect(placeholderSource('')).toBe('{{private}}')
  })

  it('caps the length', () => {
    expect(sanitizeLabel('x'.repeat(100)).length).toBeLessThanOrEqual(MAX_LABEL_LENGTH)
  })

  it('recognises a whole marker but not partial text', () => {
    expect(isWholePlaceholder('{{a}}')).toBe(true)
    expect(isWholePlaceholder('call {{a}}')).toBe(false)
    expect(readPlaceholderLabel('{{client phone}}')).toBe('client phone')
    expect(readPlaceholderLabel('call {{a}}')).toBeNull()
  })
})

describe('the model instruction', () => {
  it('lists each marker this deck has, once', () => {
    const text2 = placeholderInstruction(['phone', 'address', 'phone'])
    expect(text2).toContain('- {{phone}}')
    expect(text2).toContain('- {{address}}')
    expect(text2.split('\n').filter((l) => l === '- {{phone}}')).toHaveLength(1)
  })

  it('forbids each way a marker gets damaged', () => {
    const text2 = placeholderInstruction(['a'])
    expect(text2).toMatch(/never split/i)
    expect(text2).toMatch(/never merge/i)
    expect(text2).toMatch(/never rename/i)
    expect(text2).toMatch(/never drop/i)
    expect(text2).toMatch(/never add a placeholder/i)
  })

  it('says a placeholder may stand in for withheld media', () => {
    // slides extends the docs wording: a deck withholds video and audio too
    const text2 = placeholderInstruction(['a'])
    expect(text2).toMatch(/picture/i)
    expect(text2).toMatch(/video/i)
    expect(text2).toMatch(/audio/i)
  })
})

describe('collectPlaceholders', () => {
  it('keeps duplicates', () => {
    expect(collectPlaceholders('{{a}} {{b}} {{a}}')).toEqual(['{{a}}', '{{b}}', '{{a}}'])
  })
})

// ---------------------------------------------------------------------------
// Media: the part slides adds on top of the ported module
// ---------------------------------------------------------------------------

const withheldVideo: MediaNodeLike = {
  type: 'picture',
  name: 'Quarterly review',
  redact: 'customer meeting recording',
  media: { kind: 'video', target: SIGNED_URL, external: true },
  transform: { x: 100, y: 200, w: 1920000, h: 1080000 },
  durationMs: 1_845_000,
  mediaRef: 'ppt/media/media1.mp4',
  dataUrl: 'data:video/mp4;base64,AAAA',
  descr: 'internal review, do not share',
}

const withheldPicture: MediaNodeLike = {
  type: 'picture',
  name: 'Passport scan',
  redact: 'ID photo',
  mediaRef: 'ppt/media/image3.png',
  dataUrl: 'data:image/png;base64,AAAA',
  transform: { x: 0, y: 0, w: 914400, h: 914400 },
}

const plainVideo: MediaNodeLike = {
  type: 'picture',
  name: 'Product demo',
  media: { kind: 'video', target: 'ppt/media/media2.mp4' },
  transform: { x: 0, y: 0, w: 1920000, h: 1080000 },
}

const withheldAudio: MediaNodeLike = {
  type: 'picture',
  name: 'Voicemail',
  redact: 'customer voice message',
  media: { kind: 'audio', target: SIGNED_URL, external: true },
}

describe('isWithheldMedia / mediaLabelOf', () => {
  it('recognises a withheld node and reads its label', () => {
    expect(isWithheldMedia(withheldVideo)).toBe(true)
    expect(mediaLabelOf(withheldVideo)).toBe('customer meeting recording')
  })

  it('reports no mark on a plain media node', () => {
    expect(isWithheldMedia(plainVideo)).toBe(false)
    expect(mediaLabelOf(plainVideo)).toBeNull()
  })

  it('treats an empty label as no mark', () => {
    // the engine's readRedactLabel already returns undefined for an empty
    // label; a blank string must not become a {{private}} marker either
    expect(isWithheldMedia({ type: 'picture', redact: '' })).toBe(false)
    expect(mediaLabelOf({ type: 'picture', redact: '' })).toBeNull()
  })
})

describe('mediaModelView — a withheld node', () => {
  it('never emits the target, so a signed link is not handed over', () => {
    // the whole reason media is treated separately from text: for a linked
    // clip, target is an external URL the model could fetch and replay
    const flat = JSON.stringify(mediaModelView(withheldVideo))
    expect(flat).not.toContain('cdn.example.com')
    expect(flat).not.toContain('sig=abc123')
    expect(flat).not.toContain('http')
    expect(mediaModelView(withheldVideo).target).toBeUndefined()
  })

  it('never emits duration or dimensions', () => {
    // a confidential recording's length and size are information about it
    const view = mediaModelView(withheldVideo)
    const flat = JSON.stringify(view)
    expect(flat).not.toContain('1845000')
    expect(view.frame).toBeUndefined()
    expect(view).not.toHaveProperty('durationMs')
  })

  it('never emits a media path or the decoded bytes', () => {
    const flat = JSON.stringify(mediaModelView(withheldVideo))
    expect(flat).not.toContain('ppt/media/media1.mp4')
    expect(flat).not.toContain('base64')
  })

  it('never emits the engine descr metadata', () => {
    // <p:cNvPr descr> is passed through verbatim on save and is free-form
    expect(JSON.stringify(mediaModelView(withheldVideo))).not.toContain('do not share')
  })

  it('gives the model the label and a marker, and nothing else', () => {
    expect(mediaModelView(withheldVideo)).toEqual({
      kind: 'video',
      label: 'customer meeting recording',
      placeholder: '{{customer meeting recording}}',
    })
  })

  it('withholds a picture the same way, since it has no text either', () => {
    const view = mediaModelView(withheldPicture)
    expect(view.kind).toBe('picture')
    expect(view.placeholder).toBe('{{ID photo}}')
    const flat = JSON.stringify(view)
    expect(flat).not.toContain('image3.png')
    expect(flat).not.toContain('base64')
  })

  it('withholds audio, whose target is a link just as often', () => {
    const view = mediaModelView(withheldAudio)
    expect(view.kind).toBe('audio')
    expect(view.placeholder).toBe('{{customer voice message}}')
    expect(view.target).toBeUndefined()
  })

  it('sanitises the label in the marker, so braces cannot be smuggled in', () => {
    expect(mediaModelView({ type: 'picture', redact: 'a{b}' }).placeholder).toBe('{{ab}}')
  })

  it('falls back to a neutral marker for an unusable label', () => {
    expect(mediaModelView({ type: 'picture', redact: '<<>>' }).placeholder).toBe('{{private}}')
  })

  it('keeps the media kind, which is what the model writes prose around', () => {
    expect(mediaModelView(withheldAudio).kind).toBe('audio')
    expect(mediaModelView(withheldPicture).kind).toBe('picture')
  })
})

describe('mediaModelView — a node that is not withheld', () => {
  it('passes an ordinary video through with its target and frame', () => {
    // the fix is scoped: a node the reader did not withhold is not the business
    // of this layer, and blanking it would break every non-redacted deck
    expect(mediaModelView(plainVideo)).toEqual({
      kind: 'video',
      name: 'Product demo',
      frame: { x: 0, y: 0, w: 1920000, h: 1080000 },
      target: 'ppt/media/media2.mp4',
    })
  })

  it('marks an external link so the model can tell a reference from a path', () => {
    const view = mediaModelView({
      type: 'picture',
      media: { kind: 'video', target: SIGNED_URL, external: true },
    })
    expect(view.external).toBe(true)
    expect(view.target).toBe(SIGNED_URL)
  })

  it('omits keys it has no value for rather than emitting undefined', () => {
    const view = mediaModelView({ type: 'picture' })
    expect(view).toEqual({ kind: 'picture' })
    expect(Object.keys(view)).toEqual(['kind'])
  })
})

describe('mediaModelView — a whitelist, not a blacklist', () => {
  it('drops any field it was not told about', () => {
    // the property that keeps holding as PictureElement grows: a new field is
    // withheld by construction rather than by someone remembering a deny-list
    const view = mediaModelView({
      ...withheldVideo,
      someFutureField: 'internal revenue figures',
      videoCodec: 'h264',
    })
    const flat = JSON.stringify(view)
    expect(flat).not.toContain('internal revenue figures')
    expect(flat).not.toContain('h264')
  })

  it('does not alias the node, so a later edit cannot leak into the view', () => {
    const node: MediaNodeLike = { ...withheldVideo }
    const view = mediaModelView(node)
    node.media!.target = 'https://elsewhere.example.com/other.mp4'
    expect(view).not.toHaveProperty('target')
    expect(JSON.stringify(view)).not.toContain('elsewhere.example.com')
  })

  it('copies the frame instead of holding the same object', () => {
    const node: MediaNodeLike = { type: 'picture', transform: { x: 1, y: 2, w: 3, h: 4 } }
    const view = mediaModelView(node)
    node.transform!.w = 9999
    expect(view.frame!.w).toBe(3)
  })
})
