import { describe, expect, it } from 'vitest'
import {
  adoptImageRewrites,
  rewriteDocumentImageSources,
} from '../src/renderer/document/image-rewrites'

describe('HTML Save As image rewrites', () => {
  it('adopts returned paths in image sources without changing lookalike text', () => {
    const source =
      '<p>images/photo.png</p><IMG alt="preview" SRC=\'images&#47;photo.png\'><img src=images/photo.png>'

    expect(
      rewriteDocumentImageSources(source, new Map([['images/photo.png', 'assets/photo-1.png']])),
    ).toBe(
      '<p>images/photo.png</p><IMG alt="preview" SRC=\'assets/photo-1.png\'><img src=assets/photo-1.png>',
    )
  })

  it('updates the saved baseline and keeps edits made during an async save', () => {
    const textAtSave = '<img src="images/photo.png">\n<p>Saved text</p>'
    const liveText = '<img src="images/photo.png">\n<p>Saved text plus an edit</p>'

    expect(
      adoptImageRewrites(textAtSave, liveText, [
        { from: 'images/photo.png', to: 'assets/photo-1.png' },
      ]),
    ).toEqual({
      savedText: '<img src="assets/photo-1.png">\n<p>Saved text</p>',
      liveText: '<img src="assets/photo-1.png">\n<p>Saved text plus an edit</p>',
    })
  })
})
