import { describe, expect, it } from 'vitest'
import { isDocEmpty } from '../src/renderer/document/blank'

describe('isDocEmpty', () => {
  it('treats an untitled buffer and a bare skeleton as empty', () => {
    expect(isDocEmpty('')).toBe(true)
    expect(
      isDocEmpty(
        '<!doctype html>\n<html lang="en"><head><meta charset="utf-8"><title>Untitled</title><style>body{margin:0}</style></head><body>\n<div class="wrap">&nbsp;</div>\n<script>console.log("hi")</script></body></html>',
      ),
    ).toBe(true)
  })

  it('ignores head, comments and style/script text', () => {
    expect(
      isDocEmpty('<head><title>Real title</title></head><body><!-- todo: hero --></body>'),
    ).toBe(true)
  })

  it('handles an omitted </head> and upper-case tag names', () => {
    expect(isDocEmpty('<head><title>Real title</title><body></body>')).toBe(true)
    expect(isDocEmpty('<head><title>Real title</title>')).toBe(true)
    expect(isDocEmpty('<HEAD><TITLE>T</TITLE></HEAD><BODY><STYLE>p{}</style></BODY>')).toBe(true)
    expect(isDocEmpty('<head><title>T</title><body><p>Real copy</p></body>')).toBe(false)
    expect(isDocEmpty('<head><title>T</title><h1>Implicit body</h1>')).toBe(false)
  })

  it('counts body text and media as content', () => {
    expect(isDocEmpty('<body><h1>Hello</h1></body>')).toBe(false)
    expect(isDocEmpty('<body><img src="a.png"></body>')).toBe(false)
    expect(isDocEmpty('<body><svg viewBox="0 0 1 1"></svg></body>')).toBe(false)
    expect(isDocEmpty('<body>plain text without tags</body>')).toBe(false)
  })
})
