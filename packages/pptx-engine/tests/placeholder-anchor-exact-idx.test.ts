import { describe, expect, it } from 'vitest'
import {
  parsePlaceholderMap,
  resolvePlaceholderAnchor,
  resolvePlaceholderInsets,
} from '../src/placeholder'

const NS =
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"'
const sp = (ph: string, bodyPr: string) =>
  `<p:sp><p:nvSpPr><p:cNvPr id="2" name="x"/><p:cNvSpPr/><p:nvPr>${ph}</p:nvPr></p:nvSpPr>` +
  `<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm></p:spPr>` +
  `<p:txBody>${bodyPr}<a:lstStyle/><a:p/></p:txBody></p:sp>`
const layer = (...shapes: string[]) =>
  `<p:sldLayout ${NS}><p:cSld><p:spTree>${shapes.join('')}</p:spTree></p:cSld></p:sldLayout>`

// Two-content layout: the caption body (idx 1) is bottom-anchored, the content
// body (idx 2) says nothing and inherits the master's top anchor
const layout = parsePlaceholderMap(
  layer(
    sp('<p:ph type="body" idx="1"/>', '<a:bodyPr anchor="b" lIns="0"/>'),
    sp('<p:ph sz="half" idx="2"/>', '<a:bodyPr><a:normAutofit/></a:bodyPr>'),
  ),
)
const master = parsePlaceholderMap(
  layer(sp('<p:ph type="body" idx="1"/>', '<a:bodyPr anchor="t" tIns="45720"/>')),
)

describe('placeholder bodyPr inheritance follows the exact (type, idx) layout entry', () => {
  it('a content placeholder does not borrow a sibling body anchor from its layout', () => {
    expect(resolvePlaceholderAnchor(layout, master, undefined, '2')).toBe('top')
    expect(resolvePlaceholderAnchor(layout, master, 'body', '1')).toBe('bottom')
  })

  it('insets skip the sibling too and merge from the master', () => {
    expect(resolvePlaceholderInsets(layout, master, undefined, '2')).toEqual({ t: 45720 })
    expect(resolvePlaceholderInsets(layout, master, 'body', '1')).toEqual({ l: 0, t: 45720 })
  })

  it('without an exact layout entry the same-type fallback still applies', () => {
    expect(resolvePlaceholderAnchor(layout, master, 'body', '7')).toBe('bottom')
  })
})
