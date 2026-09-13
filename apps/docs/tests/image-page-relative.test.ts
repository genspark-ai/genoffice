/**
 * Anchored pictures whose wp:positionV is page/margin-relative: the posOffset
 * is a position on the landing page (Word), not a distance below the anchor
 * paragraph. Overlays (wrapNone/behind) render in content-area coordinates
 * and hand the page re-pin to the pagination; CSS-floated side wraps start at
 * their anchor line and let the engine shift them down.
 */
import { Editor } from '@tiptap/core'
import { parseDocx } from '@genoffice/docx-engine'
import { describe, expect, it } from 'vitest'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'

const PIC =
  '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
  '<pic:pic><pic:blipFill><a:blip r:embed="rId10"/></pic:blipFill>' +
  '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1806363" cy="738967"/></a:xfrm></pic:spPr></pic:pic></a:graphicData></a:graphic>'

const anchor = (relV: string, wrap: string, behind: string) =>
  '<w:r><w:drawing>' +
  `<wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="251658240" behindDoc="${behind}" locked="0" layoutInCell="1" allowOverlap="1">` +
  '<wp:simplePos x="0" y="0"/>' +
  '<wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>' +
  `<wp:positionV relativeFrom="${relV}"><wp:posOffset>1417320</wp:posOffset></wp:positionV>` +
  '<wp:extent cx="1806363" cy="738967"/>' +
  wrap +
  '<wp:docPr id="1" name="Image 1"/>' +
  PIC +
  '</wp:anchor></w:drawing></w:r>'

const SQUARE = '<wp:wrapSquare wrapText="bothSides"/>'
const TEXT = '<w:r><w:t>Purpose: the picture floats beside this paragraph.</w:t></w:r>'

async function open(bodyXml: string) {
  const parsed = await parseDocx(await buildDocx({ bodyXml, withImage: true }))
  return new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
}

describe('run-level anchors (paragraph with text)', () => {
  it('renders a page-relative behind-text logo in content-area coordinates and marks it for re-pin', async () => {
    const editor = await open(`<w:p>${anchor('page', '<wp:wrapNone/>', '1')}${TEXT}</w:p>`)
    const img = editor.view.dom.querySelector<HTMLElement>('img.doc-inline-img--behind')!
    const style = img.getAttribute('style')!
    // 1417320 EMU = 148.8px from the page top; the page margin comes off in CSS
    expect(style).toContain('translate(0.0px, calc(148.8px - var(--doc-margin-top,0px)))')
    expect(img.dataset.pageRelV).toBe('1')
    expect(img.dataset.pageRelFrom).toBe('page')
    editor.destroy()
  })

  it('margin-relative offsets already are content-area coordinates', async () => {
    const editor = await open(`<w:p>${anchor('margin', '<wp:wrapNone/>', '0')}${TEXT}</w:p>`)
    const img = editor.view.dom.querySelector<HTMLElement>('img.doc-inline-img')!
    expect(img.getAttribute('style')).toContain('translate(0.0px, 148.8px)')
    expect(img.dataset.pageRelV).toBe('1')
    expect(img.dataset.pageRelFrom).toBeUndefined()
    editor.destroy()
  })

  it('paragraph-relative overlays keep the plain offset and no re-pin marker', async () => {
    const editor = await open(`<w:p>${anchor('paragraph', '<wp:wrapNone/>', '0')}${TEXT}</w:p>`)
    const img = editor.view.dom.querySelector<HTMLElement>('img.doc-inline-img')!
    expect(img.getAttribute('style')).toContain('translate(0.0px, 148.8px)')
    expect(img.dataset.pageRelV).toBeUndefined()
    editor.destroy()
  })

  it('a page-relative side wrap starts at its anchor line instead of a page offset below it', async () => {
    const editor = await open(`<w:p>${anchor('page', SQUARE, '0')}${TEXT}</w:p>`)
    const img = editor.view.dom.querySelector<HTMLElement>('img.doc-inline-img--wrap-square-left')!
    expect(parseFloat(img.style.marginTop || '0')).toBe(0)
    expect(img.getAttribute('style')).not.toContain('shape-outside')
    editor.destroy()
  })
})

describe('block-level anchors (picture-only paragraph)', () => {
  it('page-relative overlay: content-area top plus the re-pin markers on the wrap', async () => {
    const editor = await open(`<w:p>${anchor('page', '<wp:wrapNone/>', '1')}</w:p>`)
    const wrap = editor.view.dom.querySelector<HTMLElement>('.img-wrap-behind .doc-img-wrap')!
    expect(wrap.style.top).toBe('calc(148.8px - var(--doc-margin-top,0px))')
    expect(wrap.dataset.pageRelV).toBe('1')
    expect(wrap.dataset.pageRelFrom).toBe('page')
    editor.destroy()
  })

  it('page-relative side wrap hands the vertical target to the pagination engine', async () => {
    const editor = await open(`<w:p>${anchor('page', SQUARE, '0')}</w:p>`)
    const block = editor.view.dom.querySelector<HTMLElement>('.doc-protected.img-wrap-square-left')!
    expect(block.dataset.tblpVy).toBe('148.8')
    expect(block.dataset.tblpVanchor).toBe('page')
    expect(block.getAttribute('style')).toMatch(/margin: var\(--tblp-dy,0px\)/)
    expect(block.getAttribute('style')).toContain('shape-outside: inset(var(--tblp-dy,0px) 0 0 0)')
    editor.destroy()
  })

  it('paragraph-relative side wrap keeps the offset as a flow margin', async () => {
    const editor = await open(`<w:p>${anchor('paragraph', SQUARE, '0')}</w:p>`)
    const block = editor.view.dom.querySelector<HTMLElement>('.doc-protected.img-wrap-square-left')!
    expect(block.dataset.tblpVy).toBeUndefined()
    expect(block.style.marginTop).toBe('148.8px')
    editor.destroy()
  })
})
