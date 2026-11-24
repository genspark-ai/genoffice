import { afterEach, describe, expect, it, vi } from 'vitest'

const insertImageFromDataUrl = vi.fn(async () => true)
vi.mock('../src/renderer/components/ribbon-tabs', () => ({ insertImageFromDataUrl }))
const clipboardPicture = vi.fn(async () => ({
  dataUrl: 'data:image/png;base64,AAAA',
  widthPx: 300,
  heightPx: 80,
}))
vi.mock('../src/renderer/editor/paste-picture', async (importOriginal) => ({
  ...(await importOriginal<typeof import('../src/renderer/editor/paste-picture')>()),
  clipboardPicture,
}))

const { pastePayload } = await import('../src/renderer/editor/paste-actions')
const { pictureContent } = await import('../src/renderer/editor/paste-picture')

afterEach(() => vi.clearAllMocks())

describe('Paste as Picture', () => {
  it('draws copied HTML as one picture at its laid-out size', async () => {
    const editor = {} as never
    const html = '<table><tr><td>North</td><td>12</td></tr></table>'
    expect(pastePayload(editor, { html, text: 'North\t12', image: null }, 'picture')).toBe(true)
    await vi.waitFor(() => expect(insertImageFromDataUrl).toHaveBeenCalled())
    expect(clipboardPicture).toHaveBeenCalledWith(html, 'North\t12')
    expect(insertImageFromDataUrl).toHaveBeenCalledWith(
      editor,
      'data:image/png;base64,AAAA',
      'Picture (pasted)',
      { width: 300, height: 80 },
    )
  })

  it('has nothing to draw from an empty clipboard', () => {
    expect(pastePayload({} as never, { html: '', text: '  ', image: null }, 'picture')).toBe(false)
    expect(clipboardPicture).not.toHaveBeenCalled()
  })

  it("keeps the content's own styles and drops what is active or remote", () => {
    const frame = pictureContent(
      '<html><head><style>.xl65{font-weight:bold}</style><script>alert(1)</script></head>' +
        '<body><table><tr><td class="xl65">Total</td></tr></table>' +
        '<img src="https://example.com/a.png"><img src="data:image/png;base64,AAAA">' +
        '<iframe src="https://example.com"></iframe></body></html>',
      '',
    )
    expect(frame.querySelector('style')?.textContent).toContain('.xl65')
    expect(frame.querySelector('td.xl65')?.textContent).toBe('Total')
    expect(frame.querySelector('script, iframe')).toBeNull()
    expect([...frame.querySelectorAll('img')].map((img) => img.getAttribute('src'))).toEqual([
      'data:image/png;base64,AAAA',
    ])
  })

  it('keeps the line breaks of plain text', () => {
    const frame = pictureContent('', 'one\ntwo')
    const block = frame.firstElementChild as HTMLElement
    expect(block.style.whiteSpace).toBe('pre-wrap')
    expect(block.textContent).toBe('one\ntwo')
  })
})
