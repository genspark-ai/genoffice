import { afterEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { editorExtensions } from '../src/renderer/editor/extensions'
import {
  buildDocContext,
  getDocImagePlan,
  setDocImagePlan,
} from '../src/renderer/ai/protocol'
import { collectDocImages, withImageOffset } from '../src/renderer/ai/doc-images'
import { setModuleLang } from '../src/renderer/i18n/locale'
import type { AiSettings } from '../src/shared/ipc'

interface JsonNode {
  type: string
  attrs?: Record<string, unknown>
  content?: JsonNode[]
  text?: string
  marks?: Array<{ type: string; attrs?: Record<string, unknown> }>
}

const PNG_A =
  'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg=='
// a second, distinct tiny png (differs in the IDAT bytes)
const PNG_B =
  'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=='
const SVG = 'data:image/svg+xml;base64,PHN2Zy8+'

const picture = (dataUrl = PNG_A): JsonNode => ({
  type: 'docInlineImage',
  attrs: { dataUrl, widthPx: 120, heightPx: 80 },
})
const text = (t: string): JsonNode => ({ type: 'text', text: t })
const para = (content: JsonNode[]): JsonNode => ({
  type: 'docParagraph',
  attrs: { docxIndex: null },
  content,
})
const imageBlock = (dataUrl = PNG_A): JsonNode => ({
  type: 'docProtected',
  attrs: { blockType: 'image', label: 'Image', imageDataUrl: dataUrl, previewText: '' },
})

const editors = new Set<Editor>()
afterEach(() => {
  for (const e of editors) e.destroy()
  editors.clear()
  setDocImagePlan(null)
})

function createEditor(content: JsonNode[]): Editor {
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: { type: 'doc', content },
  })
  editors.add(editor)
  return editor
}

/** vision-capable vs text-only settings, going through the real registry */
const VISION_ON: AiSettings = {
  provider: 'anthropic',
  providers: { anthropic: { model: 'claude-opus-4-6', apiKey: 'k' } },
} as unknown as AiSettings
const VISION_OFF: AiSettings = {
  provider: 'glm',
  providers: { glm: { model: 'glm-4.6', apiKey: 'k' } },
} as unknown as AiSettings

setModuleLang('en')

describe('document pictures are injected as vision input', () => {
  it('text-only endpoint: nothing is collected, no plan, block list keeps the analyze_media wording', async () => {
    const editor = createEditor([imageBlock(), para([text('Body')])])
    const r = await collectDocImages(editor, VISION_OFF)
    expect(r).toBeNull()
    setDocImagePlan(null)
    const ctx = buildDocContext(editor)
    expect(ctx).toContain('use analyze_media with blockIndex 0')
    expect(ctx).not.toContain('attached as image')
  })

  it('vision endpoint: protected image block and inline pictures are collected in block order', async () => {
    const editor = createEditor([
      imageBlock(PNG_A),
      para([text('Intro')]),
      para([picture(PNG_A), picture(PNG_B)]),
    ])
    const r = (await collectDocImages(editor, VISION_ON))!
    expect(r).not.toBeNull()
    // PNG_A appears twice (block 0 and block 2) but rides the message once (dedupe)
    expect(r.images).toHaveLength(2)
    expect(r.ordinalsByBlock.get(0)).toEqual([1])
    expect(r.ordinalsByBlock.get(1)).toBeUndefined()
    expect(r.ordinalsByBlock.get(2)).toEqual([1, 2])
    expect(r.skipped).toBe(0)
  })

  it('the block list annotates attached pictures with their image ordinals', async () => {
    const editor = createEditor([imageBlock(PNG_A), para([picture(PNG_B)])])
    const r = (await collectDocImages(editor, VISION_ON))!
    setDocImagePlan(withImageOffset(r, 0))
    const ctx = buildDocContext(editor)
    expect(ctx).toContain('(picture — attached as image #1)')
    expect(ctx).toContain('(picture — attached as image #2)')
  })

  it('a mixed text+picture paragraph names the attached ordinal after its text preview', async () => {
    const editor = createEditor([para([text('See figure: '), picture(PNG_B)])])
    const r = (await collectDocImages(editor, VISION_ON))!
    setDocImagePlan(withImageOffset(r, 0))
    const ctx = buildDocContext(editor)
    expect(ctx).toContain('See figure:')
    expect(ctx).toContain('(+1 attached image #1)')
  })

  it('attachment images shift the doc image ordinals so #K is images[K-1] of the merged array', async () => {
    const editor = createEditor([imageBlock(PNG_A)])
    const r = (await collectDocImages(editor, VISION_ON))!
    const plan = withImageOffset(r, 2)
    expect(plan.ordinalsByBlock.get(0)).toEqual([3])
    setDocImagePlan(plan)
    expect(buildDocContext(editor)).toContain('attached as image #3')
  })

  it('non-vision mimes are skipped and counted', async () => {
    const editor = createEditor([imageBlock(SVG), imageBlock(PNG_A)])
    const r = (await collectDocImages(editor, VISION_ON))!
    expect(r.images).toHaveLength(1)
    expect(r.skipped).toBe(1)
    expect(r.ordinalsByBlock.get(0)).toBeUndefined()
    expect(r.ordinalsByBlock.get(1)).toEqual([1])
  })

  it('the budget caps collection at 8 pictures and reports the overflow', async () => {
    const blocks: JsonNode[] = []
    for (let i = 0; i < 10; i++) {
      // distinct payloads so dedupe does not collapse them: embed the index in a comment chunk
      blocks.push(imageBlock(`${PNG_A.slice(0, -4)}${btoa(String(i)).replace(/=+$/, '')}==`))
    }
    const editor = createEditor(blocks)
    const r = (await collectDocImages(editor, VISION_ON))!
    expect(r.images).toHaveLength(8)
    expect(r.skipped).toBe(2)
  })

  it('lazy media URLs resolve through the preload bridge', async () => {
    const lazy = `genoffice-docx-media://${'a'.repeat(64)}/word/media/image1.png`
    const editor = createEditor([imageBlock(lazy)])
    const calls: string[] = []
    const prev = window.desktop
    ;(window as unknown as Record<string, unknown>).desktop = {
      ...prev,
      readImageBytes: async (src: string) => {
        calls.push(src)
        return { base64: 'aGVsbG8=', mime: 'image/png' }
      },
    }
    try {
      const r = (await collectDocImages(editor, VISION_ON))!
      expect(calls).toEqual([lazy])
      expect(r.images[0]).toEqual({ base64: 'aGVsbG8=', mime: 'image/png' })
    } finally {
      ;(window as unknown as Record<string, unknown>).desktop = prev
    }
  })

  it('plan holder round-trips and clears', () => {
    expect(getDocImagePlan()).toBeNull()
    setDocImagePlan({ ordinalsByBlock: new Map([[0, [1]]]), skipped: 0 })
    expect(getDocImagePlan()?.ordinalsByBlock.get(0)).toEqual([1])
    setDocImagePlan(null)
    expect(getDocImagePlan()).toBeNull()
  })
})
