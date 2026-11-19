import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { buildParseMap, type ParseMap } from '../src/renderer/document/parse-map'
import { compileOps, type HtmlOp } from '../src/renderer/document/ops'
import { applyPatches } from '../src/renderer/document/patch'
import {
  AGENT_TOOLS,
  createHtmlSkill,
  createHtmlSkillCore,
  IMAGE_GEN_OFF_NOTE,
  type HtmlDocAccess,
} from '../src/renderer/ai/tools'

/** in-memory stand-in for the App: same compile → apply → version bookkeeping */
function fakeAccess(initial: string) {
  let text = initial
  let version = 1
  let lastManual = 0
  let map: ParseMap | null = null
  const getMap = () => {
    if (!map || map.version !== version) map = buildParseMap(text, version, map)
    return map
  }
  const access: HtmlDocAccess = {
    getText: () => text,
    getVersion: () => version,
    getMap,
    getLastManualVersion: () => lastManual,
    getFilePath: () => '/tmp/page.html',
    getSelectedSid: () => null,
    applyOps: (ops: HtmlOp[]) => {
      const compiled = compileOps(text, getMap(), ops)
      if (compiled.errors.length) return { ok: false, errors: compiled.errors }
      const ranges: Array<[number, number]> = []
      let delta = 0
      for (const p of [...compiled.patches].sort((a, b) => a.from - b.from)) {
        ranges.push([p.from + delta, p.from + delta + p.text.length])
        delta += p.text.length - (p.to - p.from)
      }
      text = applyPatches(text, compiled.patches)
      version++
      return { ok: true, ranges }
    },
    replaceAll: (html) => {
      text = html
      version++
    },
  }
  return {
    access,
    userTypes(next: string) {
      text = next
      version++
      lastManual = version
    },
    get text() {
      return text
    },
  }
}

const DOC = `<!doctype html>
<html>
<head><title>Report</title></head>
<body>
<h1 id="t">Quarterly</h1>
<p>Closing note.</p>
</body>
</html>`

const call = (name: string, input: Record<string, unknown>) => ({ id: 'c1', name, input })

/** png magic bytes as base64 — what the fetch handler returns for a real png */
const PNG_BASE64 = btoa(String.fromCharCode(0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a))

const bridge = {
  imageSearch: vi.fn(),
  aiGenerateImage: vi.fn(),
  fetchImage: vi.fn(),
  saveImage: vi.fn(),
}

beforeEach(() => {
  window.htmlApi = bridge as unknown as typeof window.htmlApi
})

afterEach(() => {
  bridge.imageSearch.mockReset()
  bridge.aiGenerateImage.mockReset()
  bridge.fetchImage.mockReset()
  bridge.saveImage.mockReset()
})

describe('image tool schemas', () => {
  it('declare search_images and generate_image with the sibling-required fields', () => {
    const byName = new Map(AGENT_TOOLS.map((tl) => [tl.name, tl]))
    const search = byName.get('search_images')
    const generate = byName.get('generate_image')
    expect(search?.inputSchema).toMatchObject({ required: ['query'] })
    expect(generate?.inputSchema).toMatchObject({ required: ['prompt'] })
    expect(String(search?.description)).toContain('apply_ops')
    for (const tl of AGENT_TOOLS) {
      expect(tl.name).toBeTruthy()
      expect(tl.description).toBeTruthy()
      expect(tl.inputSchema.type).toBe('object')
    }
  })
})

describe('search_images', () => {
  it('formats the gallery as a numbered list of urls', async () => {
    bridge.imageSearch.mockResolvedValue({
      images: [
        { title: 'Poster', imageUrl: 'https://a.example/p.jpg', width: 300, height: 200 },
        { imageUrl: 'https://b.example/q.png' },
      ],
      method: 'serper',
    })
    const core = createHtmlSkillCore(fakeAccess(DOC).access)
    const r = await core.executeTool(call('search_images', { query: 'poster', maxResults: 2 }))
    expect(r.isError).toBeUndefined()
    expect(r.mutated).toBe(false)
    expect(r.output).toBe(
      '1. Poster [300x200]\n   https://a.example/p.jpg\n2. (untitled) [?x?]\n   https://b.example/q.png',
    )
    expect(bridge.imageSearch).toHaveBeenCalledWith('poster', 2)
  })

  it('keeps an empty result and a backend failure apart', async () => {
    const core = createHtmlSkillCore(fakeAccess(DOC).access)
    bridge.imageSearch.mockResolvedValue({ images: [], method: 'serper' })
    const empty = await core.executeTool(call('search_images', { query: 'nothing' }))
    expect(empty.output).toBe('(no images)')
    expect(empty.isError).toBeUndefined()

    bridge.imageSearch.mockResolvedValue({ images: [], method: 'error', error: 'quota' })
    const failed = await core.executeTool(call('search_images', { query: 'nothing' }))
    expect(failed.isError).toBe(true)
    expect(failed.output).toContain('you may retry')
    expect(failed.output).toContain('quota')
  })

  it('refuses an empty query', async () => {
    const core = createHtmlSkillCore(fakeAccess(DOC).access)
    const r = await core.executeTool(call('search_images', { query: '  ' }))
    expect(r.isError).toBe(true)
    expect(r.output).toBe('query must not be empty')
    expect(bridge.imageSearch).not.toHaveBeenCalled()
  })
})

describe('generate_image', () => {
  it('generates, stores the bytes in assets/, and appends the <img> to <body>', async () => {
    const doc = fakeAccess(DOC)
    const core = createHtmlSkillCore(doc.access)
    core.buildContext() // freshness baseline, as the loop does before every run
    bridge.aiGenerateImage.mockResolvedValue({ url: 'https://cdn.example/gen.png' })
    bridge.fetchImage.mockResolvedValue({ base64: PNG_BASE64, mime: 'image/png' })
    bridge.saveImage.mockResolvedValue('assets/gen.png')
    const r = await core.executeTool(
      call('generate_image', { prompt: 'a chart mascot', aspectRatio: '16:9' }),
    )
    expect(r.isError).toBeUndefined()
    expect(r.mutated).toBe(true)
    expect(r.output).toContain('assets/gen.png')
    expect(bridge.aiGenerateImage).toHaveBeenCalledWith({
      prompt: 'a chart mascot',
      aspectRatio: '16:9',
    })
    expect(bridge.saveImage).toHaveBeenCalledWith({ base64: PNG_BASE64, ext: 'png' })
    expect(doc.text).toContain('<img src="assets/gen.png" style="max-width: 100%;">')
    expect(doc.text.indexOf('<img')).toBeGreaterThan(doc.text.indexOf('<h1'))
  })

  it('inlines a data URL when the document is unsaved and hot-links when the download fails', async () => {
    const unsaved = fakeAccess(DOC)
    const core = createHtmlSkillCore(unsaved.access)
    core.buildContext()
    bridge.aiGenerateImage.mockResolvedValue({ url: 'https://cdn.example/gen.png' })
    bridge.fetchImage.mockResolvedValue({ base64: PNG_BASE64, mime: 'image/png' })
    bridge.saveImage.mockResolvedValue(null)
    const r = await core.executeTool(call('generate_image', { prompt: 'mascot' }))
    expect(r.mutated).toBe(true)
    expect(r.output).toContain('data URL')
    expect(unsaved.text).toContain(`<img src="data:image/png;base64,${PNG_BASE64}"`)

    const hot = fakeAccess(DOC)
    const hotCore = createHtmlSkillCore(hot.access)
    hotCore.buildContext()
    bridge.aiGenerateImage.mockResolvedValue({ url: 'https://cdn.example/gen.png' })
    bridge.fetchImage.mockResolvedValue(null)
    const hotR = await hotCore.executeTool(call('generate_image', { prompt: 'mascot', alt: 'm' }))
    expect(hotR.mutated).toBe(true)
    expect(hot.text).toContain('<img src="https://cdn.example/gen.png" alt="m"')
  })

  it('surfaces the generation error and refuses an empty prompt', async () => {
    const doc = fakeAccess(DOC)
    const core = createHtmlSkillCore(doc.access)
    core.buildContext()
    bridge.aiGenerateImage.mockResolvedValue({ error: 'signed out' })
    const failed = await core.executeTool(call('generate_image', { prompt: 'x' }))
    expect(failed.isError).toBe(true)
    expect(failed.output).toBe('signed out')
    expect(doc.text).toBe(DOC)

    const empty = await core.executeTool(call('generate_image', { prompt: ' ' }))
    expect(empty.isError).toBe(true)
    expect(empty.output).toBe('prompt must not be empty')
  })

  it('refuses to write when the user edited during the generation', async () => {
    const doc = fakeAccess(DOC)
    const core = createHtmlSkillCore(doc.access)
    core.buildContext()
    let releaseGenerate: ((v: { url: string }) => void) | undefined
    bridge.aiGenerateImage.mockReturnValue(
      new Promise<{ url: string }>((resolve) => {
        releaseGenerate = resolve
      }),
    )
    const pending = core.executeTool(call('generate_image', { prompt: 'mascot' }))
    doc.userTypes(`${DOC}\n<!-- user typed -->`)
    releaseGenerate?.({ url: 'https://cdn.example/gen.png' })
    bridge.fetchImage.mockResolvedValue({ base64: PNG_BASE64, mime: 'image/png' })
    bridge.saveImage.mockResolvedValue('assets/gen.png')
    const r = await pending
    expect(r.isError).toBe(true)
    expect(r.output).toContain('changed since you last looked')
    expect(doc.text).not.toContain('<img')
  })
})

describe('generate_image gating', () => {
  it('hides generate_image and notes it while no image provider is reachable', () => {
    const access = fakeAccess(DOC).access
    const off = createHtmlSkill(access, 'PROMPT', () => false)
    expect(off.tools.map((tl) => tl.name)).not.toContain('generate_image')
    expect(off.tools.map((tl) => tl.name)).toContain('search_images')
    expect(off.systemPrompt).toContain('generate_image is currently unavailable')
    expect(off.systemPrompt).toBe(`PROMPT${IMAGE_GEN_OFF_NOTE}`)

    const on = createHtmlSkill(access, 'PROMPT', () => true)
    expect(on.tools.map((tl) => tl.name)).toContain('generate_image')
    expect(on.systemPrompt).toBe('PROMPT')
  })
})
