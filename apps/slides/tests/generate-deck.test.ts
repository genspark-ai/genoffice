/**
 * generate_deck self-driving pipeline verification:
 * the AI calls generate_deck once with the whole plan; the tool's internal code loops page by page,
 * calling the LLM to write HTML and landing it. Page count is guaranteed by the for loop -> cures
 * "planned N pages but only 1 remains" for good.
 */
import { describe, it, expect } from 'vitest'
import { createSlidesSkill, type DeckAccess } from '../src/renderer/ai/slides-skill'
import type { RenderSlide } from '@genoffice/pptx-render'
import type { AgentToolCall } from '../src/shared/ipc'

function makeAccess(opts?: {
  failPages?: number[]
  failAttempts?: Record<number, number>
  cloudEnabled?: boolean
  localImageFails?: Record<number, string[]>
  landFailOnce?: number[]
  searchImagesFail?: boolean
  searchImagesMulti?: number
  planImageQueries?: string[]
  styleTemplates?: Record<string, { topic: string; styleSkill: string; createdAt: string }>
}) {
  const failPages = new Set(opts?.failPages ?? [])
  const failAttempts = { ...(opts?.failAttempts ?? {}) } // pageIndex -> how many more times to fail
  const landFailOnce = new Set(opts?.landFailOnce ?? []) // these pages fail their first "landing" once (simulated conversion failure)
  let pages = 0
  const genPageCalls: number[] = [] // records each generatePageCloud call's pageIndex
  const localPageCalls: number[] = [] // records each generatePageLocal call's pageIndex
  const stylesSeen: string[] = [] // records the style each page received (verifies styleSkill reached single pages)
  const landOrder: string[] = [] // records the landing order
  const imageSearchCalls: string[] = [] // records the queries searchImages was called with
  const imagesSeen: string[][] = [] // records the images each generatePageCloud call received
  const sidecarSaves: Array<{ topic: string; styleSkill: string; createdAt: string }> = []
  const savedTemplates: Record<
    string,
    { name: string; topic: string; styleSkill: string; createdAt: string }
  > = {
    ...(opts?.styleTemplates
      ? Object.fromEntries(
          Object.entries(opts.styleTemplates).map(([k, v]) => [k, { name: k, ...v }]),
        )
      : {}),
  }
  const access: DeckAccess = {
    getSlides: () => [] as RenderSlide[],
    getCurrent: () => 0,
    getSelectedIds: () => [],
    applySlide: () => {},
    applyDeck: () => {},
    fitWidthPx: 1280,
    retryBackoffMs: 0,
    generateFromHtml: async (html, mode = 'replace', _deckName?: string, insertAt?: number) => {
      const failNo = [...landFailOnce].find((n) => html[0]?.includes(`PAGE${n}:`))
      if (failNo !== undefined) {
        landFailOnce.delete(failNo)
        return { ok: false, error: 'mock land fail' }
      }
      if (mode === 'insert_at') {
        pages += 1
        landOrder.push(`insert@${insertAt}:` + html[0])
        return { ok: true, pages, insertedIndex: insertAt }
      }
      if (mode === 'append') {
        const from = pages
        pages += html.length
        landOrder.push('append:' + html[0])
        return { ok: true, pages, appendedFrom: from }
      }
      pages = html.length
      landOrder.push('replace:' + html[0])
      return { ok: true, pages }
    },
    isCloudPageGenEnabled: async () => opts?.cloudEnabled !== false,
    generatePageCloud: async (args) => {
      genPageCalls.push(args.pageIndex)
      stylesSeen.push(args.style)
      imagesSeen.push([...args.images])
      if (failPages.has(args.pageIndex)) return { ok: false, error: 'mock fail' }
      if (failAttempts[args.pageIndex] && failAttempts[args.pageIndex] > 0) {
        failAttempts[args.pageIndex] -= 1
        return { ok: false, error: 'transient' }
      }
      // Return an identifiable marker (with page order); the mock landing treats it like the real cloudpptx: marker
      return {
        ok: true,
        marker: `<!doctype html><html><body>PAGE${args.pageIndex}:${args.title}</body></html>`,
      }
    },
    generatePageLocal: async (args) => {
      localPageCalls.push(args.pageIndex)
      stylesSeen.push(args.style)
      imagesSeen.push([...args.images])
      if (failPages.has(args.pageIndex)) return { ok: false, error: 'mock fail' }
      if (failAttempts[args.pageIndex] && failAttempts[args.pageIndex] > 0) {
        failAttempts[args.pageIndex] -= 1
        return { ok: false, error: 'transient' }
      }
      // Same marker contract as the cloud path, distinguishable in landOrder assertions
      const fails = opts?.localImageFails?.[args.pageIndex]
      return {
        ok: true,
        marker: `localpptx:PAGE${args.pageIndex}:${args.title}`,
        ...(fails?.length ? { imageFailures: [...fails] } : {}),
      }
    },
    generateStyleSkill: async (a) => {
      return { ok: true, styleSkill: 'STYLE_SKILL_FOR:' + a.topic }
    },
    planDeckOutline: async (a) => {
      // In-tool planning: emit pages per count, titles carry the page order, core_hook given in the first batch (style no longer produced here)
      const pagesOut = Array.from({ length: a.count }, (_, i) => ({
        title: `Planned page ${a.startPage + i}`,
        brief: `b${a.startPage + i}`,
        layout: 'data',
        image_queries: opts?.planImageQueries ? [...opts.planImageQueries] : [],
      }))
      return { ok: true, outline: { core_hook: 'CH:' + a.topic, pages: pagesOut } }
    },
    searchImages: async (query: string, maxResults: number) => {
      imageSearchCalls.push(query)
      if (opts?.searchImagesFail) return []
      if (opts?.searchImagesMulti) {
        const n = Math.min(opts.searchImagesMulti, maxResults)
        return Array.from(
          { length: n },
          (_, i) => `https://mock.img/${encodeURIComponent(query)}-${i + 1}.jpg`,
        )
      }
      return [`https://mock.img/${encodeURIComponent(query)}.jpg`]
    },
    saveSidecar: async (data) => {
      sidecarSaves.push(data)
    },
    saveStyleTemplate: async (name, data) => {
      savedTemplates[name] = { name, ...data }
      return { ok: true }
    },
    listStyleTemplates: async () =>
      Object.values(savedTemplates).map(({ name, topic, createdAt }) => ({
        name,
        topic,
        createdAt,
      })),
    loadStyleTemplate: async (name) => {
      const t = savedTemplates[name]
      if (!t) return { ok: false, error: `Template "${name}" not found` }
      return { ok: true, styleSkill: t.styleSkill, topic: t.topic }
    },
  }
  return {
    access,
    genPageCalls,
    localPageCalls,
    stylesSeen,
    landOrder,
    imageSearchCalls,
    imagesSeen,
    sidecarSaves,
    savedTemplates,
    getPages: () => pages,
  }
}

const deckCall = (n: number, insertMode?: 'replace' | 'append'): AgentToolCall => ({
  id: 'c-deck',
  name: 'generate_deck',
  input: {
    core_hook: 'Doubled in a decade',
    style: 'Deep navy + gold',
    ...(insertMode ? { insert_mode: insertMode } : {}),
    pages: Array.from({ length: n }, (_, i) => ({
      title: `Page ${i + 1} Title`,
      brief: `brief${i + 1}`,
      layout: 'data',
      image_queries: i === 0 ? ['https://example.com/a.jpg'] : [],
    })),
  },
})

describe('generate_deck self-driven page-by-page generation', () => {
  it('plans 5 pages → tool loop calls the LLM 5 times to write pages → lands 5 pages → progress complete', async () => {
    const { access, genPageCalls, getPages } = makeAccess()
    const skill = createSlidesSkill(access)

    const res = (await skill.executeTool(deckCall(5))) as { output: string; summary: string }

    // The tool internally calls generatePageCloud once per each of the 5 pages (the AI never touches per-page work)
    expect(genPageCalls.sort((a, b) => a - b)).toEqual([1, 2, 3, 4, 5])
    // All 5 pages landed
    expect(getPages()).toBe(5)
    expect(res.summary).toContain('5/5')
    expect(res.output).toContain('Self-driven generation produced 5/5 pages')

    // Each round's context shows everything generated; no more nagging
    const ctx = skill.buildContext?.() ?? ''
    expect(ctx).toContain('all generated')
    expect(ctx).not.toContain('still missing')
  })

  it('landing order strictly follows page order (first page replace, rest append, no mutual overwrites)', async () => {
    const { access, landOrder } = makeAccess()
    const skill = createSlidesSkill(access)
    await skill.executeTool(deckCall(3))
    // First page replace, next two append, and content is in PAGE1/2/3 order
    expect(landOrder[0]).toBe('replace:<!doctype html><html><body>PAGE1:Page 1 Title</body></html>')
    expect(landOrder[1]).toContain('append:')
    expect(landOrder[1]).toContain('PAGE2')
    expect(landOrder[2]).toContain('PAGE3')
  })

  it('LLM fails on one page → other pages land + tool result names the failed page to prompt a fix', async () => {
    const { access, getPages } = makeAccess({ failPages: [3] })
    const skill = createSlidesSkill(access)
    const res = (await skill.executeTool(deckCall(5))) as { output: string }

    // Page 3 failed; the other 4 pages landed
    expect(getPages()).toBe(4)
    expect(res.output).toContain('page 3 (mock fail) still failed after retry')
    expect(res.output).toContain('append')
    // Progress reflects 4/5 and still reports the shortfall
    const ctx = skill.buildContext?.() ?? ''
    expect(ctx).toContain('4 generated')
  })

  it('neither pages nor topic → errors out without executing', async () => {
    const { access } = makeAccess()
    const skill = createSlidesSkill(access)
    const bad = (await skill.executeTool({ id: 'x', name: 'generate_deck', input: {} })) as {
      isError?: boolean
    }
    expect(bad.isError).toBe(true)
  })
})

describe('generate_deck local page generation (cloud/gsk unavailable)', () => {
  it('cloud disabled → the local spec builder generates every page through the same marker landing', async () => {
    const { access, genPageCalls, localPageCalls, landOrder, getPages } = makeAccess({
      cloudEnabled: false,
    })
    const skill = createSlidesSkill(access)
    const res = (await skill.executeTool(deckCall(3))) as { output: string }
    expect(genPageCalls).toEqual([]) // cloud never touched
    expect(localPageCalls.sort((a, b) => a - b)).toEqual([1, 2, 3])
    expect(getPages()).toBe(3)
    expect(landOrder[0]).toBe('replace:localpptx:PAGE1:Page 1 Title')
    expect(res.output).toContain('3/3')
  })

  it('page failing both inline attempts gets another generation in the retry round and lands at its original position', async () => {
    const { access, landOrder, getPages } = makeAccess({
      cloudEnabled: false,
      failAttempts: { 2: 2 },
    })
    const skill = createSlidesSkill(access)
    const res = (await skill.executeTool(deckCall(3))) as { output: string }
    expect(getPages()).toBe(3)
    expect(res.output).toContain('3/3')
    expect(landOrder[2]).toContain('insert@1:')
    expect(landOrder[2]).toContain('PAGE2')
  })

  it("local image download failures surface in the tool output with the page's number", async () => {
    const { access, getPages } = makeAccess({
      cloudEnabled: false,
      localImageFails: { 2: ['https://img.example/broken.jpg'] },
    })
    const skill = createSlidesSkill(access)
    const res = (await skill.executeTool(deckCall(3))) as { output: string }
    expect(getPages()).toBe(3)
    expect(res.output).toContain('Missing images')
    expect(res.output).toContain('page 2')
    expect(res.output).toContain('https://img.example/broken.jpg')
  })

  it('neither cloud nor local pipeline available → fails fast', async () => {
    const { access } = makeAccess({ cloudEnabled: false })
    delete (access as { generatePageLocal?: unknown }).generatePageLocal
    const skill = createSlidesSkill(access)
    const r = (await skill.executeTool(deckCall(2))) as { isError?: boolean }
    expect(r.isError).toBe(true)
  })
})

describe('generate_deck batched planning (topic mode)', () => {
  const topicCall = (topic: string, approx: number): AgentToolCall => ({
    id: 'c-topic',
    name: 'generate_deck',
    input: { topic, approx_pages: approx, context: 'real source material' },
  })

  it('topic+approx_pages only → plans the outline internally → generates every page one by one', async () => {
    const { access, genPageCalls, getPages } = makeAccess()
    const skill = createSlidesSkill(access)
    const res = (await skill.executeTool(topicCall('Jinan Travel Guide', 5))) as { output: string }
    // Internal planning produced 5 pages -> generated page by page
    expect(genPageCalls.sort((a, b) => a - b)).toEqual([1, 2, 3, 4, 5])
    expect(getPages()).toBe(5)
    expect(res.output).toContain('5/5')
  })

  it('over 12 pages → recursive batched planning (no single huge JSON) → all pages generated', async () => {
    const { access, getPages } = makeAccess()
    const skill = createSlidesSkill(access)
    // 20 pages -> PLAN_BATCH=12 -> should plan in 2 batches (12+8)
    const res = (await skill.executeTool(topicCall('Annual Report', 20))) as { output: string }
    expect(getPages()).toBe(20)
    expect(res.output).toContain('20/20')
    const ctx = skill.buildContext?.() ?? ''
    expect(ctx).toContain('all generated')
  })

  it('style generated independently → the same styleSkill is passed to every page (consistent across pages)', async () => {
    const { access, stylesSeen } = makeAccess()
    const skill = createSlidesSkill(access)
    await skill.executeTool(topicCall('Jinan Travel Guide', 4))
    // The style each page received is the independently generated styleSkill, and all are identical
    expect(stylesSeen.length).toBe(4)
    expect(stylesSeen.every((s) => s === 'STYLE_SKILL_FOR:Jinan Travel Guide')).toBe(true)
  })
})

describe('generate_deck lands pages while generating + retries', () => {
  const deckCall = (n: number): AgentToolCall => ({
    id: 'c',
    name: 'generate_deck',
    input: {
      core_hook: 'h',
      style: 's',
      pages: Array.from({ length: n }, (_, i) => ({
        title: `T${i + 1}`,
        brief: 'b',
        layout: 'data',
        image_queries: [],
      })),
    },
  })

  it('a page fails once but retry succeeds → all pages still land', async () => {
    // Page 2 fails once at first; the second (retry) attempt succeeds
    const { access, getPages } = makeAccess({ failAttempts: { 2: 1 } })
    const skill = createSlidesSkill(access)
    await skill.executeTool(deckCall(3))
    expect(getPages()).toBe(3) // retry as safety net, no missing pages
  })

  it('a page keeps failing (beyond retries) → other pages still land in page order, no blocking', async () => {
    const { access, getPages, landOrder } = makeAccess({ failPages: [2] })
    const skill = createSlidesSkill(access)
    const res = (await skill.executeTool(deckCall(4))) as { output: string }
    // Page 2 hangs; 1/3/4 land (later pages not blocked by page 2)
    expect(getPages()).toBe(3)
    expect(res.output).toContain('page 2 (mock fail) still failed after retry')
    // Landed content doesn't contain PAGE2
    expect(landOrder.some((l) => l.includes('T1'))).toBe(true)
    expect(landOrder.some((l) => l.includes('T3'))).toBe(true)
    expect(landOrder.some((l) => l.includes('T2'))).toBe(false)
    // The progress checklist names pages precisely: page 2 is missing (not misattributed to the last page)
    const ctx = skill.buildContext?.() ?? ''
    expect(ctx).toContain('page 2 "T2"')
  })

  it('a page failing beyond one retry → only 2 calls then skipped (no final-round regeneration)', async () => {
    const { access, getPages, genPageCalls } = makeAccess({ failPages: [2] })
    const skill = createSlidesSkill(access)
    await skill.executeTool(deckCall(3))
    expect(getPages()).toBe(2)
    // One initial attempt + one retry, nothing more
    expect(genPageCalls.filter((n) => n === 2).length).toBe(2)
  })

  it('all pages fail to land initially (zero landed) → retry round still inserts: first page via insertMode, rest via insert_at', async () => {
    const { access, getPages, landOrder } = makeAccess({ landFailOnce: [1, 2] })
    const skill = createSlidesSkill(access)
    const res = (await skill.executeTool(deckCall(2))) as { output: string }
    expect(getPages()).toBe(2)
    expect(res.output).toContain('2/2')
    expect(landOrder.some((l) => l.startsWith('replace:') && l.includes('PAGE1'))).toBe(true)
    expect(landOrder.some((l) => l.startsWith('insert@1:') && l.includes('PAGE2'))).toBe(true)
  })

  it('pages that failed to land → after retry inserted back at their original position via insert_at', async () => {
    // Page 2's HTML generates fine but the first landing (conversion) fails once
    const { access, getPages, landOrder } = makeAccess({ landFailOnce: [2] })
    const skill = createSlidesSkill(access)
    const res = (await skill.executeTool(deckCall(3))) as { output: string }
    // All three pages end up present; page 2 is inserted back at index=1 via insert_at (after page 1)
    expect(getPages()).toBe(3)
    expect(landOrder.some((l) => l.startsWith('insert@1:') && l.includes('PAGE2'))).toBe(true)
    expect(res.output).toContain('3/3')
    const ctx = skill.buildContext?.() ?? ''
    expect(ctx).toContain('all generated')
  })
})

describe('generate_deck in-tool image search', () => {
  it('image_queries are keywords → tool searches images → real URLs passed to generatePageCloud', async () => {
    const { access, imageSearchCalls, imagesSeen } = makeAccess()
    const skill = createSlidesSkill(access)
    const call: AgentToolCall = {
      id: 'c-img',
      name: 'generate_deck',
      input: {
        core_hook: 'h',
        style: 's',
        pages: [
          {
            title: 'Cover',
            brief: 'b',
            layout: 'cover',
            image_queries: ['summer palace kunming lake'],
          },
          { title: 'Body', brief: 'b', layout: 'content', image_queries: [] },
        ],
      },
    }
    await skill.executeTool(call)
    // searchImages should be called once (a keyword only, not a URL)
    expect(imageSearchCalls).toEqual(['summer palace kunming lake'])
    // Page 1 (pageIndex=1) should receive the real searched URL
    expect(imagesSeen[0]).toEqual(['https://mock.img/summer%20palace%20kunming%20lake.jpg'])
    // Page 2 has no image slot
    expect(imagesSeen[1]).toEqual([])
  })

  it('image_queries already http(s) URLs → used directly without re-searching', async () => {
    const { access, imageSearchCalls, imagesSeen } = makeAccess()
    const skill = createSlidesSkill(access)
    const existingUrl = 'https://example.com/preloaded.jpg'
    const call: AgentToolCall = {
      id: 'c-url',
      name: 'generate_deck',
      input: {
        core_hook: 'h',
        style: 's',
        pages: [{ title: 'P1', brief: 'b', layout: 'content', image_queries: [existingUrl] }],
      },
    }
    await skill.executeTool(call)
    // Already a URL; searchImages should not be called
    expect(imageSearchCalls).toEqual([])
    // generatePageCloud should receive the original URL
    expect(imagesSeen[0]).toEqual([existingUrl])
  })

  it('same keyword across pages → searched once per deck, each page gets a different candidate image (deduped across pages)', async () => {
    const { access, imageSearchCalls, imagesSeen } = makeAccess({ searchImagesMulti: 5 })
    const skill = createSlidesSkill(access)
    const call: AgentToolCall = {
      id: 'c-dedup',
      name: 'generate_deck',
      input: {
        core_hook: 'h',
        style: 's',
        pages: [
          { title: 'P1', brief: 'b', layout: 'content', image_queries: ['beijing hutong'] },
          { title: 'P2', brief: 'b', layout: 'content', image_queries: ['Beijing  Hutong'] }, // case/whitespace differences count as the same keyword
          { title: 'P3', brief: 'b', layout: 'content', image_queries: ['beijing hutong'] },
        ],
      },
    }
    await skill.executeTool(call)
    // Three pages sharing one keyword trigger only one search
    expect(imageSearchCalls).toEqual(['beijing hutong'])
    // Each of the three pages gets one image, all distinct
    const firstImages = imagesSeen.map((a) => a[0]).filter(Boolean)
    expect(firstImages.length).toBe(3)
    expect(new Set(firstImages).size).toBe(3)
  })

  it('not enough candidate images to go around → falls back to reusing the first one (an image beats no image)', async () => {
    const { access, imagesSeen } = makeAccess({ searchImagesMulti: 1 })
    const skill = createSlidesSkill(access)
    const call: AgentToolCall = {
      id: 'c-fallback',
      name: 'generate_deck',
      input: {
        core_hook: 'h',
        style: 's',
        pages: [
          { title: 'P1', brief: 'b', layout: 'content', image_queries: ['great wall'] },
          { title: 'P2', brief: 'b', layout: 'content', image_queries: ['great wall'] },
        ],
      },
    }
    await skill.executeTool(call)
    // Only 1 candidate; both pages get an image (reuse fallback)
    expect(imagesSeen.every((a) => a.length === 1)).toBe(true)
  })

  it('topic mode: planned keywords → internal image search end to end → every page prompt receives real URLs', async () => {
    const { access, imageSearchCalls, imagesSeen } = makeAccess({
      searchImagesMulti: 5,
      planImageQueries: ['city skyline night'],
    })
    const skill = createSlidesSkill(access)
    const call: AgentToolCall = {
      id: 'c-topic-img',
      name: 'generate_deck',
      input: { topic: 'City Nightscape', approx_pages: 3, context: 'real source material' },
    }
    await skill.executeTool(call)
    // Three pages planned with the same keyword -> searched only once
    expect(imageSearchCalls).toEqual(['city skyline night'])
    // Every page receives a real URL, all distinct
    const firstImages = imagesSeen.map((a) => a[0]).filter(Boolean)
    expect(firstImages.length).toBe(3)
    expect(new Set(firstImages).size).toBe(3)
    expect(firstImages.every((u) => u!.startsWith('https://mock.img/'))).toBe(true)
  })

  it('image search fails → fail-open, that image slot stays empty, generation completes normally', async () => {
    const { access, getPages } = makeAccess({ searchImagesFail: true })
    const skill = createSlidesSkill(access)
    const call: AgentToolCall = {
      id: 'c-fail',
      name: 'generate_deck',
      input: {
        core_hook: 'h',
        style: 's',
        pages: [{ title: 'P1', brief: 'b', layout: 'content', image_queries: ['beijing hutong'] }],
      },
    }
    const res = (await skill.executeTool(call)) as { output: string; isError?: boolean }
    // Image-search failure doesn't affect the main chain: pages still land
    expect(res.isError).toBeFalsy()
    expect(getPages()).toBe(1)
  })

  it('outline planning fails → plan progress marked error (no longer stuck at running)', async () => {
    const { access } = makeAccess()
    const events: Array<{ stage: string; status?: string }> = []
    access.planDeckOutline = async () => ({
      ok: false,
      error: "outline JSON parse failed: Expected ','",
    })
    access.onProgress = (e) => {
      events.push({ stage: e.stage, status: 'status' in e ? e.status : undefined })
    }
    const skill = createSlidesSkill(access)
    const res = (await skill.executeTool({
      id: 'c-plan-fail',
      name: 'generate_deck',
      input: { topic: 'Test Topic', approx_pages: 10 },
    })) as { output: string; isError?: boolean }
    expect(res.isError).toBeTruthy()
    expect(res.output).toContain('Planning failed')
    const planEvents = events.filter((e) => e.stage === 'plan')
    expect(planEvents.some((e) => e.status === 'running')).toBe(true)
    expect(planEvents.some((e) => e.status === 'error')).toBe(true)
    expect(planEvents.at(-1)?.status).toBe('error')
  })
})

describe('generate_deck Style Skill template persistence', () => {
  const topicCall = (topic: string, opts?: { style_template?: string }): AgentToolCall => ({
    id: 'c-tpl',
    name: 'generate_deck',
    input: {
      topic,
      approx_pages: 2,
      ...(opts?.style_template ? { style_template: opts.style_template } : {}),
    },
  })

  it('saveSidecar is called after generate_deck completes, with styleSkill and topic', async () => {
    const { access, sidecarSaves } = makeAccess()
    const skill = createSlidesSkill(access)
    await skill.executeTool(topicCall('Beijing Travel'))
    expect(sidecarSaves.length).toBe(1)
    expect(sidecarSaves[0]!.styleSkill).toBe('STYLE_SKILL_FOR:Beijing Travel')
    expect(sidecarSaves[0]!.topic).toBe('Beijing Travel')
    expect(sidecarSaves[0]!.createdAt).toBeTruthy()
  })

  it('saveSidecar is not called when all pages fail to generate', async () => {
    const { access, sidecarSaves } = makeAccess({ failPages: [1, 2] })
    const skill = createSlidesSkill(access)
    await skill.executeTool(topicCall('Shanghai Travel'))
    // landedPages=0 -> no sidecar written
    expect(sidecarSaves.length).toBe(0)
  })

  it('state.lastStyleSkill is recorded after generate_deck and usable by save_style_template', async () => {
    const { access, savedTemplates } = makeAccess()
    const skill = createSlidesSkill(access)
    await skill.executeTool(topicCall('Chengdu Cuisine'))
    // Then call save_style_template
    const saveCall: AgentToolCall = {
      id: 'c-save',
      name: 'save_style_template',
      input: { name: 'Chengdu Style' },
    }
    const res = (await skill.executeTool(saveCall)) as { output: string; isError?: boolean }
    expect(res.isError).toBeFalsy()
    expect(res.output).toContain('Chengdu Style')
    expect(savedTemplates['Chengdu Style']?.styleSkill).toBe('STYLE_SKILL_FOR:Chengdu Cuisine')
  })

  it('save_style_template without a prior generate_deck → errors', async () => {
    const { access } = makeAccess()
    const skill = createSlidesSkill(access)
    const saveCall: AgentToolCall = {
      id: 'c-save2',
      name: 'save_style_template',
      input: { name: 'invalid' },
    }
    const res = (await skill.executeTool(saveCall)) as { isError?: boolean }
    expect(res.isError).toBe(true)
  })

  it('list_style_templates lists saved templates', async () => {
    const { access } = makeAccess({
      styleTemplates: {
        'Minimal Blue White': {
          topic: 'Tech',
          styleSkill: 'SK1',
          createdAt: '2024-01-01T00:00:00Z',
        },
      },
    })
    const skill = createSlidesSkill(access)
    const res = (await skill.executeTool({
      id: 'c-list',
      name: 'list_style_templates',
      input: {},
    })) as { output: string }
    expect(res.output).toContain('Minimal Blue White')
    expect(res.output).toContain('Tech')
  })

  it('generate_deck with style_template → skips generateStyleSkill and uses the template styleSkill directly', async () => {
    const { access, stylesSeen } = makeAccess({
      styleTemplates: {
        'Minimal Blue White': {
          topic: 'Tech',
          styleSkill: 'TEMPLATE_SKILL',
          createdAt: '2024-01-01T00:00:00Z',
        },
      },
    })
    const skill = createSlidesSkill(access)
    await skill.executeTool(topicCall('AI Product', { style_template: 'Minimal Blue White' }))
    // The style each page receives should be the template's styleSkill (not generateStyleSkill output)
    expect(stylesSeen.every((s) => s === 'TEMPLATE_SKILL')).toBe(true)
  })

  it('style_template not found → fail-open, degrades to normal generation (no error)', async () => {
    const { access, stylesSeen, getPages } = makeAccess()
    const skill = createSlidesSkill(access)
    const res = (await skill.executeTool(
      topicCall('Urban Planning', { style_template: 'nonexistent-template' }),
    )) as { isError?: boolean; output: string }
    // Fail-open when missing: degrade to normal generation without isError
    expect(res.isError).toBeFalsy()
    expect(getPages()).toBe(2)
    // After degrading, uses the style produced by generateStyleSkill
    expect(stylesSeen.every((s) => s === 'STYLE_SKILL_FOR:Urban Planning')).toBe(true)
  })
})

describe('generate_deck attachment gate', () => {
  const topicCall: AgentToolCall = {
    id: 'c-gate',
    name: 'generate_deck',
    input: { topic: 'Weekly Sales Report', approx_pages: 3, context: 'real source material' },
  }

  it('unread text attachments present → generate_deck refuses and names them', async () => {
    const { access, getPages } = makeAccess()
    access.unreadTextAttachments = () => ['weekly-report-jordan-jul20-24.docx']
    const skill = createSlidesSkill(access)
    const res = (await skill.executeTool(topicCall)) as { isError?: boolean; output: string }
    expect(res.isError).toBe(true)
    expect(res.output).toContain('weekly-report-jordan-jul20-24.docx')
    expect(res.output).toContain('read_attachment')
    // Nothing was generated or landed
    expect(getPages()).toBe(0)
  })

  it('all text attachments read (or none attached) → generate_deck proceeds', async () => {
    const { access, getPages } = makeAccess()
    access.unreadTextAttachments = () => []
    const skill = createSlidesSkill(access)
    const res = (await skill.executeTool(topicCall)) as { isError?: boolean }
    expect(res.isError).toBeFalsy()
    expect(getPages()).toBe(3)
  })
})
