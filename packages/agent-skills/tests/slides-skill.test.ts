/**
 * Tests for slides-skill (W7 deliverable).
 */

import { describe, expect, it, beforeEach, vi } from 'vitest'
import { ReactUIAdapter } from '@genoffice/agent-runtime'
import {
  createSlidesSkillExtension,
  createReadSlideTool,
  createPlanDeckTool,
  createExecuteSlideScriptTool,
  createRegenerateSlideTool,
  type SlidesEditor,
  type SlideContent,
} from '../src/extensions/slides-skill'

class MockSlidesEditor implements SlidesEditor {
  pages: SlideContent[]
  appliedOps: { index: number; count: number }[] = []
  regenerated: { index: number; title: string; body: string }[] = []
  appliedPlans: { title: string; pages: number }[] = []

  constructor() {
    this.pages = [
      { index: 0, title: 'Welcome', body: 'Intro to the deck', notes: '', layoutName: 'Title', slots: [{ kind: 'title', text: 'Welcome' }] },
      { index: 1, title: 'Agenda', body: '1. Goals 2. Plan 3. Risks', notes: '', layoutName: 'Title + Content', slots: [{ kind: 'title', text: 'Agenda' }, { kind: 'body', text: '1. Goals 2. Plan 3. Risks' }] },
      { index: 2, title: 'Q1 Results', body: 'Revenue grew 23%', notes: 'Mention FX impact', layoutName: 'Title + Content', slots: [{ kind: 'title', text: 'Q1 Results' }, { kind: 'body', text: 'Revenue grew 23%' }] },
    ]
  }

  getDeckSummary() {
    return {
      pageCount: this.pages.length,
      title: this.pages[0]?.title ?? '',
      pages: this.pages.map((p) => ({ index: p.index, title: p.title, layoutName: p.layoutName, hasImages: false, hasCharts: false })),
    }
  }

  readSlide(index: number): SlideContent | null {
    return this.pages[index] ?? null
  }

  regenerateSlide(index: number, plan: { title: string; body: string; notes?: string }) {
    this.regenerated.push({ index, title: plan.title, body: plan.body })
    if (this.pages[index]) {
      this.pages[index]!.title = plan.title
      this.pages[index]!.body = plan.body
    }
  }

  executeSlideScript(index: number, script: { operations: ReadonlyArray<unknown> }) {
    this.appliedOps.push({ index, count: script.operations.length })
    return { applied: script.operations.length, errors: [] }
  }

  applyDeckPlan(plan: { title: string; pages: Array<{ title: string }> }) {
    this.appliedPlans.push({ title: plan.title, pages: plan.pages.length })
    const oldLen = this.pages.length
    this.pages = plan.pages.map((p, i) => ({ index: i, title: p.title, body: '', notes: '', layoutName: 'Blank', slots: [] }))
    return { updated: 0, created: this.pages.length - oldLen }
  }
}

describe('slides-skill: read_slide', () => {
  let adapter: ReactUIAdapter
  let editor: MockSlidesEditor

  beforeEach(() => {
    adapter = new ReactUIAdapter()
    editor = new MockSlidesEditor()
    adapter.setEditorInstance(editor)
  })

  it('returns title, body, notes, slot summary', async () => {
    const tool = createReadSlideTool({ uiAdapter: adapter })
    const result = await tool.execute('id', { index: 1 }, undefined, undefined, {} as never)
    const text = (result.content[0] as { type: 'text'; text: string }).text
    expect(text).toContain('"Agenda"')
    expect(text).toContain('layout: Title + Content')
    expect(text).toContain('1. Goals 2. Plan 3. Risks')
    expect(text).toContain('[title')
    expect(text).toContain('[body')
    expect(result.details.found).toBe(true)
  })

  it('reports out-of-range slide', async () => {
    const tool = createReadSlideTool({ uiAdapter: adapter })
    const result = await tool.execute('id', { index: 99 }, undefined, undefined, {} as never)
    expect((result.content[0] as { type: 'text'; text: string }).text).toContain('not found')
    expect(result.details.found).toBe(false)
  })
})

describe('slides-skill: plan_deck', () => {
  let adapter: ReactUIAdapter
  let editor: MockSlidesEditor

  beforeEach(() => {
    adapter = new ReactUIAdapter()
    editor = new MockSlidesEditor()
    adapter.setEditorInstance(editor)
  })

  it('applies the plan when host supports applyDeckPlan', async () => {
    const tool = createPlanDeckTool({ uiAdapter: adapter })
    const result = await tool.execute('id', {
      title: 'Q1 2026 Review',
      audience: 'Executives',
      pages: [
        { title: 'Welcome', keyPoints: ['intro'] },
        { title: 'Goals', keyPoints: ['grow revenue'] },
      ],
    }, undefined, undefined, {} as never)
    expect(editor.appliedPlans).toHaveLength(1)
    expect(editor.appliedPlans[0]!.title).toBe('Q1 2026 Review')
    expect(result.details.applied).toBe(true)
  })

  it('returns plan as text when host missing applyDeckPlan', async () => {
    const editorNoApply: SlidesEditor = {
      getDeckSummary: () => ({ pageCount: 0, title: '', pages: [] }),
      readSlide: () => null,
    }
    adapter.setEditorInstance(editorNoApply)
    const tool = createPlanDeckTool({ uiAdapter: adapter })
    const result = await tool.execute('id', {
      title: 'X', audience: 'devs',
      pages: [{ title: 'A', keyPoints: ['1', '2'] }],
    }, undefined, undefined, {} as never)
    expect(result.details.applied).toBe(false)
    expect((result.content[0] as { type: 'text'; text: string }).text).toContain('not applied')
  })
})

describe('slides-skill: execute_slide_script', () => {
  it('counts operations applied', async () => {
    const adapter = new ReactUIAdapter()
    adapter.setEditorInstance(new MockSlidesEditor())
    const tool = createExecuteSlideScriptTool({ uiAdapter: adapter })
    const result = await tool.execute('id', { index: 0, operations: [{ move: 'left' }, { style: 'bold' }] }, undefined, undefined, {} as never)
    expect(result.details.applied).toBe(2)
    expect(result.details.errors).toEqual([])
  })

  it('handles host without executeSlideScript', async () => {
    const adapter = new ReactUIAdapter()
    const editor: SlidesEditor = { getDeckSummary: () => ({ pageCount: 0, title: '', pages: [] }), readSlide: () => null }
    adapter.setEditorInstance(editor)
    const tool = createExecuteSlideScriptTool({ uiAdapter: adapter })
    const result = await tool.execute('id', { index: 0, operations: [] }, undefined, undefined, {} as never)
    expect(result.details.errors).toContain('not_supported')
  })
})

describe('slides-skill: regenerate_slide', () => {
  it('replaces content and records the change', async () => {
    const adapter = new ReactUIAdapter()
    const editor = new MockSlidesEditor()
    adapter.setEditorInstance(editor)
    const tool = createRegenerateSlideTool({ uiAdapter: adapter })
    const result = await tool.execute('id', { index: 0, title: 'New Welcome', body: 'New intro' }, undefined, undefined, {} as never)
    expect(editor.regenerated).toHaveLength(1)
    expect(editor.regenerated[0]!.title).toBe('New Welcome')
    expect(result.details.regenerated).toBe(true)
  })
})

describe('slides-skill extension factory', () => {
  it('registers all 4 slides tools by default', () => {
    const adapter = new ReactUIAdapter()
    adapter.setEditorInstance(new MockSlidesEditor())
    const registered: string[] = []
    const pi = { registerTool: (t: { name: string }) => registered.push(t.name), on: () => {} } as never
    createSlidesSkillExtension({ uiAdapter: adapter })(pi)
    expect(registered).toEqual(['read_slide', 'plan_deck', 'execute_slide_script', 'regenerate_slide'])
  })

  it('enabledTools restricts the set', () => {
    const adapter = new ReactUIAdapter()
    adapter.setEditorInstance(new MockSlidesEditor())
    const registered: string[] = []
    const pi = { registerTool: (t: { name: string }) => registered.push(t.name), on: () => {} } as never
    createSlidesSkillExtension({ uiAdapter: adapter, enabledTools: ['read_slide'] })(pi)
    expect(registered).toEqual(['read_slide'])
  })

  it('loads into a real pi session', async () => {
    const { createOfficeSession } = await import('@genoffice/agent-runtime')
    const adapter = new ReactUIAdapter()
    adapter.setEditorInstance(new MockSlidesEditor())
    const { session, dispose } = await createOfficeSession({
      cwd: process.cwd(),
      extensionFactories: [createSlidesSkillExtension({ uiAdapter: adapter })],
    })
    const names = session.getAllTools().map((t) => t.name)
    expect(names).toContain('read_slide')
    expect(names).toContain('plan_deck')
    dispose()
  }, 30_000)
})
