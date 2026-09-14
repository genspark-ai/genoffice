/**
 * Smoke tests for the AI channel split after the M1 refactor:
 *   - every former placeholder channel throws WebUnsupportedError
 *   - the new media handlers (style-template) succeed with a fixture
 *
 * Run with: `tsx --test src/ai/__tests__/web-unsupported.test.ts`
 *
 * The tests use the shared registry directly (no Electron / HTTP), so they
 * run in any plain Node environment.
 */
import { test } from 'node:test'
import assert from 'node:assert/strict'

import { getHandler, handlerCount } from '../../common/index.js'
import { registerDocAiSkillHandlers } from '../doc-skill.js'
import { registerSheetAiSkillHandlers } from '../sheet-skill.js'
import { registerSlideAiSkillHandlers } from '../slide-skill.js'
import { registerAiMediaSkillHandlers } from '../media-skill.js'
import { WebUnsupportedError } from '../errors.js'

const WEB_DOC_CHANNELS = [
  'ai:doc-write-continue',
  'ai:doc-write-expand',
  'ai:doc-write-shrink',
  'ai:doc-write-rewrite',
  'ai:doc-tone-adjust',
  'ai:doc-format-suggest',
  'ai:doc-format-apply',
]
const WEB_SHEET_CHANNELS = [
  'ai:sheets-formula-suggest',
  'ai:sheets-formula-explain',
  'ai:sheets-data-analyze',
  'ai:sheets-trend-predict',
  'ai:sheets-chart-suggest',
  'ai:sheets-chart-create',
  'ai:sheets-data-clean',
  'ai:sheets-data-fill',
]
const WEB_SLIDE_CHANNELS = [
  'ai:slides-generate-outline',
  'ai:slides-generate-content',
  'ai:slides-generate-full',
  'ai:slides-style-apply',
  'ai:slides-style-consistency',
  'ai:slides-generate-notes',
  'ai:slides-translate',
]
const MEDIA_CHANNELS = [
  'ai:save-style-template',
  'ai:list-style-templates',
  'ai:load-style-template',
]

function snapshotRegistered(): number {
  return handlerCount()
}

test('every former placeholder channel is registered exactly once', () => {
  const before = snapshotRegistered()
  registerDocAiSkillHandlers()
  registerSheetAiSkillHandlers()
  registerSlideAiSkillHandlers()
  registerAiMediaSkillHandlers()
  const expectedNew =
    WEB_DOC_CHANNELS.length +
    WEB_SHEET_CHANNELS.length +
    WEB_SLIDE_CHANNELS.length +
    MEDIA_CHANNELS.length
  assert.equal(snapshotRegistered() - before, expectedNew)
})

for (const channel of [...WEB_DOC_CHANNELS, ...WEB_SHEET_CHANNELS, ...WEB_SLIDE_CHANNELS]) {
  test(`${channel} throws WebUnsupportedError`, () => {
    const handler = getHandler(channel)
    assert.ok(handler, `handler for ${channel} should be registered`)
    let caught: unknown
    try {
      handler!(null, { sample: 'input' })
    } catch (err) {
      caught = err
    }
    assert.ok(caught, `expected ${channel} to throw`)
    assert.ok(caught instanceof WebUnsupportedError, `expected WebUnsupportedError, got ${String(caught)}`)
    assert.equal((caught as WebUnsupportedError).code, 'WEB_UNSUPPORTED')
    assert.equal((caught as WebUnsupportedError).channel, channel)
    assert.equal((caught as WebUnsupportedError).reason, 'renderer-side skill')
  })
}

test('ai:save-style-template persists a record and returns it', async () => {
  const save = getHandler('ai:save-style-template')!
  const result = (await save(null, {
    name: 'Smoke Test Template',
    style: { colors: ['#3498db'], fonts: { title: 'Inter' } },
  })) as { ok: boolean; template: { id: string; name: string; style: Record<string, unknown> } }
  assert.equal(result.ok, true)
  assert.equal(result.template.name, 'Smoke Test Template')
  assert.deepEqual(result.template.style.colors, ['#3498db'])

  const load = getHandler('ai:load-style-template')!
  const loaded = (await load(null, { id: result.template.id })) as {
    template: { id: string; name: string }
  }
  assert.equal(loaded.template.id, result.template.id)

  const list = getHandler('ai:list-style-templates')!
  const listed = (await list(null)) as { templates: Array<{ id: string }> }
  assert.ok(listed.templates.some(t => t.id === result.template.id))
})

test('ai:load-style-template rejects missing ids with NotFoundError', async () => {
  const load = getHandler('ai:load-style-template')!
  let caught: unknown
  try {
    await load(null, { id: 'definitely-not-here' })
  } catch (err) {
    caught = err
  }
  assert.ok(caught)
  assert.equal((caught as { code: string }).code, 'NOT_FOUND')
})
