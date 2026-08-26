/** set_speaker_notes tool: writes/clears a page's speaker notes through DeckAccess. */
import { describe, it, expect, vi } from 'vitest'
import { createSlidesSkill, type DeckAccess } from '../src/renderer/ai/slides-skill'
import type { RenderSlide } from '@genoffice/pptx-render'
import type { AgentToolCall } from '../src/shared/ipc'

const slide = {
  widthPx: 1280,
  heightPx: 720,
  nodes: [],
} as unknown as RenderSlide

function mkAccess(overrides: Partial<DeckAccess> = {}): DeckAccess {
  return {
    getSlides: () => [slide],
    getCurrent: () => 0,
    getSelectedIds: () => [],
    applySlide: () => {},
    applyDeck: () => {},
    fitWidthPx: 1280,
    retryBackoffMs: 0,
    ...overrides,
  } as unknown as DeckAccess
}

const notesCall = (input: Record<string, unknown>): AgentToolCall => ({
  id: 'notes',
  name: 'set_speaker_notes',
  input,
})

describe('set_speaker_notes', () => {
  it('writes notes for a valid page via access.setSpeakerNotes', async () => {
    const setSpeakerNotes = vi.fn(async () => true)
    const r = await createSlidesSkill(mkAccess({ setSpeakerNotes })).executeTool!(
      notesCall({ slideIndex: 0, text: 'Talk slowly.\nEmphasize the ROI slide.' }),
    )
    expect(setSpeakerNotes).toHaveBeenCalledWith(0, 'Talk slowly.\nEmphasize the ROI slide.')
    expect(r.isError).toBeUndefined()
    expect(r.mutated).toBe(true)
    expect(r.output).toContain('Wrote speaker notes')
  })

  it('rejects an out-of-range slideIndex without calling the writer', async () => {
    const setSpeakerNotes = vi.fn(async () => true)
    const r = await createSlidesSkill(mkAccess({ setSpeakerNotes })).executeTool!(
      notesCall({ slideIndex: 5, text: 'x' }),
    )
    expect(r.isError).toBe(true)
    expect(setSpeakerNotes).not.toHaveBeenCalled()
  })

  it('clears notes when text is empty', async () => {
    const setSpeakerNotes = vi.fn(async () => true)
    const r = await createSlidesSkill(mkAccess({ setSpeakerNotes })).executeTool!(
      notesCall({ slideIndex: 0, text: '' }),
    )
    expect(setSpeakerNotes).toHaveBeenCalledWith(0, '')
    expect(r.output).toContain('Cleared speaker notes')
  })

  it('reports failure when the write fails', async () => {
    const setSpeakerNotes = vi.fn(async () => false)
    const r = await createSlidesSkill(mkAccess({ setSpeakerNotes })).executeTool!(
      notesCall({ slideIndex: 0, text: 'x' }),
    )
    expect(r.isError).toBe(true)
    expect(r.output).toContain('failed')
  })
})
