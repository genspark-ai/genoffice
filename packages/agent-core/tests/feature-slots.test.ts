import { describe, expect, it } from 'vitest'
import { SLOT_LIMITS, slotsAt, validateSlot } from '../src/feature-slots'
import type { FeatureRecord } from '../src/features-skill'

const button = (action: unknown) => ({ type: 'button', label: 'Go', action })

function record(id: string, place: string): FeatureRecord {
  return {
    id,
    name: id,
    description: 'x',
    kind: 'slot',
    slot: { at: place as never, element: { type: 'text', value: 'hi' } },
    createdAt: 'x',
    updatedAt: 'x',
  }
}

describe('slot contract', () => {
  it('accepts a small tree and reports the place', () => {
    const result = validateSlot({
      at: 'home.sidebar',
      element: {
        type: 'group',
        children: [
          { type: 'text', value: 'Files', tone: 'muted', size: 'caption' },
          { type: 'button', label: 'Open', action: { kind: 'open', path: '/tmp/a.docx' } },
          { type: 'divider' },
        ],
      },
    })
    expect(result.ok).toBe(true)
    if (result.ok) expect(result.slot.at).toBe('home.sidebar')
  })

  it('refuses a place no host draws, and names the ones that exist', () => {
    // ribbon.button is not in the list: a place is only accepted once a host
    // actually draws it, so a feature cannot declare a slot nobody renders
    const result = validateSlot({ at: 'ribbon.button', element: { type: 'text', value: 'x' } })
    expect(result.ok).toBe(false)
    if (!result.ok) {
      expect(result.errors[0]!.field).toBe('slot.at')
      expect(result.errors[0]!.message).toContain('editor.side')
      expect(result.errors[0]!.message).not.toContain('ribbon.button,')
    }
  })

  it('accepts the editor place the docs app draws', () => {
    const result = validateSlot({
      at: 'editor.side',
      element: {
        type: 'button',
        label: 'Tidy headings',
        action: { kind: 'ask-assistant', prompt: 'tidy' },
      },
    })
    expect(result.ok).toBe(true)
  })

  it('names the field of every broken element', () => {
    const result = validateSlot({
      at: 'home.card',
      element: {
        type: 'group',
        children: [
          { type: 'button', label: 'No action' },
          { type: 'nope' },
          { type: 'input', action: { kind: 'open' } },
        ],
      },
    })
    expect(result.ok).toBe(false)
    if (!result.ok) {
      const fields = result.errors.map((error) => error.field)
      expect(fields).toContain('slot.element.children[0].action')
      expect(fields).toContain('slot.element.children[1].type')
      expect(fields).toContain('slot.element.children[2].action.path')
    }
  })

  it('enforces the limits', () => {
    const deep = {
      type: 'group',
      children: [
        {
          type: 'group',
          children: [
            {
              type: 'group',
              children: [{ type: 'group', children: [{ type: 'text', value: 'x' }] }],
            },
          ],
        },
      ],
    }
    const tooDeep = validateSlot({ at: 'home.card', element: deep })
    expect(tooDeep.ok).toBe(false)

    const longText = validateSlot({
      at: 'home.card',
      element: { type: 'text', value: 'x'.repeat(SLOT_LIMITS.maxTextChars + 1) },
    })
    expect(longText.ok).toBe(false)

    const tooManyItems = validateSlot({
      at: 'home.card',
      element: {
        type: 'list',
        items: Array.from({ length: SLOT_LIMITS.maxItems + 1 }, (_, i) => ({ label: 'item ' + i })),
      },
    })
    expect(tooManyItems.ok).toBe(false)

    const tooManyElements = validateSlot({
      at: 'home.card',
      element: {
        type: 'group',
        children: Array.from({ length: SLOT_LIMITS.maxElements }, () => ({ type: 'divider' })),
      },
    })
    expect(tooManyElements.ok).toBe(false)
    if (!tooManyElements.ok) {
      expect(tooManyElements.errors.some((error) => error.message.includes('more than the'))).toBe(
        true,
      )
    }
  })

  it('rejects a non-object slot', () => {
    expect(validateSlot('home.card').ok).toBe(false)
    expect(validateSlot({ at: 'home.card' }).ok).toBe(false)
    expect(validateSlot(undefined).ok).toBe(false)
  })

  it('accepts the four declared actions and nothing else', () => {
    for (const action of [
      { kind: 'open', path: '/tmp/a.docx' },
      { kind: 'ask-assistant', prompt: 'add a feature' },
      { kind: 'panel', id: 'file-switcher' },
      { kind: 'run-feature', id: 'file-switcher' },
    ]) {
      expect(validateSlot({ at: 'home.card', element: button(action) }).ok).toBe(true)
    }
    expect(validateSlot({ at: 'home.card', element: button({ kind: 'rm -rf' }) }).ok).toBe(false)
    expect(validateSlot({ at: 'home.card', element: button({ kind: 'open' }) }).ok).toBe(false)
  })

  it('accepts a host-filled file list and its placeholders', () => {
    const result = validateSlot({
      at: 'home.panel',
      element: {
        type: 'file-list',
        source: 'recent-files',
        limit: 6,
        emptyText: 'No recent files',
        action: { kind: 'open', path: '$path' },
      },
    })
    expect(result.ok).toBe(true)
  })

  it('refuses an unknown source, a bad limit and a bad action on a file list', () => {
    const unknown = validateSlot({
      at: 'home.panel',
      element: { type: 'file-list', source: 'ssh-keys' },
    })
    expect(unknown.ok).toBe(false)
    if (!unknown.ok) expect(unknown.errors[0]!.field).toBe('slot.element.source')

    const badLimit = validateSlot({
      at: 'home.panel',
      element: { type: 'file-list', source: 'starred-files', limit: 500 },
    })
    expect(badLimit.ok).toBe(false)
    if (!badLimit.ok) expect(badLimit.errors[0]!.field).toBe('slot.element.limit')

    const badAction = validateSlot({
      at: 'home.panel',
      element: { type: 'file-list', source: 'recent-files', action: { kind: 'open' } },
    })
    expect(badAction.ok).toBe(false)
    if (!badAction.ok) expect(badAction.errors[0]!.field).toBe('slot.element.action.path')
  })

  it('accepts a browser view and refuses an address the shell will not open', () => {
    const ok = validateSlot({
      at: 'home.panel',
      element: { type: 'view', view: 'browser', url: 'https://example.com/news', title: 'News' },
    })
    expect(ok.ok).toBe(true)

    const local = validateSlot({
      at: 'home.panel',
      element: { type: 'view', view: 'browser', url: 'http://localhost:5173' },
    })
    expect(local.ok).toBe(true)

    const file = validateSlot({
      at: 'home.panel',
      element: { type: 'view', view: 'browser', url: 'file:///etc/passwd' },
    })
    expect(file.ok).toBe(false)
    if (!file.ok) expect(file.errors[0]!.field).toBe('slot.element.url')

    const other = validateSlot({
      at: 'home.panel',
      element: { type: 'view', view: 'terminal', url: 'https://example.com' },
    })
    expect(other.ok).toBe(false)
    if (!other.ok) expect(other.errors[0]!.field).toBe('slot.element.view')

    const tiny = validateSlot({
      at: 'home.panel',
      element: { type: 'view', view: 'browser', url: 'https://example.com', minHeight: 20 },
    })
    expect(tiny.ok).toBe(false)
  })

  it('picks the slots of one place in registry order', () => {
    const features = [
      record('a', 'home.sidebar'),
      record('b', 'home.card'),
      record('c', 'home.sidebar'),
    ]
    expect(slotsAt(features, 'home.sidebar').map((f) => f.id)).toEqual(['a', 'c'])
    expect(slotsAt(features, 'home.panel')).toEqual([])
  })
})
