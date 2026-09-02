import { readFileSync } from 'node:fs'

import { describe, expect, it } from 'vitest'

import { workbookFileSchema, workbookRangeResultSchema } from '../src/shared/desktop-api'

/**
 * Tripwire for the preload-whitelist field-drop trap: the sidecar wire schema
 * gained a field but the hand-written preload validator never mentions it, so
 * the field silently vanishes between main and renderer. Three regressions
 * shipped this way (the recalc budget fields, sourceXmlBytes, arrayRef).
 *
 * A name appearing in the source is a necessary, not sufficient, condition —
 * but every shipped instance of the trap was a field the preload never named
 * at all, which this catches at unit-test speed.
 */
const preloadSource = readFileSync(new URL('../src/preload/index.ts', import.meta.url), 'utf8')

function objectKeys(schema: unknown, path: string, out: Map<string, string>): void {
  const node = schema as {
    shape?: Record<string, unknown>
    element?: unknown
    unwrap?: () => unknown
    _def?: { innerType?: unknown; schema?: unknown }
  }
  if (node.shape) {
    for (const [key, child] of Object.entries(node.shape)) {
      if (!out.has(key)) out.set(key, `${path}.${key}`)
      objectKeys(child, `${path}.${key}`, out)
    }
    return
  }
  if (node.element) {
    objectKeys(node.element, `${path}[]`, out)
    return
  }
  const inner = node._def?.innerType ?? node._def?.schema
  if (inner) objectKeys(inner, path, out)
}

describe('preload whitelist names every sidecar wire field', () => {
  it('covers workbookRangeResultSchema', () => {
    const keys = new Map<string, string>()
    objectKeys(workbookRangeResultSchema, 'rangeResult', keys)
    expect(keys.size).toBeGreaterThan(30)
    const missing = [...keys.entries()]
      .filter(([key]) => !preloadSource.includes(key))
      .map(([, where]) => where)
    expect(missing).toEqual([])
  })

  it('covers the visual objects of workbookFileSchema', () => {
    const keys = new Map<string, string>()
    objectKeys(workbookFileSchema, 'file', keys)
    // mediaDataUrl is renderer-only (session previews); the sidecar never
    // sends it, so the preload rightly does not name it.
    const visualKeys = [...keys.entries()].filter(
      ([key, where]) => where.startsWith('file.visuals[]') && key !== 'mediaDataUrl',
    )
    expect(visualKeys.length).toBeGreaterThan(20)
    // OLE embeds ride the same visual record; the preload must name them.
    expect(visualKeys.map(([key]) => key)).toContain('progId')
    const missing = visualKeys
      .filter(([key]) => !preloadSource.includes(key))
      .map(([, where]) => where)
    expect(missing).toEqual([])
  })
})
