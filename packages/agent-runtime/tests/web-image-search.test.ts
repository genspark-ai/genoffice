/**
 * Verify the web-search + image-search extensions register their tools
 * inside a real pi AgentSession built by createOfficeSession. We don't
 * actually call the network — just confirm the tools are wired.
 */
import { describe, expect, it } from 'vitest'
import {
  createOfficeSession,
  ReactUIAdapter,
} from '../src/index'
import {
  createWebSearchExtension,
  createImageSearchExtension,
  webSearchExtensionDefaults,
  imageSearchExtensionDefaults,
} from '@genoffice/agent-skills'

describe('web-search + image-search extensions', () => {
  it('exports stable defaults', () => {
    expect(webSearchExtensionDefaults.tools).toEqual(['web_search'])
    expect(imageSearchExtensionDefaults.tools).toEqual(['image_search', 'fetch_image'])
  })

  it('registers tools in a real pi AgentSession', async () => {
    const session = await createOfficeSession({
      uiAdapter: new ReactUIAdapter(),
      extensionFactories: [
        createWebSearchExtension(),
        createImageSearchExtension(),
      ],
      extensionMode: 'print',
    })

    try {
      const tools = session.session.getAllTools()
      const names = new Set(tools.map((t) => t.name))
      expect(names.has('web_search')).toBe(true)
      expect(names.has('image_search')).toBe(true)
      expect(names.has('fetch_image')).toBe(true)

      const ws = tools.find((t) => t.name === 'web_search')!
      expect(ws.description.length).toBeGreaterThan(40)
      expect(ws.description).toMatch(/DuckDuckGo/i)

      const is = tools.find((t) => t.name === 'image_search')!
      expect(is.description.length).toBeGreaterThan(40)
      expect(is.description).toMatch(/image/i)

      const fi = tools.find((t) => t.name === 'fetch_image')!
      expect(fi.description.length).toBeGreaterThan(20)
    } finally {
      session.session.dispose()
    }
  }, 15_000)
})
