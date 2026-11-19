import { describe, expect, it } from 'vitest'
import type { AgentStreamCallbacks, AgentTransport } from '@genoffice/agent-core'
import {
  buildPageWriterRequest,
  extractPageHtml,
  streamPage,
  PAGE_MAX_CHARS,
} from '../src/renderer/ai/page-writer'

const PAGE =
  '<!doctype html>\n<html lang="en"><head><title>T</title></head><body><p>x</p></body></html>'

/** scripted transport: feeds the callbacks the given events on the next tick */
function fakeTransport(
  script: (cb: AgentStreamCallbacks) => void,
  onCancel?: () => void,
): AgentTransport {
  return {
    stream: (_req, cb) => {
      queueMicrotask(() => script(cb))
      return { cancel: () => onCancel?.() }
    },
  }
}

describe('extractPageHtml', () => {
  it('drops chatter and fences around the document', () => {
    expect(extractPageHtml(`Here is the page:\n\`\`\`html\n${PAGE}\n\`\`\`\nDone.`)).toEqual({
      html: PAGE,
      complete: true,
    })
  })
  it('returns a partial document without closing tag as incomplete', () => {
    const r = extractPageHtml('```html\n<!doctype html>\n<html><body><p>half')
    expect(r.complete).toBe(false)
    expect(r.html).toBe('<!doctype html>\n<html><body><p>half')
  })
  it('normalizes CRLF and keeps content up to the last </html>', () => {
    const r = extractPageHtml('<html>\r\n<body>a</body></html>\r\n<!-- trailing -->')
    expect(r).toEqual({ html: '<html>\n<body>a</body></html>', complete: true })
  })
})

describe('buildPageWriterRequest', () => {
  it('gives the design writer the brief without the alternatives, plus context and request', () => {
    const { system, user } = buildPageWriterRequest(
      {
        kind: 'design',
        brief: {
          core_hook: 'Hook',
          style: { tone: 'bold', palette: {}, typography: {} },
          alternatives: [{ tone: 'soft', palette: {}, typography: {} }],
          sections: [{ title: 'Hero', brief: 'b' }],
          version: 1,
        },
        instruction: 'make a landing page',
        context: 'founded 1999',
      },
      '\n\nAnswer in French.',
    )
    expect(system).toContain('--brief-primary')
    expect(system).toContain('start at <!doctype html>')
    expect(system.endsWith('Answer in French.')).toBe(true)
    expect(user).toContain('"core_hook": "Hook"')
    expect(user).not.toContain('soft')
    expect(user).toContain('founded 1999')
    expect(user).toContain('make a landing page')
  })
  it('gives the content writer the plan and title', () => {
    const { system, user } = buildPageWriterRequest(
      { kind: 'content', title: 'Guide', plan: 'steps', instruction: 'write it' },
      '',
    )
    expect(system).toContain('single-column')
    expect(user).toContain('Title: Guide')
    expect(user).toContain('steps')
  })
})

describe('streamPage', () => {
  it('assembles a complete page from deltas and reports progress', async () => {
    const progress: string[] = []
    const outcome = await streamPage({
      transport: fakeTransport((cb) => {
        cb.onDelta('Sure!\n```html\n<!doctype html>\n<html><body>')
        cb.onDelta('<p>x</p></body></html>\n```')
        cb.onStopReason?.('end_turn')
        cb.onDone()
      }),
      system: 's',
      user: 'u',
      onProgress: (html) => progress.push(html),
    })
    expect(outcome).toEqual({
      status: 'complete',
      html: '<!doctype html>\n<html><body><p>x</p></body></html>',
    })
    expect(progress[0]).toBe('<!doctype html>\n<html><body>')
  })
  it('a dropped connection yields the partial page, an empty one yields empty', async () => {
    const partial = await streamPage({
      transport: fakeTransport((cb) => {
        cb.onDelta('<!doctype html><html><body><p>half')
        cb.onError('connection dropped')
      }),
      system: 's',
      user: 'u',
    })
    expect(partial).toEqual({
      status: 'partial',
      html: '<!doctype html><html><body><p>half',
      reason: 'error',
      error: 'connection dropped',
    })
    const empty = await streamPage({
      transport: fakeTransport((cb) => cb.onError('gateway 502')),
      system: 's',
      user: 'u',
    })
    expect(empty).toEqual({ status: 'empty', error: 'gateway 502' })
  })
  it('a finished page is complete even when the turn stopped at max_tokens', async () => {
    const outcome = await streamPage({
      transport: fakeTransport((cb) => {
        cb.onDelta(PAGE)
        cb.onStopReason?.('max_tokens')
        cb.onDone()
      }),
      system: 's',
      user: 'u',
    })
    expect(outcome).toEqual({ status: 'complete', html: PAGE })
  })
  it('max_tokens marks a page that ends before </html> as partial', async () => {
    const outcome = await streamPage({
      transport: fakeTransport((cb) => {
        cb.onDelta('<html><body>')
        cb.onStopReason?.('max_tokens')
        cb.onDone()
      }),
      system: 's',
      user: 'u',
    })
    expect(outcome).toMatchObject({ status: 'partial', reason: 'max_tokens' })
  })
  it('stop cancels the transport and keeps what arrived; later deltas are ignored', async () => {
    let cancelled = false
    const progress: string[] = []
    const controller = new AbortController()
    const promise = streamPage({
      transport: fakeTransport(
        (cb) => {
          cb.onDelta('<html><body><p>a')
          controller.abort()
          cb.onDelta('<p>late</p>')
          cb.onDone()
        },
        () => {
          cancelled = true
        },
      ),
      system: 's',
      user: 'u',
      signal: controller.signal,
      onProgress: (html) => progress.push(html),
    })
    expect(await promise).toMatchObject({
      status: 'partial',
      reason: 'stopped',
      html: '<html><body><p>a',
    })
    expect(cancelled).toBe(true)
    expect(progress).toEqual(['<html><body><p>a'])
  })
  it('an oversized reply is cut off as partial', async () => {
    const outcome = await streamPage({
      transport: fakeTransport((cb) => {
        cb.onDelta('<html>' + 'x'.repeat(PAGE_MAX_CHARS))
        cb.onDone()
      }),
      system: 's',
      user: 'u',
    })
    expect(outcome).toMatchObject({ status: 'partial', reason: 'max_tokens' })
  })
})
