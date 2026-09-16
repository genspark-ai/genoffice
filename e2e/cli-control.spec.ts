import { test, expect } from '@playwright/test'
import { readFileSync } from 'node:fs'
import { connect } from 'node:net'
import { join, resolve } from 'node:path'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'

/**
 * The genoffice CLI's control channel (`open --slide/--el`, `selection`):
 * the shell publishes userData/control.json, a token-checked local socket
 * takes one request per connection and relays it to the tab's renderer.
 * This drives the socket the way packages/cli/src/control.ts does.
 */

const DECK = resolve(__dirname, '../packages/pptx-engine/tests/fixtures/01_standard_business.pptx')
const DOC = resolve(__dirname, 'assets/justify-pagegap-fr.docx')
const BOOK = resolve(__dirname, '../apps/sheets/fixtures/generated/compatibility-basic.xlsx')

interface Endpoint {
  endpoint: string
  token: string
}

async function waitForEndpoint(userDataDir: string): Promise<Endpoint> {
  const deadline = Date.now() + 30_000
  for (;;) {
    try {
      return JSON.parse(readFileSync(join(userDataDir, 'control.json'), 'utf8')) as Endpoint
    } catch {
      if (Date.now() > deadline) throw new Error('control.json never appeared')
      await new Promise((r) => setTimeout(r, 200))
    }
  }
}

function request(ep: Endpoint, token: string, body: unknown): Promise<string> {
  return new Promise((resolvePromise, reject) => {
    const socket = connect(ep.endpoint)
    let reply = ''
    socket.setEncoding('utf8')
    socket.on('error', reject)
    socket.on('connect', () => socket.write(JSON.stringify({ token, request: body }) + '\n'))
    socket.on('data', (chunk: string) => (reply += chunk))
    socket.on('close', () => resolvePromise(reply))
  })
}

test.describe('cli control channel', () => {
  test('goto selects a slide element and selection reads it back', async () => {
    test.setTimeout(120_000)
    const launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'cli-control',
      openFile: DECK,
    })
    try {
      const editor = await waitForPageWithUrl(launched.app, '://slides/')
      const ep = await waitForEndpoint(launched.userDataDir)
      expect(ep.token).toHaveLength(48)

      const wrongToken = await request(ep, 'nope', { cmd: 'selection', path: DECK })
      expect(wrongToken).toBe('')

      const before = JSON.parse(await request(ep, ep.token, { cmd: 'selection', path: DECK }))
      expect(before).toMatchObject({ ok: true, result: { slide: 0, elements: [] } })

      const outOfRange = JSON.parse(
        await request(ep, ep.token, {
          cmd: 'open',
          path: DECK,
          target: { kind: 'slide', slide: 99 },
        }),
      )
      expect(outOfRange).toMatchObject({ ok: false, error: { reason: 'out_of_range' } })
      expect(outOfRange.error.detail.valid_range).toMatch(/^0-\d+$/)

      const missing = JSON.parse(
        await request(ep, ep.token, {
          cmd: 'open',
          path: DECK,
          target: { kind: 'slide', slide: 1, el: 'e_nope' },
        }),
      )
      expect(missing).toMatchObject({ ok: false, error: { reason: 'target_not_found' } })
      const available: string[] = missing.error.detail.available
      expect(available.length).toBeGreaterThan(0)

      const moved = JSON.parse(
        await request(ep, ep.token, {
          cmd: 'open',
          path: DECK,
          target: { kind: 'slide', slide: 1, el: available[0] },
        }),
      )
      expect(moved).toMatchObject({ ok: true, result: { slide: 1, element: available[0] } })

      const after = JSON.parse(await request(ep, ep.token, { cmd: 'selection', path: DECK }))
      expect(after).toMatchObject({ ok: true, result: { slide: 1, elements: [available[0]] } })
      await expect(editor.locator('.thumb').nth(1)).toHaveClass(/\bactive\b/)

      const wrongKind = JSON.parse(
        await request(ep, ep.token, { cmd: 'open', path: DECK, target: { kind: 'page', page: 1 } }),
      )
      expect(wrongKind).toMatchObject({ ok: false, error: { reason: 'unsupported' } })
    } finally {
      await closeAndSaveVideo(launched, 'cli-control')
    }
  })

  test('goto selects a docx block and selection reports the block range', async () => {
    test.setTimeout(120_000)
    const launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'cli-control-docs',
      openFile: DOC,
    })
    try {
      await waitForPageWithUrl(launched.app, '://docs/')
      const ep = await waitForEndpoint(launched.userDataDir)
      const moved = JSON.parse(
        await request(ep, ep.token, {
          cmd: 'open',
          path: DOC,
          target: { kind: 'block', block: 2 },
        }),
      )
      expect(moved).toMatchObject({ ok: true, result: { block: 2 } })
      const sel = JSON.parse(await request(ep, ep.token, { cmd: 'selection', path: DOC }))
      expect(sel).toMatchObject({ ok: true, result: { blocks: [2, 2] } })
      const tooFar = JSON.parse(
        await request(ep, ep.token, {
          cmd: 'open',
          path: DOC,
          target: { kind: 'block', block: 9999 },
        }),
      )
      expect(tooFar).toMatchObject({ ok: false, error: { reason: 'out_of_range' } })
    } finally {
      await closeAndSaveVideo(launched, 'cli-control-docs')
    }
  })

  test('goto activates a worksheet range and selection reads it back', async () => {
    test.setTimeout(120_000)
    const launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'cli-control-sheets',
      openFile: BOOK,
    })
    try {
      await waitForPageWithUrl(launched.app, '://sheets/')
      const ep = await waitForEndpoint(launched.userDataDir)
      const moved = JSON.parse(
        await request(ep, ep.token, {
          cmd: 'open',
          path: BOOK,
          target: { kind: 'range', range: 'B2:C3' },
        }),
      )
      expect(moved, JSON.stringify(moved)).toMatchObject({ ok: true, result: { range: 'B2:C3' } })
      expect(typeof moved.result.sheet).toBe('string')
      const sel = JSON.parse(await request(ep, ep.token, { cmd: 'selection', path: BOOK }))
      expect(sel).toMatchObject({ ok: true, result: { sheet: moved.result.sheet, range: 'B2:C3' } })
      expect(sel.result.values).toHaveLength(2)
      const noSheet = JSON.parse(
        await request(ep, ep.token, {
          cmd: 'open',
          path: BOOK,
          target: { kind: 'range', sheet: 'NoSuchSheet', range: 'A1' },
        }),
      )
      expect(noSheet).toMatchObject({ ok: false, error: { reason: 'sheet_not_found' } })
      expect(noSheet.error.detail.sheets).toContain(moved.result.sheet)
    } finally {
      await closeAndSaveVideo(launched, 'cli-control-sheets')
    }
  })
})
