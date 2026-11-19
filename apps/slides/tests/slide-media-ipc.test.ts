import { tmpdir } from 'node:os'
import { beforeAll, beforeEach, describe, expect, it, vi } from 'vitest'

const handlers = new Map<string, (...args: never[]) => unknown>()
const broadcasts: Array<{ id: number; channel: string; payload: unknown }> = []

vi.mock('electron', () => ({
  app: {
    getPath: () => tmpdir(),
    getName: () => 'slides-test',
    getVersion: () => '0',
    whenReady: () => Promise.resolve(),
  },
  dialog: {},
  clipboard: { writeBuffer: () => {} },
  ipcMain: {
    handle: (channel: string, fn: (...args: never[]) => unknown) => handlers.set(channel, fn),
    on: () => {},
    removeHandler: () => {},
  },
  BrowserWindow: class {},
  webContents: {
    getAllWebContents: () => [],
    fromId: (id: number) => ({
      send: (channel: string, payload: unknown) => broadcasts.push({ id, channel, payload }),
    }),
  },
  Menu: { buildFromTemplate: () => ({}), setApplicationMenu: () => {} },
  session: {
    defaultSession: { protocol: { handle: () => {} }, setDisplayMediaRequestHandler: () => {} },
  },
  nativeImage: {},
  shell: {},
  desktopCapturer: { getSources: async () => [] },
  WebContentsView: class {},
}))
// The shaping chain loads a harfbuzz wasm that cannot resolve in the unit-test env;
// none of the handlers under test measure text.
vi.mock('../src/main/shaped-metrics', () => ({
  shapedMetricsReady: async () => {},
  refineComplexWidths: async () => {},
  complexScriptOf: () => null,
  initShapedMetrics: () => {},
  gtMeasure: (t: string) => ({ width: t.length * 8 }),
  shapedFamily: () => null,
  shapedMeasure: () => ({ width: 0 }),
}))

import { serialize } from 'node:v8'
import {
  addPicture,
  addElement,
  createBlankPptx,
  duplicateSlide,
  openPptx,
  savePptx,
  replacePictureBytes,
  type PictureElement,
  type TextElement,
} from '@genoffice/pptx-engine'
import { registerSlidesIpc } from '../src/main/slides-main'
import {
  buildAllRenderSlides,
  makeMediaResolver,
  rebuildSlide,
  readSlideMedia,
  pruneMediaSources,
  takeSnapshot,
  sessions,
  type Session,
} from '../src/main/session-state'
import { mergeSlideUpdate } from '../src/shared/slide-update'
import type { PasteSlideResult } from '../src/shared/ipc'
import type { PictureRenderNode, RenderSlide } from '@genoffice/pptx-render'

const FIT = 1280
const WC = 1
let session: Session
function call(channel: string, ...args: unknown[]): any {
  return (handlers.get(channel) as (...args: unknown[]) => unknown)({ sender: { id: WC } }, ...args)
}
beforeAll(() => registerSlidesIpc())
beforeEach(async () => {
  await new Promise<void>((resolve) => setImmediate(resolve))
  broadcasts.length = 0
  sessions.clear()
  session = {
    path: '',
    fitWidthPx: FIT,
    undoStack: [],
    redoStack: [],
    opened: await openPptx(await createBlankPptx()),
  }
  sessions.set(WC, session)
})

function addImage(size: number): Uint8Array {
  const bytes = new Uint8Array(size)
  // A valid 1px PNG followed by padding models a large embedded image without a fixture.
  bytes.set(
    Buffer.from(
      'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+aS9sAAAAASUVORK5CYII=',
      'base64',
    ),
  )
  expect(
    addPicture(session.opened, session.opened.deck.slides[0]!, {
      bytes,
      ext: 'png',
      offset: { x: 0, y: 0, cx: 914400, cy: 914400 },
    }),
  ).toBeTruthy()
  expect(duplicateSlide(session.opened, 0)).toBeTruthy()
  return bytes
}

describe('slide media IPC payload', () => {
  it.each(
    ['slides:paste-slide', 'slides:duplicate-slides', 'slides:add-slide'].flatMap((channel) =>
      [1024, 50 * 1024 * 1024].map((size) => [channel, size] as const),
    ),
  )('%s keeps a %i byte image outside the response', (channel, size) => {
    addImage(size)
    call('slides:copy-slides', { slideIndexes: [0] })
    const r = call(channel, { afterIndex: 0, sourceIndex: 0, slideIndexes: [0], fitWidthPx: FIT })
    const payloadBytes = serialize(r).byteLength
    process.stdout.write(`${channel} (${size} byte image): V8 payload ${payloadBytes} bytes\n`)
    expect(JSON.stringify(r).includes('data:image/')).toBe(false)
    expect(payloadBytes).toBeLessThan(32 * 1024)
    expect(r.slides).toHaveLength(1)
    const url = (r.slides[0].nodes[0] as PictureRenderNode).dataUrl!
    expect(url.startsWith('genoffice-slide-media://')).toBe(true)
    expect(readSlideMedia(url)?.body.byteLength).toBe(size)
  })

  it.each([1024, 50 * 1024 * 1024])(
    'inserts and deletes slides with %i byte media without sending the deck',
    async (size) => {
      addImage(size)
      const before = buildAllRenderSlides(session.opened, FIT)
      const added = call('slides:add-blank-slide', {
        sourceIndex: 0,
        before: true,
        fitWidthPx: FIT,
      })
      expect(serialize(added).byteLength).toBeLessThan(32 * 1024)
      let current = mergeSlideUpdate(before, added)
      expect(current.map((s) => s.partPath)).toEqual(added.order)
      expect(current.slice(1)).toEqual(before)
      const deleted = call('slides:delete-slides', { slideIndexes: [0, 2] })
      expect(serialize(deleted).byteLength).toBeLessThan(32 * 1024)
      current = mergeSlideUpdate(current, deleted)
      expect(current.map((s) => s.partPath)).toEqual([before[0]!.partPath])
      expect((call('slides:undo') as RenderSlide[]).map((s) => s.partPath)).toEqual(added.order)
      expect((call('slides:redo') as RenderSlide[]).map((s) => s.partPath)).toEqual(deleted.order)
      const saved = await openPptx(await savePptx(session.opened))
      expect(saved.deck.slides.map((s) => s.path)).toEqual(deleted.order)
      const picture = saved.deck.slides[0]!.elements.find(
        (e) => e.type === 'picture',
      ) as PictureElement
      expect(saved.archive.readBytes(picture.mediaRef!)?.byteLength).toBe(size)
    },
  )

  it.each([1024, 50 * 1024 * 1024])(
    'reorders, hides and deletes without transferring %i byte media',
    async (size) => {
      const bytes = addImage(size)
      duplicateSlide(session.opened, 1)
      let current = buildAllRenderSlides(session.opened, FIT)
      const original = current.map((s) => s.partPath)
      for (const [channel, op] of [
        ['slides:move-slide', { fromIndex: 0, toIndex: 2 }],
        ['slides:move-slides', { slideIndexes: [0, 1], insertAt: 3 }],
        ['slides:set-slides-hidden', { slideIndexes: [0, 1], hidden: true }],
        ['slides:delete-slide', 1],
      ] as const) {
        const result = call(channel, op)
        const payloadBytes = serialize(result).byteLength
        process.stdout.write(`${channel} (${size} byte image): V8 payload ${payloadBytes} bytes\n`)
        expect(payloadBytes).toBeLessThan(32 * 1024)
        current = mergeSlideUpdate(current, result)
        expect(current).toEqual(buildAllRenderSlides(session.opened, FIT))
        const undone = call('slides:undo') as RenderSlide[]
        expect(serialize(undone).byteLength).toBeLessThan(32 * 1024)
        const redone = call('slides:redo') as RenderSlide[]
        expect(redone).toEqual(current)
      }
      expect(current.map((s) => s.partPath)).not.toEqual(original)
      const reopened = await openPptx(await savePptx(session.opened))
      expect(reopened.deck.slides.map((s) => s.path)).toEqual(current.map((s) => s.partPath))
      const picture = reopened.deck.slides
        .flatMap((s) => s.elements)
        .find((e) => e.type === 'picture') as PictureElement
      expect(
        Buffer.compare(
          Buffer.from(reopened.archive.readBytes(picture.mediaRef!)!),
          Buffer.from(bytes),
        ),
      ).toBe(0)
    },
  )

  it.each(['slides:paste-slide', 'slides:duplicate-slides'] as const)(
    'preserves untouched package parts after %s and reopening',
    async (channel) => {
      const bytes = addImage(1024)
      const fixture = await savePptx(session.opened)
      session.opened = await openPptx(fixture)
      const JSZip = (await import('jszip')).default
      const original = await JSZip.loadAsync(fixture)
      call('slides:copy-slides', { slideIndexes: [0] })
      const result = call(channel, {
        afterIndex: 0,
        slideIndexes: [0],
        fitWidthPx: FIT,
      }) as PasteSlideResult
      const saved = await savePptx(session.opened)
      const reopened = await openPptx(saved)
      const output = await JSZip.loadAsync(saved)
      // Insertion changes the presentation manifest and its relationships.
      const changed = new Set([
        '[Content_Types].xml',
        'ppt/presentation.xml',
        'ppt/_rels/presentation.xml.rels',
      ])
      for (const [path, entry] of Object.entries(original.files)) {
        if (entry.dir || changed.has(path)) continue
        expect(output.file(path), path).not.toBeNull()
        expect(await output.file(path)!.async('uint8array'), path).toEqual(
          await entry.async('uint8array'),
        )
      }
      const picture = reopened.deck.slides[result.index]!.elements.find(
        (element) => element.type === 'picture',
      ) as PictureElement
      expect(reopened.archive.readBytes(picture.mediaRef!)).toEqual(bytes)
    },
  )

  it('keeps a paste into a 200-slide deck below 1 MB', () => {
    addImage(1024 * 1024)
    for (let i = 2; i < 200; i++) expect(duplicateSlide(session.opened, 0)).toBeTruthy()
    const before = buildAllRenderSlides(session.opened, FIT)
    call('slides:copy-slides', { slideIndexes: [0] })
    const result = call('slides:paste-slide', { afterIndex: 99, fitWidthPx: FIT })
    const payloadBytes = serialize(result).byteLength
    process.stdout.write(`200-slide paste: V8 payload ${payloadBytes} bytes\n`)
    expect(payloadBytes).toBeLessThan(1024 * 1024)
    expect(mergeSlideUpdate(before, result).map((s) => s.partPath)).toEqual(result.order)
  })

  it.each([1024, 50 * 1024 * 1024])(
    'updates only the background target with %i byte media',
    async (size) => {
      const bytes = addImage(size)
      const before = buildAllRenderSlides(session.opened, FIT)
      const result = await call('slides:edit-background', {
        slideIndex: 0,
        kind: 'solid',
        color: 'FF0000',
        fitWidthPx: FIT,
      })
      expect(serialize(result).byteLength).toBeLessThan(32 * 1024)
      expect(result.slides).toHaveLength(1)
      const updated = mergeSlideUpdate(before, result)
      expect(updated).toEqual(buildAllRenderSlides(session.opened, FIT))
      expect(call('slides:undo')).toEqual(before)
      expect(call('slides:redo')).toEqual(updated)
      const reopened = await openPptx(await savePptx(session.opened))
      const picture = reopened.deck.slides[0]!.elements.find(
        (e) => e.type === 'picture',
      ) as PictureElement
      expect(
        Buffer.compare(
          Buffer.from(reopened.archive.readBytes(picture.mediaRef!)!),
          Buffer.from(bytes),
        ),
      ).toBe(0)
    },
  )

  it.each([1024, 50 * 1024 * 1024])(
    'replaces text without transferring %i byte media',
    async (size) => {
      const bytes = addImage(size)
      addElement(session.opened.deck.slides[0]!, {
        kind: 'textbox',
        offset: { x: 0, y: 0, cx: 914400, cy: 914400 },
        paragraphs: [{ runs: [{ text: 'needle' }] }],
      })
      const before = buildAllRenderSlides(session.opened, FIT)
      const result = call('slides:find-replace', { find: 'needle', replace: 'replacement' })
      expect(result.count).toBe(1)
      expect(result.update.slides).toHaveLength(1)
      expect(serialize(result).byteLength).toBeLessThan(32 * 1024)
      const updated = mergeSlideUpdate(before, result.update)
      expect(updated).toEqual(buildAllRenderSlides(session.opened, FIT))
      expect(call('slides:undo')).toEqual(before)
      expect(call('slides:redo')).toEqual(updated)
      const reopened = await openPptx(await savePptx(session.opened))
      const picture = reopened.deck.slides[0]!.elements.find(
        (e) => e.type === 'picture',
      ) as PictureElement
      expect(
        Buffer.compare(
          Buffer.from(reopened.archive.readBytes(picture.mediaRef!)!),
          Buffer.from(bytes),
        ),
      ).toBe(0)
    },
  )

  it('keeps multiple large images outside the update and the open result', () => {
    const bytes = addImage(8 * 1024 * 1024)
    for (let i = 1; i < 6; i++) {
      addPicture(session.opened, session.opened.deck.slides[0]!, {
        bytes,
        ext: 'png',
        offset: { x: i * 914400, y: 0, cx: 914400, cy: 914400 },
      })
    }
    expect(serialize(buildAllRenderSlides(session.opened, FIT)).byteLength).toBeLessThan(32 * 1024)
    call('slides:copy-slides', { slideIndexes: [0] })
    const result = call('slides:paste-slide', { afterIndex: 0, fitWidthPx: FIT })
    expect(serialize(result).byteLength).toBeLessThan(32 * 1024)
    expect(
      result.slides[0].nodes.filter((n: { type: string }) => n.type === 'picture'),
    ).toHaveLength(6)
  })

  it('broadcasts only new pages to both attached windows and coalesces partial updates', async () => {
    addImage(1024 * 1024)
    sessions.set(2, session)
    const before = buildAllRenderSlides(session.opened, FIT)
    call('slides:copy-slides', { slideIndexes: [0] })
    call('slides:paste-slide', { afterIndex: 0, fitWidthPx: FIT })
    call('slides:duplicate-slides', { slideIndexes: [0], fitWidthPx: FIT })
    const expected = buildAllRenderSlides(session.opened, FIT)
    await new Promise<void>((resolve) => setImmediate(resolve))
    const messages = broadcasts.filter((b) => b.channel === 'slides:deck-changed')
    expect(messages.map((b) => b.id)).toEqual([1, 2])
    for (const message of messages) {
      const update = message.payload as PasteSlideResult
      expect(update.slides).toHaveLength(2)
      expect(serialize(update).byteLength).toBeLessThan(32 * 1024)
      expect(mergeSlideUpdate(before, update)).toEqual(expected)
    }
  })

  it('resyncs when undo shares a broadcast task with an insertion', async () => {
    addImage(1024)
    sessions.set(2, session)
    call('slides:duplicate-slides', { slideIndexes: [0], fitWidthPx: FIT })
    call('slides:undo')
    const expected = buildAllRenderSlides(session.opened, FIT)
    await new Promise<void>((resolve) => setImmediate(resolve))
    const messages = broadcasts.filter((b) => b.channel === 'slides:deck-changed')
    expect(messages).toHaveLength(2)
    const state = messages[0]!.payload as { slides: RenderSlide[]; order?: string[] }
    expect(state.order).toBeUndefined()
    expect(state.slides).toEqual(expected)
  })

  it('duplicates a noncontiguous selection in order after the last selected page', async () => {
    addImage(1024)
    duplicateSlide(session.opened, 1)
    duplicateSlide(session.opened, 2)
    const before = buildAllRenderSlides(session.opened, FIT)
    const result = call('slides:duplicate-slides', { slideIndexes: [2, 0, 2], fitWidthPx: FIT })
    const current = mergeSlideUpdate(before, result)
    expect(result.index).toBe(3)
    expect(result.slides).toHaveLength(2)
    expect(current.map((s) => s.partPath)).toEqual([
      before[0]!.partPath,
      before[1]!.partPath,
      before[2]!.partPath,
      result.slides[0].partPath,
      result.slides[1].partPath,
      before[3]!.partPath,
    ])
    const saved = await openPptx(await savePptx(session.opened))
    expect(saved.deck.slides).toHaveLength(6)
    expect(saved.deck.slides[3]!.elements.some((e) => e.type === 'picture')).toBe(true)
  })

  it('refreshes shifted slide-number fields and reverses the renumber on picture repaste', () => {
    const bytes = addImage(1024)
    const numbered = session.opened.deck.slides[1]!
    const text = addElement(numbered, {
      kind: 'textbox',
      offset: { x: 0, y: 0, cx: 914400, cy: 914400 },
      paragraphs: [{ runs: [{ text: '2', field: 'slidenum' }] }],
    }) as TextElement
    expect(text).toBeTruthy()
    const before = buildAllRenderSlides(session.opened, FIT)
    call('slides:copy-slides', { slideIndexes: [0], pngs: [Buffer.from(bytes).toString('base64')] })
    const pasted = call('slides:paste-slide', { afterIndex: 0, fitWidthPx: FIT })
    expect(pasted.slides.map((s: RenderSlide) => s.partPath)).toContain(numbered.path)
    let current = mergeSlideUpdate(before, pasted)
    expect(current[2]).toEqual(rebuildSlide(session, 2))
    expect(current[2]).not.toEqual(before[1])
    const picture = call('slides:repaste-slide', { fitWidthPx: FIT, mode: 'picture' })
    current = mergeSlideUpdate(current, picture)
    expect(current[1]).toEqual(rebuildSlide(session, 1))
    expect(current[1]).toEqual(before[1])
  })

  it.each(['theme', 'source'] as const)(
    'pastes between two pages in %s mode, undo/redoes and saves image bytes',
    async (mode) => {
      const bytes = addImage(1024 * 1024)
      const before = buildAllRenderSlides(session.opened, FIT)
      call('slides:copy-slides', { slideIndexes: [0] })
      const result = call('slides:paste-slide', {
        afterIndex: 0,
        fitWidthPx: FIT,
        mode,
      }) as PasteSlideResult
      const merged = mergeSlideUpdate(before, result)
      expect(result.index).toBe(1)
      expect(result.count).toBe(1)
      expect(merged).toHaveLength(3)
      expect(merged[0]).toBe(before[0])
      expect(merged[2]).toBe(before[1])
      const picture = merged[1]!.nodes.find((n) => n.type === 'picture') as PictureRenderNode
      expect(
        Buffer.compare(Buffer.from(readSlideMedia(picture.dataUrl!)!.body), Buffer.from(bytes)),
      ).toBe(0)
      const undone = call('slides:undo') as RenderSlide[]
      expect(undone.map((s) => s.partPath)).toEqual(before.map((s) => s.partPath))
      const redone = call('slides:redo') as RenderSlide[]
      expect(redone.map((s) => s.partPath)).toEqual(merged.map((s) => s.partPath))
      expect(serialize(redone).byteLength).toBeLessThan(32 * 1024)
      const saved = await openPptx(await savePptx(session.opened))
      expect(saved.deck.slides.map((s) => s.path)).toEqual(result.order)
      const savedPicture = saved.deck.slides[1]!.elements.find(
        (el) => el.type === 'picture',
      ) as PictureElement
      expect(
        Buffer.compare(
          Buffer.from(saved.archive.readBytes(savedPicture.mediaRef!)!),
          Buffer.from(bytes),
        ),
      ).toBe(0)
    },
  )

  it('repastes from slides to picture and back without leaving a stale inserted page', () => {
    const bytes = addImage(1024)
    const before = buildAllRenderSlides(session.opened, FIT)
    call('slides:copy-slides', { slideIndexes: [0], pngs: [Buffer.from(bytes).toString('base64')] })
    const pasted = call('slides:paste-slide', { afterIndex: 0, fitWidthPx: FIT, mode: 'theme' })
    let current = mergeSlideUpdate(before, pasted)
    const picture = call('slides:repaste-slide', { fitWidthPx: FIT, mode: 'picture' })
    current = mergeSlideUpdate(current, picture)
    expect(picture.count).toBe(0)
    expect(current).toHaveLength(2)
    expect(current[0]!.nodes.filter((n) => n.type === 'picture')).toHaveLength(2)
    const source = call('slides:repaste-slide', { fitWidthPx: FIT, mode: 'source' })
    current = mergeSlideUpdate(current, source)
    expect(current).toHaveLength(3)
    // Restoring the snapshot also removes the picture pasted on the anchor.
    expect(current[0]!.nodes.filter((n) => n.type === 'picture')).toHaveLength(1)
    expect(session.undoStack).toHaveLength(1)
  })

  it('prunes media absent from the archive and history while retaining rollback bytes', () => {
    addImage(1024)
    const slide = session.opened.deck.slides[0]!
    const pic = slide.elements.find((el) => el.type === 'picture') as PictureElement
    const original = makeMediaResolver(session.opened, slide.path)(pic.mediaRef!)!
    pruneMediaSources(session)
    expect(readSlideMedia(original)).not.toBeNull()
    const oldPath = pic.mediaRef!
    session.undoStack.push(takeSnapshot(session))
    const replacement = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 1])
    expect(replacePictureBytes(session.opened, slide, pic.id, replacement, 'png')).toBe(true)
    const updated = makeMediaResolver(session.opened, slide.path)(pic.mediaRef!)!
    session.opened.archive.entries.delete(oldPath)
    pruneMediaSources(session)
    expect(readSlideMedia(original)).not.toBeNull()
    expect(readSlideMedia(updated)).not.toBeNull()
    session.undoStack = []
    pruneMediaSources(session)
    expect(readSlideMedia(original)).toBeNull()
    expect(readSlideMedia(updated)).not.toBeNull()
  })

  it('gives a replaced image a new URL and restores the original URL on undo', () => {
    addImage(1024)
    const slide = session.opened.deck.slides[0]!
    const pic = slide.elements.find((el) => el.type === 'picture') as PictureElement
    const original = makeMediaResolver(session.opened, slide.path)(pic.mediaRef!)!
    call('slides:copy-slides', { slideIndexes: [0] })
    call('slides:paste-slide', { afterIndex: 0, fitWidthPx: FIT })
    const replacement = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 1])
    expect(replacePictureBytes(session.opened, slide, pic.id, replacement, 'png')).toBe(true)
    const updated = makeMediaResolver(session.opened, slide.path)(pic.mediaRef!)!
    expect(updated).not.toBe(original)
    expect(readSlideMedia(updated)?.body.byteLength).toBe(replacement.length)
    call('slides:undo')
    const restored = session.opened.deck.slides[0]!.elements.find(
      (el) => el.type === 'picture',
    ) as PictureElement
    expect(makeMediaResolver(session.opened, slide.path)(restored.mediaRef!)).toBe(original)
    expect(readSlideMedia('genoffice-slide-media://archive/unknown')).toBeNull()
  })
})
