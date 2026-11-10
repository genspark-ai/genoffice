import { existsSync } from 'node:fs'
import { extname, isAbsolute } from 'node:path'
import { webContents } from 'electron'
import { elementDurableId, slideDurableId, type SlideElement } from '@genoffice/pptx-engine'
import { applySessionTxn, saveSessionDeckTo } from '../../../../slides/src/main/slides-main'
import { attachedIds, sessions, type Session } from '../../../../slides/src/main/session-state'
import type { SlidesControl, SlidesTxnRequest } from './tools/slides-tools'

/**
 * Shell-main half of the MCP → slides bridge.
 *
 * Unlike docs, a slides editing session lives in the main process (keyed by the
 * tab's webContents id), so there is no renderer protocol here: the control
 * opens a blank slides tab, waits for the session the renderer's boot pull
 * creates (`slides:new-blank`), then drives it through the same functions the
 * app's own `slides:apply-txn` and save pipeline use. Every transaction lands
 * in the session's undo history and journal, and the tab re-renders live.
 */

const READY_TIMEOUT_MS = 20_000
const EMU_PER_PX_96 = 9525
/** per-text cap in read_deck, mirroring the CLI's describeDeck clipping */
const MAX_TEXT_CHARS = 300
/** table rows echoed per element (the CLI caps its own table output the same way) */
const MAX_TABLE_ROWS = 3

function clip(text: string, max = MAX_TEXT_CHARS): string {
  return text.length > max ? `${text.slice(0, max)}…` : text
}

/** the session a tab's renderer creates on boot; the renderer applies the blank deck right after */
async function waitSession(wcId: number): Promise<Session> {
  const deadline = Date.now() + READY_TIMEOUT_MS
  for (;;) {
    const session = sessions.get(wcId)
    if (session) return session
    const wc = webContents.fromId(wcId)
    if (!wc || wc.isDestroyed()) {
      throw new Error('the presentation tab was closed before it became ready')
    }
    if (Date.now() > deadline) {
      throw new Error(`the presentation did not become ready within ${READY_TIMEOUT_MS}ms`)
    }
    await new Promise((resolve) => setTimeout(resolve, 100))
  }
}

function requireSession(wcId: number): Session {
  const session = sessions.get(wcId)
  if (!session) {
    throw new Error('the deck session is gone — its tab was closed; call create_session again')
  }
  return session
}

/**
 * Compact model readout for the agent: ids, geometry (EMU) and text per element.
 *
 * Group children carry the group's child coordinate system, not document-space
 * EMU, so their box is reported as `local` (with the parent's own box alongside)
 * instead of being passed off as an absolute position — an agent placing a
 * follow-up setTransform needs to know which frame it is addressing.
 */
function elementInfo(el: SlideElement): Record<string, unknown> {
  const t = el.transform?.offset ?? { x: 0, y: 0, cx: 0, cy: 0 }
  const info: Record<string, unknown> = {
    id: elementDurableId(el) ?? el.id,
    type: el.type,
    x: t.x,
    y: t.y,
    w: t.cx,
    h: t.cy,
  }
  if (el.type === 'text' || el.type === 'shape') {
    const paragraphs = el.text?.paragraphs ?? []
    if (paragraphs.length) {
      const full = paragraphs.map((p) => p.runs.map((r) => r.text).join('')).join('\n')
      info.text = clip(full)
      if (full.length > MAX_TEXT_CHARS) info.textTruncated = true
    }
  } else if (el.type === 'table') {
    const rows = el.rows
    const totalRows = rows.length
    const body = rows
      .slice(0, MAX_TABLE_ROWS)
      .map((row) =>
        row
          .map((cell) =>
            (cell.text?.paragraphs ?? []).map((p) => p.runs.map((r) => r.text).join('')).join(''),
          )
          .join(' | '),
      )
      .join('\n')
    info.text = clip(body)
    if (totalRows > MAX_TABLE_ROWS) {
      info.textTruncated = true
      info.rows = totalRows
    }
  } else if (el.type === 'group') {
    // children are in the group's local frame: say so rather than implying
    // document-space coordinates
    info.children = el.children.map((child) => ({
      ...elementInfo(child),
      coordinates: 'local-to-group',
    }))
  }
  return info
}

function readDeckModel(session: Session): Record<string, unknown> {
  const deck = session.opened.deck
  return {
    emuPerPx: EMU_PER_PX_96,
    note: 'x/y/w/h are EMU in slide space; elements inside a group use the group-local frame (marked coordinates:"local-to-group").',
    slideSize: {
      widthPx: Math.round(deck.size.cx / EMU_PER_PX_96),
      heightPx: Math.round(deck.size.cy / EMU_PER_PX_96),
    },
    slides: deck.slides.map((slide, index) => ({
      index,
      id: slideDurableId(slide),
      elements: slide.elements.map(elementInfo),
    })),
  }
}

export interface SlidesBridgeDeps {
  /** open a fresh blank slides tab; returns its webContents id */
  openBlankTab: () => number
  /** close a blank tab whose session never became ready, so a failed create_session leaves no orphan */
  abandonBlankTab?: (wcId: number) => void
}

export function createSlidesControl(deps: SlidesBridgeDeps): SlidesControl {
  return {
    openBlankTab: async () => {
      const wcId = await deps.openBlankTab()
      try {
        await waitSession(wcId)
      } catch (error) {
        deps.abandonBlankTab?.(wcId)
        throw error
      }
      return wcId
    },
    runTxn: async (wcId: number, req: SlidesTxnRequest) => {
      const session = requireSession(wcId)
      const result = applySessionTxn(session, req as Parameters<typeof applySessionTxn>[1])
      if (!result) throw new Error('the deck session is gone — call create_session again')
      // The ordinary `slides:apply-txn` path has an originating renderer that
      // applies its own IPC return value; MCP runs in the main process, so the
      // tab would never hear about the change — scheduleDeckBroadcast no-ops
      // for single-window sessions. Push the fresh render to the tab ourselves
      // (multi-window sessions already get the scheduled broadcast).
      if (result.applied && result.slides) {
        const ids = attachedIds(session)
        if (ids.length < 2) {
          const payload = {
            slides: result.slides,
            size: { cx: session.opened.deck.size.cx, cy: session.opened.deck.size.cy },
          }
          for (const id of ids) webContents.fromId(id)?.send('slides:deck-changed', payload)
        }
      }
      // The render tree was only needed for the canvas push above: it carries
      // base64 dataUrls for pictures and fills, so echoing it in the tool result
      // would put the whole deck (and its images) into the agent's context on
      // every op. The agent sees the transaction outcome; read_deck is the model
      // readout.
      const { slides: _renderTree, ...summary } = result
      return summary
    },
    readDeck: async (wcId: number) => readDeckModel(requireSession(wcId)),
    saveDeck: async (wcId: number, filePath: string, overwrite: boolean) => {
      if (!isAbsolute(filePath)) throw new Error('path must be absolute')
      const ext = extname(filePath).toLowerCase()
      const targetPath = ext === '' ? `${filePath}.pptx` : ext === '.pptx' ? filePath : null
      if (!targetPath) throw new Error('path must point to a .pptx file')
      if (existsSync(targetPath) && !overwrite) {
        throw new Error(`file already exists: ${targetPath} (pass overwrite:true to replace it)`)
      }
      const session = requireSession(wcId)
      await saveSessionDeckTo(session, targetPath)
      return { path: targetPath }
    },
  }
}
