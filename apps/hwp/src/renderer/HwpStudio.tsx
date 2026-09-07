import { useEffect, useRef } from 'react'
import { createStudio } from '@rhwp/editor'
import { HML_UNAVAILABLE } from '../shared/formats'
import type { SaveMode } from '../shared/ipc'
import { asBytes } from '../shared/as-bytes'
import { exportStudioPayload, hmlUnavailableMessage } from './export-payload'
import { createStudioFacade, fileNameOf, type HangulStudioFacade } from './studio-text'

function notifyStudioResize(host: HTMLElement): void {
  host.querySelector('iframe')?.contentWindow?.dispatchEvent(new Event('resize'))
}

export function HwpStudio({
  path,
  onPath,
  onDirty,
  onError,
  onSaveError,
  onReady,
}: {
  path: string | null
  onPath: (path: string) => void
  onDirty: (dirty: boolean) => void
  onError: (message: string) => void
  onSaveError: (message: string | null) => void
  onReady?: (facade: HangulStudioFacade | null) => void
}) {
  const hostRef = useRef<HTMLDivElement>(null)
  const studioRef = useRef<Awaited<ReturnType<typeof createStudio>> | null>(null)
  const dirtyRef = useRef(false)
  const onPathRef = useRef(onPath)
  onPathRef.current = onPath
  const onDirtyRef = useRef(onDirty)
  onDirtyRef.current = onDirty
  const onErrorRef = useRef(onError)
  onErrorRef.current = onError
  const onSaveErrorRef = useRef(onSaveError)
  onSaveErrorRef.current = onSaveError
  const onReadyRef = useRef(onReady)
  onReadyRef.current = onReady

  useEffect(() => {
    const host = hostRef.current
    if (!host) return
    let cancelled = false
    let poll: ReturnType<typeof setInterval> | undefined
    let resizeTimer: ReturnType<typeof setTimeout> | undefined
    let resizeObserver: ResizeObserver | undefined
    let inflight: Promise<boolean> | null = null
    const studioUrl = new URL('rhwp/?chrome=embed', window.location.href).href

    const markDirty = (dirty: boolean) => {
      if (dirtyRef.current === dirty) return
      dirtyRef.current = dirty
      onDirtyRef.current(dirty)
    }

    const runSave = async (mode: SaveMode): Promise<boolean> => {
      const studio = studioRef.current
      if (!studio) return false
      try {
        const payload = await exportStudioPayload(studio)
        const result = await window.hwpApi.save({
          mode,
          hwp: payload.hwp,
          hwpx: payload.hwpx,
          ...(payload.hml ? { hml: payload.hml } : {}),
        })
        if (result.ok && 'path' in result) {
          onPathRef.current(result.path)
          await studio.notifySaved(fileNameOf(result.path)).catch(() => undefined)
          markDirty(false)
          onSaveErrorRef.current(null)
          return true
        }
        if (result.ok && 'canceled' in result) return true
        if (!result.ok) {
          onSaveErrorRef.current(
            result.error === HML_UNAVAILABLE
              ? hmlUnavailableMessage(payload.hmlBlockers)
              : result.error,
          )
        }
        return false
      } catch (err) {
        const message = err instanceof Error ? err.message : String(err)
        console.error('[hwp] save failed:', err)
        onSaveErrorRef.current(message)
        return false
      }
    }

    const doSave = async (mode: SaveMode): Promise<boolean> => {
      if (inflight) {
        const ok = await inflight
        if (mode === 'save') return ok
      }
      inflight = runSave(mode).finally(() => {
        inflight = null
      })
      return inflight
    }

    void (async () => {
      try {
        const studio = await createStudio(host, { studioUrl, renderer: 'canvas2d' })
        await studio.plugins.load('hwpctrl').catch(() => undefined)
        if (cancelled) {
          studio.destroy()
          return
        }
        studioRef.current = studio
        const openPath = await window.hwpApi.consumePending()
        if (cancelled) return
        if (openPath) {
          const bytes = asBytes(await window.hwpApi.readFile(openPath))
          await studio.loadFile(bytes, fileNameOf(openPath), {
            skipUnsavedGuard: true,
            suppressDialogs: true,
          })
          onPathRef.current(openPath)
        } else {
          // embed boot skips createNewDocument(); the host owns File → New
          const created = await studio.commands.execute('file:new-doc', {}, { allowDialog: false })
          if (!created.ok) {
            throw new Error(created.message ?? created.reason)
          }
        }
        markDirty(false)
        onReadyRef.current?.(createStudioFacade(studio))
        notifyStudioResize(host)
        resizeObserver = new ResizeObserver(() => {
          clearTimeout(resizeTimer)
          resizeTimer = setTimeout(() => notifyStudioResize(host), 200)
        })
        resizeObserver.observe(host)
        poll = setInterval(() => {
          void studio.getDocumentState().then(
            (state) => markDirty(state.dirty),
            () => undefined,
          )
        }, 750)
      } catch (err) {
        if (!cancelled) {
          console.error('[hwp] studio failed:', err)
          onErrorRef.current(err instanceof Error ? err.message : String(err))
        }
      }
    })()

    const offSave = window.hwpApi.onSaveRequest(
      (mode) => void doSave(mode).then((ok) => window.hwpApi.sendSaveRequestAck(ok)),
    )
    const offClose = window.hwpApi.onCloseSaveRequest(() => {
      void (async () => {
        if (inflight) await inflight
        if (!dirtyRef.current) {
          window.hwpApi.sendCloseSaveResult(true)
          return
        }
        window.hwpApi.sendCloseSaveResult(await doSave('save'))
      })()
    })
    const onKeyDown = (event: KeyboardEvent) => {
      if (!(event.metaKey || event.ctrlKey) || event.altKey) return
      if (event.key.toLowerCase() !== 's') return
      event.preventDefault()
      void doSave(event.shiftKey ? 'saveAs' : 'save')
    }
    window.addEventListener('keydown', onKeyDown, true)

    return () => {
      cancelled = true
      if (poll) clearInterval(poll)
      clearTimeout(resizeTimer)
      resizeObserver?.disconnect()
      onReadyRef.current?.(null)
      offSave()
      offClose()
      window.removeEventListener('keydown', onKeyDown, true)
      studioRef.current?.destroy()
      studioRef.current = null
      host.replaceChildren()
    }
  }, [])

  return <div ref={hostRef} className="hwp-studio" aria-label={fileNameOf(path)} />
}
