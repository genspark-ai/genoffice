import { useEffect, useRef } from 'react'
import { createStudio, type RhwpEditor } from '@rhwp/editor'
import type { SaveMode } from '../shared/ipc'
import { asBytes } from './as-bytes'

function fileName(path: string | null): string {
  if (!path) return 'untitled.hwp'
  const parts = path.split(/[\\/]/)
  return parts[parts.length - 1] ?? path
}

async function exportPayload(studio: RhwpEditor): Promise<{
  hwp: Uint8Array
  hwpx: Uint8Array
  hml?: Uint8Array
}> {
  const [hwp, hwpx] = await Promise.all([studio.exportHwp(), studio.exportHwpx()])
  try {
    return { hwp, hwpx, hml: await studio.exportHml() }
  } catch {
    return { hwp, hwpx }
  }
}

export function HwpStudio({
  path,
  onPath,
  onDirty,
  onError,
}: {
  path: string | null
  onPath: (path: string) => void
  onDirty: (dirty: boolean) => void
  onError: (message: string) => void
}) {
  const hostRef = useRef<HTMLDivElement>(null)
  const studioRef = useRef<RhwpEditor | null>(null)
  const dirtyRef = useRef(false)
  const savingRef = useRef(false)
  const onPathRef = useRef(onPath)
  onPathRef.current = onPath
  const onDirtyRef = useRef(onDirty)
  onDirtyRef.current = onDirty
  const onErrorRef = useRef(onError)
  onErrorRef.current = onError

  useEffect(() => {
    const host = hostRef.current
    if (!host) return
    let cancelled = false
    let poll: ReturnType<typeof setInterval> | undefined
    const studioUrl = new URL('rhwp/?chrome=embed', window.location.href).href

    const markDirty = (dirty: boolean) => {
      if (dirtyRef.current === dirty) return
      dirtyRef.current = dirty
      onDirtyRef.current(dirty)
    }

    const doSave = async (mode: SaveMode): Promise<boolean> => {
      const studio = studioRef.current
      if (!studio || savingRef.current) return false
      savingRef.current = true
      try {
        const payload = await exportPayload(studio)
        const result = await window.hwpApi.save({ mode, ...payload })
        if (result.ok && 'path' in result) {
          onPathRef.current(result.path)
          await studio.notifySaved(fileName(result.path)).catch(() => undefined)
          markDirty(false)
          return true
        }
        return Boolean(result.ok && 'canceled' in result)
      } catch (err) {
        console.error('[hwp] save failed:', err)
        return false
      } finally {
        savingRef.current = false
      }
    }

    void (async () => {
      try {
        const studio = await createStudio(host, { studioUrl, renderer: 'canvas2d' })
        if (cancelled) {
          studio.destroy()
          return
        }
        studioRef.current = studio
        const openPath = await window.hwpApi.consumePending()
        if (cancelled) return
        if (openPath) {
          const bytes = asBytes(await window.hwpApi.readFile(openPath))
          await studio.loadFile(bytes, fileName(openPath), {
            skipUnsavedGuard: true,
            suppressDialogs: true,
          })
          onPathRef.current(openPath)
        }
        markDirty(false)
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
        while (savingRef.current) {
          await new Promise((resolve) => setTimeout(resolve, 50))
        }
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
      offSave()
      offClose()
      window.removeEventListener('keydown', onKeyDown, true)
      studioRef.current?.destroy()
      studioRef.current = null
      host.replaceChildren()
    }
  }, [])

  return <div ref={hostRef} className="hwp-studio" aria-label={fileName(path)} />
}
