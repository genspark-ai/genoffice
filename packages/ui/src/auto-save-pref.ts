import { useCallback, useEffect, useState } from 'react'

/** the shell-wide AutoSave default; `updatedAt` is 0 until the user has ever set it */
export interface AutoSaveDefault {
  on: boolean
  updatedAt: number
}

export interface AutoSaveDefaultApi {
  getAutoSaveDefault(): Promise<AutoSaveDefault>
  onAutoSaveDefaultChanged(handler: (value: AutoSaveDefault) => void): () => void
}

export const NO_AUTO_SAVE_DEFAULT: AutoSaveDefault = { on: false, updatedAt: 0 }

/**
 * A per-window override only counts while its `base` still equals the global
 * `updatedAt`: flipping the shell setting invalidates every override at once,
 * even for editors that were not running at the time.
 */
export function resolveAutoSave(global: AutoSaveDefault, storedRaw: string | null): boolean {
  if (storedRaw === '1' || storedRaw === '0') {
    return global.updatedAt === 0 ? storedRaw === '1' : global.on
  }
  if (storedRaw) {
    try {
      const parsed: unknown = JSON.parse(storedRaw)
      if (
        parsed &&
        typeof parsed === 'object' &&
        typeof (parsed as { on?: unknown }).on === 'boolean' &&
        (parsed as { base?: unknown }).base === global.updatedAt
      ) {
        return (parsed as { on: boolean }).on
      }
    } catch {
      // corrupt value: fall through to the global default
    }
  }
  return global.on
}

export function encodeAutoSaveOverride(on: boolean, global: AutoSaveDefault): string {
  return JSON.stringify({ on, base: global.updatedAt })
}

export function isAutoSaveDefault(value: unknown): value is AutoSaveDefault {
  return (
    !!value &&
    typeof value === 'object' &&
    typeof (value as { on?: unknown }).on === 'boolean' &&
    typeof (value as { updatedAt?: unknown }).updatedAt === 'number'
  )
}

/**
 * AutoSave pill state for an editor window: the shell default, overridable
 * per window under `storageKey`. Works without the shell (standalone dev) by
 * falling back to the legacy '1'/'0' value.
 */
export function useAutoSavePref(
  storageKey: string,
  api: Partial<AutoSaveDefaultApi> | undefined,
): [boolean, (on: boolean) => void] {
  const [global, setGlobal] = useState<AutoSaveDefault>(NO_AUTO_SAVE_DEFAULT)
  const [autoSave, setAutoSaveState] = useState(() =>
    resolveAutoSave(NO_AUTO_SAVE_DEFAULT, localStorage.getItem(storageKey)),
  )

  useEffect(() => {
    let alive = true
    const apply = (value: AutoSaveDefault) => {
      if (!alive) return
      setGlobal(value)
      setAutoSaveState(resolveAutoSave(value, localStorage.getItem(storageKey)))
    }
    api
      ?.getAutoSaveDefault?.()
      .then((value) => {
        if (isAutoSaveDefault(value)) apply(value)
      })
      .catch(() => {})
    const off = api?.onAutoSaveDefaultChanged?.((value) => {
      if (isAutoSaveDefault(value)) apply(value)
    })
    return () => {
      alive = false
      off?.()
    }
  }, [api, storageKey])

  const setAutoSave = useCallback(
    (on: boolean) => {
      localStorage.setItem(storageKey, encodeAutoSaveOverride(on, global))
      setAutoSaveState(on)
    },
    [global, storageKey],
  )

  return [autoSave, setAutoSave]
}
