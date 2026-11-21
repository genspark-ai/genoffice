import { useCallback, useEffect, useRef, useState } from 'react'

function readFlag(key: string, fallback: boolean): boolean {
  try {
    const raw = globalThis.localStorage?.getItem(key)
    return raw === '1' ? true : raw === '0' ? false : fallback
  } catch {
    return fallback
  }
}

/**
 * An on/off view preference that outlives the document, like Word's View ▸
 * Ruler: kept in this app's local storage, read when a window opens and
 * followed live when another Docs window changes it.
 */
export function useStoredFlag(
  key: string,
  fallback = false,
): [boolean, (on: boolean | ((prev: boolean) => boolean)) => void] {
  const [on, setOn] = useState(() => readFlag(key, fallback))
  const current = useRef(on)
  current.current = on
  const set = useCallback(
    (value: boolean | ((prev: boolean) => boolean)) => {
      const next = typeof value === 'function' ? value(current.current) : value
      current.current = next
      setOn(next)
      try {
        globalThis.localStorage?.setItem(key, next ? '1' : '0')
      } catch {
        /* storage unavailable: the choice lasts for this window only */
      }
    },
    [key],
  )
  useEffect(() => {
    const onStorage = (e: StorageEvent) => {
      if (e.key === key) setOn(readFlag(key, fallback))
    }
    window.addEventListener('storage', onStorage)
    return () => window.removeEventListener('storage', onStorage)
  }, [key, fallback])
  return [on, set]
}
