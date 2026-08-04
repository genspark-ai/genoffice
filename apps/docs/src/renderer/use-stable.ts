import { useRef } from 'react'

export function shallowEqual(a: object, b: object): boolean {
  if (a === b) return true
  const ra = a as Record<string, unknown>
  const rb = b as Record<string, unknown>
  const keys = Object.keys(ra)
  if (keys.length !== Object.keys(rb).length) return false
  for (const k of keys) {
    if (!Object.is(ra[k], rb[k])) return false
  }
  return true
}

/** Reuse the previous object reference while all fields stay shallow-equal (keeps React.memo props stable). */
export function useShallowStable<T extends object>(next: T): T {
  const ref = useRef(next)
  if (ref.current !== next && !shallowEqual(ref.current, next)) ref.current = next
  return ref.current
}

type AnyFn = (...args: never[]) => unknown

/**
 * Stable identities for per-render closures: the returned functions never change,
 * but always dispatch into the latest render's handlers (same pattern as fileCtxRef).
 * The key set must be identical on every call.
 */
export function useStableCallbacks<T extends Record<string, AnyFn>>(handlers: T): T {
  const latest = useRef(handlers)
  latest.current = handlers
  const stable = useRef<T | null>(null)
  if (!stable.current) {
    const out: Record<string, AnyFn> = {}
    for (const key of Object.keys(handlers)) {
      out[key] = (...args: never[]) => latest.current[key](...args)
    }
    stable.current = out as T
  }
  return stable.current
}
