import { createContext, useCallback, useContext, useEffect, useMemo, useRef, useState } from 'react'
import type { ReactNode } from 'react'
import { createPortal } from 'react-dom'

export interface ToastOptions {
  message: string
  /** label of the inline action button (e.g. Undo) */
  actionLabel?: string
  onAction?: () => void
  /** auto-dismiss delay; defaults to 5s */
  durationMs?: number
  /** fired when the toast goes away without its action being used */
  onExpire?: () => void
}

interface ToastItem extends ToastOptions {
  id: number
}

interface ToastApi {
  show: (options: ToastOptions) => number
  dismiss: (id: number) => void
}

const ToastContext = createContext<ToastApi | null>(null)

export function ToastProvider({ children }: { children: ReactNode }) {
  const [items, setItems] = useState<ToastItem[]>([])
  const nextId = useRef(1)
  const timers = useRef(new Map<number, ReturnType<typeof setTimeout>>())

  const dismiss = useCallback((id: number) => {
    const timer = timers.current.get(id)
    if (timer) clearTimeout(timer)
    timers.current.delete(id)
    setItems((prev) => prev.filter((item) => item.id !== id))
  }, [])

  const show = useCallback(
    (options: ToastOptions) => {
      const id = nextId.current++
      setItems((prev) => [...prev.slice(-2), { ...options, id }])
      timers.current.set(
        id,
        setTimeout(() => {
          options.onExpire?.()
          dismiss(id)
        }, options.durationMs ?? 5000),
      )
      return id
    },
    [dismiss],
  )

  useEffect(() => {
    const live = timers.current
    return () => live.forEach((timer) => clearTimeout(timer))
  }, [])

  const api = useMemo(() => ({ show, dismiss }), [show, dismiss])

  return (
    <ToastContext.Provider value={api}>
      {children}
      {typeof document !== 'undefined' &&
        createPortal(
          <div className="toast-region" role="status" aria-live="polite">
            {items.map((item) => (
              <div className="toast" key={item.id}>
                <span className="toast-text">{item.message}</span>
                {item.actionLabel && (
                  <button
                    className="toast-action"
                    onClick={() => {
                      item.onAction?.()
                      dismiss(item.id)
                    }}
                  >
                    {item.actionLabel}
                  </button>
                )}
              </div>
            ))}
          </div>,
          document.body,
        )}
    </ToastContext.Provider>
  )
}

/** no-op API outside a provider, so components stay testable in isolation */
const NOOP: ToastApi = { show: () => 0, dismiss: () => undefined }

export function useToast(): ToastApi {
  return useContext(ToastContext) ?? NOOP
}
