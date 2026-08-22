import { useEffect, useState } from 'react'

/**
 * Track the browser's connectivity as a reactive boolean. Fires on the
 * `online`/`offline` window events; used by the AI panels to auto-resume an
 * interrupted run when the network comes back.
 */
export function useOnlineStatus(): boolean {
  const [online, setOnline] = useState(() => navigator.onLine)
  useEffect(() => {
    const up = () => setOnline(true)
    const down = () => setOnline(false)
    window.addEventListener('online', up)
    window.addEventListener('offline', down)
    return () => {
      window.removeEventListener('online', up)
      window.removeEventListener('offline', down)
    }
  }, [])
  return online
}