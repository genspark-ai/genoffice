import { useEffect, useState } from 'react'

/**
 * Effective darkness of the UI theme: `<html data-theme>` (set by main.tsx from
 * the shell's light/dark/system setting), OS appearance in system mode.
 */
export function uiThemeIsDark(): boolean {
  const attr = document.documentElement.getAttribute('data-theme')
  if (attr === 'dark') return true
  if (attr === 'light') return false
  return window.matchMedia?.('(prefers-color-scheme: dark)').matches ?? false
}

/** live uiThemeIsDark(): follows theme broadcasts and OS appearance flips in system mode */
export function useUiThemeIsDark(): boolean {
  const [dark, setDark] = useState(uiThemeIsDark)
  useEffect(() => {
    const update = (): void => setDark(uiThemeIsDark())
    // main.tsx's listener (registered at bootstrap) updates data-theme first,
    // so reading the attribute in ours is safe
    const off = window.desktop?.onThemeChanged?.(update)
    const mq = window.matchMedia?.('(prefers-color-scheme: dark)')
    mq?.addEventListener('change', update)
    return () => {
      off?.()
      mq?.removeEventListener('change', update)
    }
  }, [])
  return dark
}
