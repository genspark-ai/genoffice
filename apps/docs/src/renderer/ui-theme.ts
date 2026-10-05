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

/**
 * Effective darkness of the document page (#1811): the `<html data-doc-theme>`
 * pin from the shell's document-page-theme setting when explicit, otherwise it
 * follows the UI theme exactly as before the setting existed.
 */
export function docThemeIsDark(): boolean {
  const attr = document.documentElement.getAttribute('data-doc-theme')
  if (attr === 'dark') return true
  if (attr === 'light') return false
  return uiThemeIsDark()
}

/** live docThemeIsDark(): follows document-theme and theme broadcasts and OS appearance flips */
export function useDocThemeIsDark(): boolean {
  const [dark, setDark] = useState(docThemeIsDark)
  useEffect(() => {
    const update = (): void => setDark(docThemeIsDark())
    // main.tsx's listeners (registered at bootstrap) update the attributes
    // first, so reading them in ours is safe; in follow mode the UI theme and
    // OS appearance still drive the result, hence the theme listeners too
    const offDoc = window.desktop?.onDocumentThemeChanged?.(update)
    const offUi = window.desktop?.onThemeChanged?.(update)
    const mq = window.matchMedia?.('(prefers-color-scheme: dark)')
    mq?.addEventListener('change', update)
    return () => {
      offDoc?.()
      offUi?.()
      mq?.removeEventListener('change', update)
    }
  }, [])
  return dark
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
