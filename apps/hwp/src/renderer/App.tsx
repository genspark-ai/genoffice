import { useEffect, useState } from 'react'

function fileName(path: string | null): string {
  if (!path) return ''
  const parts = path.split(/[\\/]/)
  return parts[parts.length - 1] ?? path
}

export default function App() {
  const [path, setPath] = useState<string | null>(null)

  useEffect(() => {
    let cancelled = false
    void window.hwpApi.consumePending().then((next) => {
      if (!cancelled) setPath(next)
    })
    const offRename = window.hwpApi.onFileRenamed((next) => setPath(next))
    return () => {
      cancelled = true
      offRename()
    }
  }, [])

  const name = fileName(path)
  return (
    <div className="hwp-shell">
      <main className="hwp-page">
        <h1 className="hwp-title">{name || 'Hangul'}</h1>
        {path ? <p className="hwp-path">{path}</p> : null}
        <p className="hwp-hint">Hangul document tab</p>
      </main>
    </div>
  )
}
