/**
 * The script editor (issue #815): a small library of scripts, a place to edit
 * one, and Run / Stop with an output pane.
 *
 * This is the thinnest thing that makes scripting usable: scripts are plain
 * JavaScript run in a worker (see scripting/script-worker.ts) against an Apps
 * Script-shaped API over the Univer facade. It deliberately does not try to be a
 * full IDE — no completions, no debugger — because the point of the feature is
 * that an AI can write the script and the user just runs it.
 */
import { useCallback, useEffect, useMemo, useRef, useState } from 'react'
import { showToast } from './toast-bus'
import { useI18n } from './i18n/locale'
import { getUniverAPI } from './scripting/script-api-access'
import { FacadeScriptHost } from './scripting/script-host'
import { createScriptRunner, type ScriptRunner } from './scripting/script-runner'
import {
  SAMPLE_SCRIPT,
  loadScripts,
  newScriptId,
  saveScripts,
  type SavedScript,
} from './scripting/script-storage'

const MAX_LOG_LINES = 300

export function ScriptEditorDialog({ onClose }: { onClose: () => void }): React.ReactElement {
  const { t } = useI18n()
  const [scripts, setScripts] = useState<SavedScript[]>(() => loadScripts())
  const [activeId, setActiveId] = useState<string | null>(null)
  const [logs, setLogs] = useState<string[]>([])
  const [running, setRunning] = useState(false)
  const runnerRef = useRef<ScriptRunner | null>(null)
  const logRef = useRef<HTMLDivElement | null>(null)
  // `t` is not referentially stable and must not sit in hook deps (CLAUDE.md) —
  // callbacks read it through this ref instead.
  const tRef = useRef(t)
  tRef.current = t
  const saveTimer = useRef<ReturnType<typeof setTimeout> | null>(null)
  const pendingSave = useRef<SavedScript[] | null>(null)

  // Start on the newest script the first time the dialog opens.
  useEffect(() => {
    if (activeId === null) setActiveId(scripts[0]?.id ?? null)
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [])

  // Keystrokes only update state; the localStorage write is debounced so typing a
  // 100k-char script does not serialize the whole library on every keypress.
  const persist = useCallback((next: SavedScript[]) => {
    setScripts(next)
    pendingSave.current = next
    if (saveTimer.current) clearTimeout(saveTimer.current)
    saveTimer.current = setTimeout(() => {
      saveTimer.current = null
      const list = pendingSave.current
      pendingSave.current = null
      if (list && !saveScripts(list)) showToast(tRef.current('appScriptsSaveFailed'), 'error')
    }, 500)
  }, [])

  const active = useMemo(() => scripts.find((s) => s.id === activeId) ?? null, [scripts, activeId])

  const patchActive = useCallback(
    (patch: Partial<SavedScript>) => {
      if (!active) return
      persist(
        scripts.map((s) => (s.id === active.id ? { ...s, ...patch, updatedAt: Date.now() } : s)),
      )
    },
    [active, persist, scripts],
  )

  const stop = useCallback(() => {
    runnerRef.current?.stop()
  }, [])

  const run = useCallback(() => {
    if (!active) return
    const api = getUniverAPI()
    if (!api) {
      setLogs((prev) => [...prev, tRef.current('dlgScriptsNoWorkbook')].slice(-MAX_LOG_LINES))
      return
    }
    setRunning(true)
    setLogs([])
    runnerRef.current ??= createScriptRunner(new FacadeScriptHost(api))
    runnerRef.current.run(active.code, {
      onLog: (text) => setLogs((prev) => [...prev, text].slice(-MAX_LOG_LINES)),
      onError: (message) => setLogs((prev) => [...prev, message].slice(-MAX_LOG_LINES)),
      onDone: () => setRunning(false),
    })
  }, [active])

  // Kill any running script when the dialog goes away, and flush a pending save so
  // quick edits are not lost to the debounce timer.
  useEffect(
    () => () => {
      runnerRef.current?.dispose()
      if (saveTimer.current) clearTimeout(saveTimer.current)
      if (pendingSave.current) saveScripts(pendingSave.current)
    },
    [],
  )

  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') onClose()
    }
    window.addEventListener('keydown', onKey)
    return () => window.removeEventListener('keydown', onKey)
  }, [onClose])

  useEffect(() => {
    logRef.current?.scrollTo?.({ top: logRef.current.scrollHeight })
  }, [logs])

  const addScript = () => {
    const script: SavedScript = {
      id: newScriptId(),
      name: t('dlgScriptsUntitled'),
      code: SAMPLE_SCRIPT,
      updatedAt: Date.now(),
    }
    persist([script, ...scripts])
    setActiveId(script.id)
  }

  const deleteActive = () => {
    if (!active) return
    const next = scripts.filter((s) => s.id !== active.id)
    persist(next)
    setActiveId(next[0]?.id ?? null)
  }

  return (
    <div className="dialog-backdrop" onMouseDown={onClose}>
      <div
        className="script-editor-dialog"
        role="dialog"
        aria-modal="true"
        aria-label={t('dlgScriptsTitle')}
        onMouseDown={(e) => e.stopPropagation()}
      >
        <div className="script-editor-head">
          <span className="script-editor-title">{t('dlgScriptsTitle')}</span>
          <span className="dialog-note">{t('dlgScriptsHelp')}</span>
        </div>
        <div className="dialog-body script-editor-body">
          <aside className="script-list">
            <button type="button" className="secondary" onClick={addScript}>
              {t('dlgScriptsNew')}
            </button>
            <div className="script-list-items">
              {scripts.length === 0 ? (
                <p className="dialog-note">{t('dlgScriptsEmpty')}</p>
              ) : (
                scripts.map((script) => (
                  <button
                    key={script.id}
                    type="button"
                    className={`script-list-item${script.id === activeId ? ' active' : ''}`}
                    onClick={() => setActiveId(script.id)}
                  >
                    {script.name}
                  </button>
                ))
              )}
            </div>
          </aside>
          <section className="script-main">
            <input
              className="script-name-input"
              value={active?.name ?? ''}
              placeholder={t('dlgScriptsUntitled')}
              disabled={!active}
              onChange={(e) => patchActive({ name: e.target.value })}
            />
            <textarea
              className="script-code"
              spellCheck={false}
              value={active?.code ?? ''}
              disabled={!active}
              placeholder={SAMPLE_SCRIPT}
              onChange={(e) => patchActive({ code: e.target.value })}
            />
            <div className="script-log-head">{t('dlgScriptsOutput')}</div>
            <div className="script-log" ref={logRef}>
              {logs.length === 0 ? (
                <span className="dialog-note">—</span>
              ) : (
                logs.map((line, i) => (
                  <div key={i} className="script-log-line">
                    {line}
                  </div>
                ))
              )}
            </div>
          </section>
        </div>
        <div className="dialog-actions">
          <button type="button" className="secondary" onClick={deleteActive} disabled={!active}>
            {t('dlgScriptsDelete')}
          </button>
          <span className="dialog-actions-spacer" />
          {running ? (
            <button type="button" onClick={stop}>
              {t('dlgScriptsStop')}
            </button>
          ) : (
            <button type="button" onClick={run} disabled={!active}>
              {t('dlgScriptsRun')}
            </button>
          )}
          <button type="button" className="secondary" onClick={onClose}>
            {t('dlgScriptsClose')}
          </button>
        </div>
      </div>
    </div>
  )
}
