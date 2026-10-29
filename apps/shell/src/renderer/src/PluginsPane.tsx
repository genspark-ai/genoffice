import { useCallback, useEffect, useState } from 'react'
import type { TFunc } from './locale'
import type { InstalledPlugin } from '../../shared/found-skill'

/**
 * The plugins installed into GenOffice. A plugin is a folder that bundles
 * skills, commands and MCP declarations under one name, and one in the Claude
 * Code / ZCode format installs as it is. The pane lists what sits in
 * GenOffice's own plugins folder and flips GenOffice's own enabled flag; every
 * change is reread from the listing. Nothing on this screen runs anything —
 * MCP declarations are shown, never started.
 */

const api = () => window.aiOffice

type Notice = { kind: 'done' | 'removed'; name: string } | { kind: 'failed'; reason: string }

export function PluginsPane({ t }: { t: TFunc }) {
  const [plugins, setPlugins] = useState<InstalledPlugin[] | null>(null)
  const [busy, setBusy] = useState(false)
  const [notice, setNotice] = useState<Notice | null>(null)

  const load = useCallback(async () => {
    // the method itself is checked, not just the bridge: a preload from before
    // this feature was added goes on answering with a partial object
    const found = await api()?.pluginsList?.()
    if (found) setPlugins(found)
  }, [])

  useEffect(() => {
    void load()
  }, [load])

  const doInstall = async (run: (() => Promise<InstalledPlugin | null>) | undefined) => {
    if (busy || !run) return
    setBusy(true)
    try {
      const plugin = await run()
      if (!plugin) return
      setNotice({ kind: 'done', name: plugin.name })
      await load()
    } catch (e) {
      setNotice({ kind: 'failed', reason: e instanceof Error ? e.message : String(e) })
    } finally {
      setBusy(false)
    }
  }

  const doToggle = async (plugin: InstalledPlugin, enabled: boolean) => {
    if (busy) return
    setBusy(true)
    try {
      await api()?.pluginsSetEnabled?.(plugin.name, enabled)
      // the relist is the truth: the row shows what the folder says, not what
      // the click hoped for
      await load()
    } catch (e) {
      setNotice({ kind: 'failed', reason: e instanceof Error ? e.message : String(e) })
    } finally {
      setBusy(false)
    }
  }

  const doRemove = async (plugin: InstalledPlugin) => {
    if (busy || !window.confirm(t('intgPluginsRemoveConfirm', { name: plugin.name }))) return
    setBusy(true)
    try {
      await api()?.pluginsRemove?.(plugin.name)
      setNotice({ kind: 'removed', name: plugin.name })
      await load()
    } catch (e) {
      setNotice({ kind: 'failed', reason: e instanceof Error ? e.message : String(e) })
    } finally {
      setBusy(false)
    }
  }

  // nothing rendered until the listing answers: an empty section that fills in
  // a moment later reads as "none installed", the one thing it must not say
  if (!plugins) return null

  return (
    <>
      <h3 className="set-pane-title">{t('setSecPlugins')}</h3>
      <div className="set-field-desc set-intg-lead">{t('intgPluginsDesc')}</div>

      <div className="set-intg-head">
        <div className="set-intg-actions">
          <button
            className="set-btn"
            disabled={busy}
            onClick={() =>
              void doInstall(() => api()?.pluginsImportZip?.() ?? Promise.resolve(null))
            }
          >
            {t('intgPluginsImportZip')}
          </button>
          <button
            className="set-btn"
            disabled={busy}
            onClick={() =>
              void doInstall(() => api()?.pluginsImportFolder?.() ?? Promise.resolve(null))
            }
          >
            {t('intgPluginsImportFolder')}
          </button>
        </div>
      </div>

      {notice && (
        <div
          className={
            notice.kind === 'failed'
              ? 'set-field-desc set-intg-warn'
              : 'set-field-desc set-intg-notice'
          }
        >
          {notice.kind === 'failed'
            ? t('intgPluginsImportFailed', { reason: notice.reason })
            : notice.kind === 'done'
              ? t('intgPluginsImportDone', { name: notice.name })
              : t('intgPluginsRemoved', { name: notice.name })}
        </div>
      )}

      {plugins.length === 0 && (
        <div className="set-field-desc set-intg-lead">{t('intgPluginsEmpty')}</div>
      )}

      {plugins.length > 0 && (
        <div className="set-intg-list">
          {plugins.map((plugin) => (
            <PluginRow
              key={plugin.name}
              t={t}
              plugin={plugin}
              busy={busy}
              onToggle={(enabled) => void doToggle(plugin, enabled)}
              onRemove={() => void doRemove(plugin)}
            />
          ))}
        </div>
      )}
    </>
  )
}

/** One installed plugin: the name line with version, author and format badge,
 *  the component counts as chips, the enabled switch, and Remove. */
function PluginRow({
  t,
  plugin,
  busy,
  onToggle,
  onRemove,
}: {
  t: TFunc
  plugin: InstalledPlugin
  busy: boolean
  onToggle: (enabled: boolean) => void
  onRemove: () => void
}) {
  const { skills, commands, mcpServers } = plugin.components
  const byline = [plugin.version, plugin.author].filter(Boolean).join(' · ')
  return (
    <div className="set-intg-row set-plug-row">
      <div className="set-field set-field-top">
        <div className="set-field-text">
          <div className="set-field-stack">
            <div className="set-plug-name">
              <span className="set-field-label">{plugin.name}</span>
              {byline && <span className="set-plug-byline">{byline}</span>}
              {plugin.format === 'claude' && (
                <span className="set-plug-badge">{t('intgPluginFormatClaude')}</span>
              )}
            </div>
            {plugin.description && <div className="set-field-desc">{plugin.description}</div>}
            {(skills.length > 0 || commands.length > 0) && (
              <div className="set-plug-meta">
                {skills.length > 0 && (
                  <span className="set-plug-chip">
                    {t('intgPluginSkillsCount', { count: skills.length })}
                  </span>
                )}
                {commands.length > 0 && (
                  <span className="set-plug-chip">
                    {t('intgPluginCommandsCount', { count: commands.length })}
                  </span>
                )}
              </div>
            )}
            {mcpServers.length > 0 && (
              <div className="set-field-desc">{t('intgPluginsMcpNote')}</div>
            )}
          </div>
        </div>
        <div className="set-intg-actions">
          <button
            className="set-switch"
            role="switch"
            aria-checked={plugin.enabled}
            aria-label={plugin.name}
            disabled={busy}
            onClick={() => onToggle(!plugin.enabled)}
          />
          <button className="set-btn" disabled={busy} onClick={onRemove}>
            {t('intgPluginsRemove')}
          </button>
        </div>
      </div>
    </div>
  )
}
