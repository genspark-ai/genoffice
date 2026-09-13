import { useCallback, useEffect, useState } from 'react'
import type { TFunc } from './locale'
import type {
  AgentId,
  AgentTarget,
  IntegrationsStatus,
  SkillInstallState,
} from '../../shared/integrations-api'

// ── Settings → Integrations ─────────────────────────────────
// Installs the bundled `genoffice` skill into the coding agents found on this
// machine. Every write starts with a click and shows the absolute path first,
// because the app does not see the shell's CODEX_HOME-style overrides and the
// user has to be able to spot a wrong target.

type AgentRow = AgentTarget & { state: SkillInstallState }

/** what the confirm block under a row is about */
interface Pending {
  kind: 'install' | 'uninstall'
  path: string
  /** extra warning line (overwriting edits, replacing a foreign SKILL.md) */
  note?: string
  target: { agentId: AgentId } | { dir: string }
  /** row the block renders under; the custom-folder install has none */
  agentId?: AgentId
}

export const NPX_INSTALL_COMMAND = 'npx skills add genspark-ai/genoffice'

/** some detected assistant holds an older copy of the skill than the bundled one */
export const skillUpdateDue = (s: IntegrationsStatus): boolean =>
  s.agents.some((a) => a.state.older === true)

const EXAMPLE_KEYS = ['intgExample1', 'intgExample2', 'intgExample3'] as const

const api = () => window.aiOfficeIntegrations

export function IntegrationsPane({
  t,
  onStatus,
}: {
  t: TFunc
  onStatus?: (s: IntegrationsStatus) => void
}) {
  const [status, setStatus] = useState<IntegrationsStatus | null>(null)
  const [pending, setPending] = useState<Pending | null>(null)
  const [busy, setBusy] = useState(false)
  /** one-line outcome shown until the next action: installed hint, saved zip path, copied */
  const [notice, setNotice] = useState<{ agentId?: AgentId; text: string } | null>(null)
  const [copied, setCopied] = useState<string | null>(null)

  const refresh = useCallback(async () => {
    const s = await api()?.status()
    if (!s) return
    setStatus(s)
    onStatus?.(s)
  }, [onStatus])

  useEffect(() => {
    void refresh()
  }, [refresh])

  const copy = (text: string, key: string) => {
    void api()?.copyText(text)
    setCopied(key)
    window.setTimeout(() => setCopied((c) => (c === key ? null : c)), 1500)
  }

  const run = async () => {
    if (!pending || busy) return
    setBusy(true)
    try {
      if (pending.kind === 'install') await api()!.installSkill(pending.target)
      else if ('agentId' in pending.target) await api()!.uninstallSkill(pending.target.agentId)
      setNotice(
        pending.kind === 'install' && !pending.agentId ? { text: t('intgInstalledHint') } : null,
      )
      setPending(null)
      await refresh()
    } finally {
      setBusy(false)
    }
  }

  const askInstall = (row: AgentRow, note?: string) => {
    setNotice(null)
    setPending({
      kind: 'install',
      path: row.state.path,
      note,
      target: { agentId: row.id },
      agentId: row.id,
    })
  }

  const askUninstall = (row: AgentRow) => {
    setNotice(null)
    setPending({
      kind: 'uninstall',
      path: row.state.path,
      target: { agentId: row.id },
      agentId: row.id,
    })
  }

  const installElsewhere = async () => {
    const dir = await api()?.pickSkillDir(t('intgPickDirTitle'))
    if (!dir) return
    setNotice(null)
    const sep = dir.includes('\\') && !dir.includes('/') ? '\\' : '/'
    setPending({
      kind: 'install',
      path: `${dir}${sep}genoffice${sep}SKILL.md`,
      target: { dir },
    })
  }

  const downloadZip = async () => {
    const path = await api()?.saveSkillZip(t('intgSaveZipTitle'))
    if (path) setNotice({ text: t('intgSavedTo', { path }) })
  }

  const confirmBlock = (p: Pending) => (
    <div className="set-intg-confirm" role="group">
      <div className="set-field-desc">
        {p.kind === 'install'
          ? t('intgConfirmWrite', { path: p.path })
          : t('intgConfirmRemove', { path: p.path })}
        {p.note ? ` ${p.note}` : ''}
      </div>
      <div className="set-intg-actions">
        <button className="set-btn" disabled={busy} onClick={() => setPending(null)}>
          {t('cancel')}
        </button>
        <button className="set-btn primary" disabled={busy} onClick={() => void run()}>
          {t('intgConfirm')}
        </button>
      </div>
    </div>
  )

  if (!status) {
    return (
      <>
        <h3 className="set-pane-title">{t('setSecIntegrations')}</h3>
        <div className="set-field-desc">{t('intgLoading')}</div>
      </>
    )
  }

  const bundled = status.skillVersion
  const cliTooOld =
    status.skillNeedsCli &&
    status.cli.version &&
    compare(status.cli.version, status.skillNeedsCli) < 0
  const anyInstalled = status.agents.some((a) => a.state.status !== 'missing')

  return (
    <>
      <h3 className="set-pane-title">{t('setSecIntegrations')}</h3>

      <section className="set-intg-hero">
        <div className="set-intg-hero-title">{t('intgHeroTitle')}</div>
        <div className="set-field-desc">{t('intgHeroDesc')}</div>
        <ol className="set-intg-hero-steps">
          <li>{t('intgHeroStep1')}</li>
          <li>{t('intgHeroStep2')}</li>
          <li>{t('intgHeroStep3')}</li>
        </ol>
      </section>

      <h4 className="set-pane-subtitle set-intg-step">
        <span className="set-intg-step-no">1</span>
        {t('intgStep1Title')}
        <span className="set-intg-version">{t('intgSkillVersion', { v: bundled })}</span>
      </h4>
      <div className="set-field-desc set-intg-lead">
        {t('intgStep1Desc')} {t('intgStep1Update')}
      </div>
      {cliTooOld && (
        <div className="set-field-desc set-intg-warn">
          {t('intgCliNeedsUpdate', { v: status.skillNeedsCli })}
        </div>
      )}

      {status.agents.map((row) => {
        const { state } = row
        return (
          <div key={row.id} className="set-intg-row" data-agent={row.id} data-state={state.status}>
            <div className="set-field">
              <div className="set-field-text">
                <div className="set-field-stack">
                  <div className="set-field-label">{row.label}</div>
                  <div className="set-field-desc set-intg-state" data-tip={state.path}>
                    {stateText(t, state, bundled)}
                  </div>
                </div>
              </div>
              <div className="set-intg-actions">
                {state.status === 'missing' && (
                  <button className="set-btn primary" onClick={() => askInstall(row)}>
                    {t('intgInstall')}
                  </button>
                )}
                {(state.status === 'outdated' || (state.status === 'foreign' && state.older)) && (
                  <button className="set-btn primary" onClick={() => askInstall(row)}>
                    {t('intgUpdate')}
                  </button>
                )}
                {state.status === 'modified' && (
                  <button
                    className="set-btn primary"
                    onClick={() => askInstall(row, t('intgConfirmOverwriteModified'))}
                  >
                    {t('intgUpdate')}
                  </button>
                )}
                {state.status === 'newer' && (
                  <button className="set-btn" onClick={() => askInstall(row)}>
                    {t('intgReinstall', { v: bundled })}
                  </button>
                )}
                {state.status === 'occupied' && (
                  <button
                    className="set-btn"
                    onClick={() => askInstall(row, t('intgConfirmOverwriteOccupied'))}
                  >
                    {t('intgOverwrite')}
                  </button>
                )}
                {(state.status === 'installed' ||
                  state.status === 'outdated' ||
                  state.status === 'modified') && (
                  <button className="set-btn" onClick={() => askUninstall(row)}>
                    {t('intgUninstall')}
                  </button>
                )}
              </div>
            </div>
            {pending?.agentId === row.id && confirmBlock(pending)}
            {notice?.agentId === row.id && (
              <div className="set-field-desc set-intg-notice">{notice.text}</div>
            )}
          </div>
        )
      })}

      <details className="set-intg-details" open={status.agents.length === 0}>
        <summary>{t('intgOtherToggle')}</summary>
        <div className="set-field-desc set-intg-lead">{t('intgOtherDesc')}</div>

        <div className="set-intg-option">
          <span className="set-intg-option-letter">A</span>
          <div className="set-field-stack">
            <div className="set-field-label">{t('intgOtherFolderTitle')}</div>
            <div className="set-field-desc">{t('intgOtherFolderDesc')}</div>
          </div>
          <button className="set-btn" onClick={() => void installElsewhere()}>
            {t('intgInstallElsewhere')}
          </button>
        </div>

        <div className="set-intg-option">
          <span className="set-intg-option-letter">B</span>
          <div className="set-field-stack">
            <div className="set-field-label">{t('intgOtherZipTitle')}</div>
            <div className="set-field-desc">{t('intgOtherZipDesc')}</div>
          </div>
          <button className="set-btn" onClick={() => void downloadZip()}>
            {t('intgDownloadZip')}
          </button>
        </div>

        <div className="set-intg-option">
          <span className="set-intg-option-letter">C</span>
          <div className="set-field-stack">
            <div className="set-field-label">{t('intgOtherNpxTitle')}</div>
            <div className="set-field-desc">{t('intgOtherNpxDesc')}</div>
            <div className="set-intg-code">
              <code>{NPX_INSTALL_COMMAND}</code>
              <button className="set-btn" onClick={() => copy(NPX_INSTALL_COMMAND, 'npx')}>
                {copied === 'npx' ? t('intgCopied') : t('intgCopy')}
              </button>
            </div>
          </div>
        </div>

        {pending && !pending.agentId && confirmBlock(pending)}
        {notice && !notice.agentId && (
          <div className="set-field-desc set-intg-notice">{notice.text}</div>
        )}
      </details>

      <h4 className="set-pane-subtitle set-intg-step">
        <span className="set-intg-step-no">2</span>
        {t('intgStep2Title')}
      </h4>
      <div className="set-field-desc set-intg-lead">
        {anyInstalled ? t('intgStep2Desc') : t('intgStep2DescBefore')}
      </div>
      <div className="set-intg-examples">
        {EXAMPLE_KEYS.map((key) => (
          <div key={key} className="set-intg-example">
            <span className="set-intg-example-text">“{t(key)}”</span>
            <button className="set-btn" onClick={() => copy(t(key), key)}>
              {copied === key ? t('intgCopied') : t('intgCopy')}
            </button>
          </div>
        ))}
      </div>
      <div className="set-field-desc set-intg-lead">{t('intgStep2Note')}</div>

      <details className="set-intg-details set-intg-cli">
        <summary>{t('intgCliTitle')}</summary>
        <div className="set-field-desc set-intg-lead">
          {status.cli.ephemeral
            ? t('intgCliEphemeral')
            : status.cli.status === 'present'
              ? t('intgCliReady', { v: status.cli.version, path: status.cli.location ?? '' })
              : t('intgCliNotOnPath', { v: status.cli.version })}
        </div>
        {status.cli.status !== 'present' && status.cli.manual && !status.cli.ephemeral && (
          <div className="set-intg-code">
            <code>{status.cli.manual}</code>
            <button className="set-btn" onClick={() => copy(status.cli.manual!, 'manual')}>
              {copied === 'manual' ? t('intgCopied') : t('intgCopy')}
            </button>
          </div>
        )}
        <div className="set-field">
          <div className="set-field-text">
            <div className="set-field-stack">
              <div className="set-field-desc">{t('intgCliLauncher')}</div>
              <div className="set-field-value set-intg-path" data-tip={status.cli.launcherDir}>
                {status.cli.launcherDir}
              </div>
            </div>
          </div>
          <div className="set-intg-actions">
            <button className="set-btn" onClick={() => copy(status.cli.launcherDir, 'path')}>
              {copied === 'path' ? t('intgCopied') : t('intgCopyPath')}
            </button>
          </div>
        </div>
      </details>
    </>
  )
}

function stateText(t: TFunc, state: SkillInstallState, bundled: string): string {
  const v = state.installedVersion ?? ''
  switch (state.status) {
    case 'missing':
      return t('intgStateMissing')
    case 'installed':
      return t('intgStateInstalled', { v })
    case 'outdated':
      return t('intgStateOutdated', { v, next: bundled })
    case 'modified':
      return t('intgStateModified', { v })
    case 'foreign':
      return t('intgStateForeign', { v })
    case 'newer':
      return t('intgStateNewer', { v })
    case 'occupied':
      return t('intgStateOccupied')
  }
}

function compare(a: string, b: string): number {
  const pa = a.split('.').map(Number)
  const pb = b.split('.').map(Number)
  for (let i = 0; i < 3; i++) {
    const d = (pa[i] ?? 0) - (pb[i] ?? 0)
    if (d) return d
  }
  return 0
}
