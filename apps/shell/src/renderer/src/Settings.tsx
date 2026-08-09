import { useCallback, useEffect, useRef, useState } from 'react'
import type { ReactNode } from 'react'
import type {
  AgentRules,
  AiModelSettings,
  InstructionScope,
  ReasoningEffort,
  SettingsSection,
  SettingsSnapshot,
  UserSkill,
} from '../../shared/settings-api'
import { INSTRUCTION_SCOPES, SETTINGS_SECTIONS } from '../../shared/settings-api'
import { useI18n } from './locale'
import type { StringKey } from './locale'

declare global {
  interface Window {
    genofficeSettings: import('../../shared/settings-api').SettingsApi
    genofficeSettingsEvents: { onChanged(handler: () => void): () => void }
  }
}

const REASONING_EFFORTS: readonly ReasoningEffort[] = ['minimal', 'low', 'medium', 'high']
const TEMPERATURE_RANGE = { min: 0, max: 2 }
const MAX_TOKENS_RANGE = { min: 1, max: 1_000_000 }

/** surfaces a rule or skill can target, in the order the UI lists them */
const SCOPE_LABEL: Record<InstructionScope, StringKey> = {
  global: 'setScopeGlobal',
  docx: 'setScopeDocx',
  pptx: 'setScopePptx',
  sheets: 'setScopeSheets',
  pdf: 'setScopePdf',
}

const SECTION_LABEL: Record<SettingsSection, StringKey> = {
  model: 'setSecModel',
  network: 'setSecNetwork',
  rules: 'setSecRules',
  skills: 'setSecSkills',
  general: 'setSecGeneral',
}

/** blank means "don't send"; anything else must parse and be in range */
function parseOptionalNumber(
  text: string,
  range: { min: number; max: number },
  integer = false,
): { value: number | null; invalid: boolean } {
  const trimmed = text.trim()
  if (!trimmed) return { value: null, invalid: false }
  const parsed = Number(trimmed)
  if (!Number.isFinite(parsed) || parsed < range.min || parsed > range.max) {
    return { value: null, invalid: true }
  }
  if (integer && !Number.isInteger(parsed)) return { value: null, invalid: true }
  return { value: parsed, invalid: false }
}

export function SettingsApp({ initial }: { initial: SettingsSnapshot }) {
  const { t } = useI18n()
  const [section, setSection] = useState<SettingsSection>('model')
  const [snapshot, setSnapshot] = useState<SettingsSnapshot>(initial)

  const reload = useCallback(() => {
    void window.genofficeSettings.load().then(setSnapshot)
  }, [])

  // main pushes this after an out-of-window change, e.g. a browser sign-in
  useEffect(() => window.genofficeSettingsEvents.onChanged(reload), [reload])

  return (
    <div className="settings">
      <nav className="settings-nav" aria-label={t('setTitle')}>
        <div className="settings-nav-title">{t('setTitle')}</div>
        {SETTINGS_SECTIONS.map((id) => (
          <button
            key={id}
            className={`settings-nav-item${section === id ? ' active' : ''}`}
            aria-current={section === id}
            onClick={() => setSection(id)}
          >
            {t(SECTION_LABEL[id])}
          </button>
        ))}
      </nav>
      <main className="settings-body">
        {section === 'model' && (
          <ModelSection ai={snapshot.ai} account={snapshot.account} onSaved={reload} />
        )}
        {section === 'network' && <NetworkSection ai={snapshot.ai} onSaved={reload} />}
        {section === 'rules' && (
          <RulesSection rules={snapshot.rules} memories={snapshot.memories} onSaved={reload} />
        )}
        {section === 'skills' && <SkillsSection skills={snapshot.skills} onChanged={reload} />}
        {section === 'general' && <GeneralSection snapshot={snapshot} onChanged={reload} />}
      </main>
    </div>
  )
}

// ── shared bits ─────────────────────────────────────────────────────

function Section({ title, hint, children }: { title: string; hint?: string; children: ReactNode }) {
  return (
    <section className="settings-section">
      <h2>{title}</h2>
      {hint && <p className="settings-hint">{hint}</p>}
      {children}
    </section>
  )
}

function SavedFlag({ shown }: { shown: boolean }) {
  const { t } = useI18n()
  return shown ? <span className="settings-saved">{t('setSaved')}</span> : null
}

/** save button + transient "Saved" flag, the pattern every section uses */
function useSaver<T>(save: (value: T) => Promise<unknown>, onSaved: () => void) {
  const [busy, setBusy] = useState(false)
  const [saved, setSaved] = useState(false)
  const timer = useRef<number | null>(null)
  useEffect(
    () => () => {
      if (timer.current !== null) window.clearTimeout(timer.current)
    },
    [],
  )
  const run = (value: T) => {
    setBusy(true)
    void save(value)
      .then(() => {
        setSaved(true)
        onSaved()
        if (timer.current !== null) window.clearTimeout(timer.current)
        timer.current = window.setTimeout(() => setSaved(false), 2000)
      })
      .finally(() => setBusy(false))
  }
  return { busy, saved, run }
}

// ── Model ───────────────────────────────────────────────────────────

function ModelSection({
  ai,
  account,
  onSaved,
}: {
  ai: AiModelSettings
  account: SettingsSnapshot['account']
  onSaved: () => void
}) {
  const { t } = useI18n()
  const [draft, setDraft] = useState<AiModelSettings>(ai)
  const [showKey, setShowKey] = useState(false)
  const [tempText, setTempText] = useState(ai.temperature === null ? '' : String(ai.temperature))
  const [maxTokText, setMaxTokText] = useState(ai.maxTokens === null ? '' : String(ai.maxTokens))
  const [test, setTest] = useState<{ state: 'idle' | 'busy' | 'ok' | 'fail'; error?: string }>({
    state: 'idle',
  })

  const custom = draft.mode === 'custom'
  const temperature = parseOptionalNumber(tempText, TEMPERATURE_RANGE)
  const maxTokens = parseOptionalNumber(maxTokText, MAX_TOKENS_RANGE, true)
  const badNumber = custom && (temperature.invalid || maxTokens.invalid)
  const incomplete = custom && (!draft.baseUrl.trim() || !draft.model.trim())
  const blocked = incomplete || badNumber

  const resolved = (): AiModelSettings => ({
    ...draft,
    temperature: temperature.value,
    maxTokens: maxTokens.value,
  })

  const saver = useSaver<AiModelSettings>((v) => window.genofficeSettings.saveAi(v), onSaved)

  /** the endpoint fields edit the selected row, so keep the row in step */
  const edit = (patch: Partial<AiModelSettings>) => {
    setDraft((prev) => {
      const next = { ...prev, ...patch }
      const touchesEndpoint =
        'baseUrl' in patch || 'model' in patch || 'apiKey' in patch || 'label' in patch
      if (touchesEndpoint && next.profileId) {
        next.profiles = next.profiles.map((p) =>
          p.id === next.profileId
            ? { ...p, baseUrl: next.baseUrl, model: next.model, apiKey: next.apiKey }
            : p,
        )
      }
      return next
    })
    setTest({ state: 'idle' })
  }

  /** Show a different row: the fields describe whichever row is selected. */
  const selectProfile = (id: string) => {
    if (id === draft.profileId) return
    setDraft((prev) => {
      const row = prev.profiles.find((p) => p.id === id)
      if (!row) return prev
      return { ...prev, profileId: id, baseUrl: row.baseUrl, model: row.model, apiKey: row.apiKey }
    })
    setTest({ state: 'idle' })
  }

  const addProfile = () => {
    // a client-side id; the main process keeps whatever it is handed, so a row
    // added and saved in one go stays the same profile afterwards
    const id = `new-${Date.now().toString(36)}`
    setDraft((prev) => ({
      ...prev,
      profiles: [...prev.profiles, { id, label: '', baseUrl: '', model: '', apiKey: '' }],
      profileId: id,
      baseUrl: '',
      model: '',
      apiKey: '',
    }))
    setTest({ state: 'idle' })
  }

  const removeProfile = (id: string) => {
    setDraft((prev) => {
      const profiles = prev.profiles.filter((p) => p.id !== id)
      if (prev.profileId !== id) return { ...prev, profiles }
      // the selection went with it: fall to the first row, or to no model at all
      const next = profiles[0]
      return {
        ...prev,
        profiles,
        profileId: next?.id ?? null,
        baseUrl: next?.baseUrl ?? '',
        model: next?.model ?? '',
        apiKey: next?.apiKey ?? '',
      }
    })
    setTest({ state: 'idle' })
  }

  const renameProfile = (label: string) => {
    setDraft((prev) => ({
      ...prev,
      profiles: prev.profiles.map((p) => (p.id === prev.profileId ? { ...p, label } : p)),
    }))
  }

  const runTest = () => {
    if (blocked) return
    setTest({ state: 'busy' })
    void window.genofficeSettings
      .testProvider(resolved())
      .then((r) => setTest(r.ok ? { state: 'ok' } : { state: 'fail', error: r.error }))
      .catch((err: unknown) => setTest({ state: 'fail', error: String(err) }))
  }

  return (
    <Section title={t('setSecModel')} hint={t('aiDlgScope')}>
      <div className="ai-mode-list" role="radiogroup" aria-label={t('setSecModel')}>
        <button
          type="button"
          role="radio"
          aria-checked={!custom}
          className={`ai-mode${!custom ? ' active' : ''}`}
          onClick={() => edit({ mode: 'genspark' })}
        >
          <span className="ai-mode-dot" aria-hidden="true" />
          <span className="ai-mode-body">
            <span className="ai-mode-title">
              {t('accountGenspark')}
              <span className={`ai-mode-badge${account.loggedIn ? ' on' : ''}`}>
                {account.loggedIn ? account.email || t('loggedIn') : t('aiDlgNotSignedIn')}
              </span>
            </span>
            <span className="ai-mode-desc">{t('aiDlgGensparkDesc')}</span>
          </span>
        </button>
        <button
          type="button"
          role="radio"
          aria-checked={custom}
          className={`ai-mode${custom ? ' active' : ''}`}
          onClick={() => edit({ mode: 'custom' })}
        >
          <span className="ai-mode-dot" aria-hidden="true" />
          <span className="ai-mode-body">
            <span className="ai-mode-title">{t('aiDlgCustom')}</span>
            <span className="ai-mode-desc">{t('aiDlgCustomDesc')}</span>
          </span>
        </button>
      </div>

      {!custom && !account.loggedIn && (
        <button
          className="ai-inline-login"
          onClick={() => void window.genofficeSettings.accountLogin()}
        >
          {t('loginGenspark')}
        </button>
      )}

      {custom && (
        <div className="ai-fields">
          {/* the library the sidebar switches between; the fields below edit the selected row */}
          <div className="ai-profile-list" role="radiogroup" aria-label={t('aiDlgModels')}>
            {draft.profiles.map((p) => (
              <div
                key={p.id}
                className={`ai-profile${p.id === draft.profileId ? ' active' : ''}`}
                role="radio"
                aria-checked={p.id === draft.profileId}
                tabIndex={0}
                onClick={() => selectProfile(p.id)}
                onKeyDown={(e) => {
                  if (e.key === 'Enter' || e.key === ' ') {
                    e.preventDefault()
                    selectProfile(p.id)
                  }
                }}
              >
                <span className="ai-profile-name">{p.label || p.model || t('aiDlgUnnamed')}</span>
                <span className="ai-profile-model">{p.model}</span>
                <button
                  type="button"
                  className="ai-profile-remove"
                  title={t('aiDlgRemoveModel')}
                  onClick={(e) => {
                    e.stopPropagation()
                    removeProfile(p.id)
                  }}
                >
                  ×
                </button>
              </div>
            ))}
            <button type="button" className="ai-profile-add" onClick={addProfile}>
              {t('aiDlgAddModel')}
            </button>
          </div>
          <label className="ai-field">
            <span className="ai-field-label">{t('aiDlgModelName')}</span>
            <input
              className="ai-input"
              spellCheck={false}
              placeholder={t('aiDlgModelNamePlaceholder')}
              value={draft.profiles.find((p) => p.id === draft.profileId)?.label ?? ''}
              onChange={(e) => renameProfile(e.target.value)}
            />
          </label>
          <label className="ai-field">
            <span className="ai-field-label">Base URL</span>
            <input
              className="ai-input"
              spellCheck={false}
              placeholder="https://api.deepseek.com/v1"
              value={draft.baseUrl}
              onChange={(e) => edit({ baseUrl: e.target.value })}
            />
            <span className="ai-field-hint">{t('aiDlgBaseUrlHint')}</span>
          </label>
          <label className="ai-field">
            <span className="ai-field-label">Model</span>
            <input
              className="ai-input"
              spellCheck={false}
              placeholder="deepseek-chat"
              value={draft.model}
              onChange={(e) => edit({ model: e.target.value })}
            />
          </label>
          <label className="ai-field">
            <span className="ai-field-label">API Key</span>
            <span className="ai-input-wrap">
              <input
                className="ai-input"
                type={showKey ? 'text' : 'password'}
                spellCheck={false}
                placeholder="sk-..."
                value={draft.apiKey}
                onChange={(e) => edit({ apiKey: e.target.value })}
              />
              <button type="button" className="ai-key-toggle" onClick={() => setShowKey((v) => !v)}>
                {showKey ? t('aiDlgHideKey') : t('aiDlgShowKey')}
              </button>
            </span>
            <span className="ai-field-hint">{t('aiDlgKeyOptional')}</span>
          </label>

          <div className="ai-tuning">
            <label className="ai-field">
              <span className="ai-field-label">Temperature</span>
              <input
                className={`ai-input${temperature.invalid ? ' invalid' : ''}`}
                inputMode="decimal"
                placeholder={t('aiDlgModelDefault')}
                value={tempText}
                onChange={(e) => {
                  setTempText(e.target.value)
                  setTest({ state: 'idle' })
                }}
              />
            </label>
            <label className="ai-field">
              <span className="ai-field-label">Max tokens</span>
              <input
                className={`ai-input${maxTokens.invalid ? ' invalid' : ''}`}
                inputMode="numeric"
                placeholder={t('aiDlgModelDefault')}
                value={maxTokText}
                onChange={(e) => {
                  setMaxTokText(e.target.value)
                  setTest({ state: 'idle' })
                }}
              />
            </label>
            <label className="ai-field">
              <span className="ai-field-label">Reasoning effort</span>
              <select
                className="ai-input"
                value={draft.reasoningEffort ?? ''}
                onChange={(e) =>
                  edit({ reasoningEffort: (e.target.value || null) as ReasoningEffort | null })
                }
              >
                <option value="">{t('aiDlgModelDefault')}</option>
                {REASONING_EFFORTS.map((eff) => (
                  <option key={eff} value={eff}>
                    {eff}
                  </option>
                ))}
              </select>
            </label>
          </div>
          <span className={`ai-field-hint${badNumber ? ' error' : ''}`}>
            {badNumber
              ? t('aiDlgBadNumber', { min: TEMPERATURE_RANGE.min, max: TEMPERATURE_RANGE.max })
              : t('aiDlgTuningHint')}
          </span>

          <div className="ai-test-row">
            <button
              className="btn btn-secondary btn-sm"
              disabled={blocked || test.state === 'busy'}
              onClick={runTest}
            >
              {test.state === 'busy' ? t('aiDlgTesting') : t('aiDlgTest')}
            </button>
            {test.state === 'ok' && <span className="ai-test-ok">{t('aiDlgTestOk')}</span>}
            {test.state === 'fail' && (
              <span className="ai-test-fail" title={test.error}>
                {test.error ?? t('aiDlgTestFail')}
              </span>
            )}
          </div>
        </div>
      )}

      <p className="settings-note">{t('setModelRecommendation')}</p>

      <div className="settings-actions">
        <SavedFlag shown={saver.saved} />
        <button
          className="btn btn-primary"
          disabled={blocked || saver.busy}
          onClick={() => saver.run(resolved())}
        >
          {t('aiDlgSave')}
        </button>
      </div>
    </Section>
  )
}

// ── Network ─────────────────────────────────────────────────────────

function NetworkSection({ ai, onSaved }: { ai: AiModelSettings; onSaved: () => void }) {
  const { t } = useI18n()
  const [tavily, setTavily] = useState(ai.tavilyApiKey)
  const [proxy, setProxy] = useState(ai.proxyUrl)
  const [showKey, setShowKey] = useState(false)
  const saver = useSaver<AiModelSettings>((v) => window.genofficeSettings.saveAi(v), onSaved)

  return (
    <Section title={t('setSecNetwork')} hint={t('setNetworkHint')}>
      <div className="ai-fields">
        <label className="ai-field">
          <span className="ai-field-label">Tavily API Key</span>
          <span className="ai-input-wrap">
            <input
              className="ai-input"
              type={showKey ? 'text' : 'password'}
              spellCheck={false}
              placeholder="tvly-..."
              value={tavily}
              onChange={(e) => setTavily(e.target.value)}
            />
            <button type="button" className="ai-key-toggle" onClick={() => setShowKey((v) => !v)}>
              {showKey ? t('aiDlgHideKey') : t('aiDlgShowKey')}
            </button>
          </span>
          <span className="ai-field-hint">{t('setTavilyHint')}</span>
        </label>

        <label className="ai-field">
          <span className="ai-field-label">{t('setProxyLabel')}</span>
          <input
            className="ai-input"
            spellCheck={false}
            placeholder="http://127.0.0.1:7890"
            value={proxy}
            onChange={(e) => setProxy(e.target.value)}
          />
          <span className="ai-field-hint">{t('setProxyHint')}</span>
        </label>
      </div>

      <div className="settings-actions">
        <SavedFlag shown={saver.saved} />
        <button
          className="btn btn-primary"
          disabled={saver.busy}
          onClick={() => saver.run({ ...ai, tavilyApiKey: tavily, proxyUrl: proxy })}
        >
          {t('aiDlgSave')}
        </button>
      </div>
    </Section>
  )
}

// ── Rules ───────────────────────────────────────────────────────────

function RulesSection({
  rules,
  memories,
  onSaved,
}: {
  rules: AgentRules
  memories: SettingsSnapshot['memories']
  onSaved: () => void
}) {
  const { t } = useI18n()
  const [draft, setDraft] = useState<AgentRules>(rules)
  const saver = useSaver<AgentRules>((v) => window.genofficeSettings.saveRules(v), onSaved)

  return (
    <Section title={t('setSecRules')} hint={t('setRulesHint')}>
      {INSTRUCTION_SCOPES.map((scope) => (
        <label key={scope} className="ai-field settings-rule">
          <span className="ai-field-label">{t(SCOPE_LABEL[scope])}</span>
          <textarea
            className="settings-textarea"
            rows={scope === 'global' ? 6 : 4}
            placeholder={scope === 'global' ? t('setRulesGlobalPh') : t('setRulesScopedPh')}
            value={draft[scope] ?? ''}
            onChange={(e) => setDraft((prev) => ({ ...prev, [scope]: e.target.value }))}
          />
        </label>
      ))}
      <div className="settings-actions">
        <SavedFlag shown={saver.saved} />
        <button className="btn btn-primary" disabled={saver.busy} onClick={() => saver.run(draft)}>
          {t('aiDlgSave')}
        </button>
      </div>
      <h3 className="settings-subhead">{t('setMemoryTitle')}</h3>
      <p className="settings-hint">{t('setMemoryHint')}</p>
      <MemorySection memories={memories} onChanged={onSaved} />
    </Section>
  )
}

/**
 * What the agent recorded about the user, and the means to delete it. This is
 * the counterweight to letting the agent write to its own prompt: memories are
 * written from conversation, a conversation can contain document text, so the
 * user has to be able to see the list and remove anything that does not belong.
 */
function MemorySection({
  memories,
  onChanged,
}: {
  memories: SettingsSnapshot['memories']
  onChanged: () => void
}) {
  const { t } = useI18n()
  const [busy, setBusy] = useState<string | null>(null)

  const remove = (id: string) => {
    setBusy(id)
    void window.genofficeSettings
      .deleteMemory(id)
      .then(onChanged)
      .finally(() => setBusy(null))
  }

  if (memories.length === 0) {
    return <p className="settings-hint">{t('setMemoryEmpty')}</p>
  }
  return (
    <ul className="settings-skills">
      {memories.map((m) => (
        <li key={m.id} className="settings-skill">
          <div className="settings-skill-head">
            <span className="settings-memory-text">{m.text}</span>
            <span className="settings-skill-id">
              {m.createdAt ? new Date(m.createdAt).toLocaleDateString() : ''}
            </span>
          </div>
          <div className="settings-actions">
            <button
              className="btn btn-secondary"
              disabled={busy === m.id}
              onClick={() => remove(m.id)}
            >
              {t('setMemoryForget')}
            </button>
          </div>
        </li>
      ))}
    </ul>
  )
}

// ── Skills ──────────────────────────────────────────────────────────

const BLANK_SKILL: UserSkill = {
  id: '',
  name: '',
  description: '',
  scopes: ['global'],
  body: '',
  enabled: true,
}

function SkillsSection({ skills, onChanged }: { skills: UserSkill[]; onChanged: () => void }) {
  const { t } = useI18n()
  const [editing, setEditing] = useState<UserSkill | null>(null)
  const [dropActive, setDropActive] = useState(false)
  const [busy, setBusy] = useState(false)

  const importFiles = () => {
    setBusy(true)
    void window.genofficeSettings
      .importSkillFiles()
      .then(onChanged)
      .finally(() => setBusy(false))
  }

  /** drag-and-drop: read the dropped .md files here, then hand main the text */
  const onDrop = (event: React.DragEvent) => {
    event.preventDefault()
    setDropActive(false)
    const files = Array.from(event.dataTransfer.files).filter((f) =>
      /\.(md|markdown)$/i.test(f.name),
    )
    if (!files.length) return
    setBusy(true)
    void Promise.all(
      files.slice(0, 50).map(async (f) => ({ filename: f.name, content: await f.text() })),
    )
      .then((payload) => window.genofficeSettings.importSkillContents(payload))
      .then(onChanged)
      .finally(() => setBusy(false))
  }

  if (editing) {
    return (
      <SkillEditor
        skill={editing}
        onCancel={() => setEditing(null)}
        onSaved={() => {
          setEditing(null)
          onChanged()
        }}
      />
    )
  }

  return (
    <Section title={t('setSecSkills')} hint={t('setSkillsHint')}>
      <div
        className={`settings-drop${dropActive ? ' active' : ''}`}
        onDragOver={(e) => {
          e.preventDefault()
          setDropActive(true)
        }}
        onDragLeave={() => setDropActive(false)}
        onDrop={onDrop}
      >
        <span>{t('setSkillsDrop')}</span>
        <span className="settings-drop-actions">
          <button className="btn btn-secondary btn-sm" disabled={busy} onClick={importFiles}>
            {t('setSkillsUpload')}
          </button>
          <button
            className="btn btn-secondary btn-sm"
            onClick={() => setEditing({ ...BLANK_SKILL })}
          >
            {t('setSkillsNew')}
          </button>
        </span>
      </div>

      {skills.length === 0 ? (
        <p className="settings-empty">{t('setSkillsEmpty')}</p>
      ) : (
        <ul className="settings-skill-list">
          {skills.map((skill) => (
            <li key={skill.id} className={`settings-skill${skill.enabled ? '' : ' off'}`}>
              <div className="settings-skill-head">
                <span className="settings-skill-name">{skill.name}</span>
                <code className="settings-skill-id">{skill.id}</code>
              </div>
              {skill.description && <div className="settings-skill-desc">{skill.description}</div>}
              <div className="settings-skill-scopes">
                {skill.scopes.map((scope) => (
                  <span key={scope} className="settings-chip">
                    {t(SCOPE_LABEL[scope])}
                  </span>
                ))}
                {!skill.enabled && (
                  <span className="settings-chip off">{t('setSkillDisabled')}</span>
                )}
              </div>
              <div className="settings-skill-actions">
                <button className="btn btn-secondary btn-sm" onClick={() => setEditing(skill)}>
                  {t('setSkillEdit')}
                </button>
                <button
                  className="btn btn-secondary btn-sm"
                  onClick={() => {
                    void window.genofficeSettings
                      .saveSkill({ ...skill, enabled: !skill.enabled })
                      .then(onChanged)
                  }}
                >
                  {skill.enabled ? t('setSkillDisable') : t('setSkillEnable')}
                </button>
                <button
                  className="btn btn-danger btn-sm"
                  onClick={() => {
                    void window.genofficeSettings.deleteSkill(skill.id).then(onChanged)
                  }}
                >
                  {t('delete')}
                </button>
              </div>
            </li>
          ))}
        </ul>
      )}
    </Section>
  )
}

function SkillEditor({
  skill,
  onCancel,
  onSaved,
}: {
  skill: UserSkill
  onCancel: () => void
  onSaved: () => void
}) {
  const { t } = useI18n()
  const [draft, setDraft] = useState<UserSkill>(skill)
  const [busy, setBusy] = useState(false)
  const valid = draft.name.trim() && draft.body.trim()

  /**
   * "All apps" and a specific app are mutually exclusive: global already
   * covers everything, so holding both would be redundant, and letting the
   * user uncheck their last scope would make the skill invisible. Picking a
   * specific app therefore drops global, and picking global clears the rest.
   */
  const toggleScope = (scope: InstructionScope) => {
    setDraft((prev) => {
      if (scope === 'global') return { ...prev, scopes: ['global'] }
      const withoutGlobal = prev.scopes.filter((s) => s !== 'global')
      const scopes = withoutGlobal.includes(scope)
        ? withoutGlobal.filter((s) => s !== scope)
        : [...withoutGlobal, scope]
      return { ...prev, scopes: scopes.length ? scopes : ['global'] }
    })
  }

  return (
    <Section title={draft.id ? t('setSkillEdit') : t('setSkillsNew')}>
      <div className="ai-fields">
        <label className="ai-field">
          <span className="ai-field-label">{t('setSkillName')}</span>
          <input
            className="ai-input"
            value={draft.name}
            onChange={(e) => setDraft((p) => ({ ...p, name: e.target.value }))}
          />
        </label>
        <label className="ai-field">
          <span className="ai-field-label">{t('setSkillDesc')}</span>
          <input
            className="ai-input"
            value={draft.description}
            onChange={(e) => setDraft((p) => ({ ...p, description: e.target.value }))}
          />
          <span className="ai-field-hint">{t('setSkillDescHint')}</span>
        </label>
        <div className="ai-field">
          <span className="ai-field-label">{t('setSkillScopes')}</span>
          <div className="settings-scope-picker">
            {INSTRUCTION_SCOPES.map((scope) => (
              <label key={scope} className="settings-scope-option">
                <input
                  type="checkbox"
                  checked={draft.scopes.includes(scope)}
                  onChange={() => toggleScope(scope)}
                />
                {t(SCOPE_LABEL[scope])}
              </label>
            ))}
          </div>
        </div>
        <label className="ai-field">
          <span className="ai-field-label">{t('setSkillBody')}</span>
          <textarea
            className="settings-textarea mono"
            rows={14}
            placeholder={t('setSkillBodyPh')}
            value={draft.body}
            onChange={(e) => setDraft((p) => ({ ...p, body: e.target.value }))}
          />
        </label>
      </div>
      <div className="settings-actions">
        <button className="btn btn-secondary" onClick={onCancel}>
          {t('cancel')}
        </button>
        <button
          className="btn btn-primary"
          disabled={!valid || busy}
          onClick={() => {
            setBusy(true)
            void window.genofficeSettings
              .saveSkill(draft.id ? draft : { ...draft, id: undefined })
              .then(onSaved)
              .finally(() => setBusy(false))
          }}
        >
          {t('aiDlgSave')}
        </button>
      </div>
    </Section>
  )
}

// ── General ─────────────────────────────────────────────────────────

const LANG_OPTIONS = [
  { value: 'ar', label: 'العربية' },
  { value: 'de', label: 'Deutsch' },
  { value: 'en', label: 'English' },
  { value: 'es', label: 'Español' },
  { value: 'fr', label: 'Français' },
  { value: 'he', label: 'עברית' },
  { value: 'hi', label: 'हिन्दी' },
  { value: 'id', label: 'Bahasa Indonesia' },
  { value: 'it', label: 'Italiano' },
  { value: 'ja', label: '日本語' },
  { value: 'ko', label: '한국어' },
  { value: 'ms', label: 'Bahasa Melayu' },
  { value: 'nl', label: 'Nederlands' },
  { value: 'pl', label: 'Polski' },
  { value: 'pt', label: 'Português' },
  { value: 'ru', label: 'Русский' },
  { value: 'th', label: 'ไทย' },
  { value: 'zh', label: '简体中文' },
  { value: 'zh-TW', label: '繁體中文' },
] as const

function GeneralSection({
  snapshot,
  onChanged,
}: {
  snapshot: SettingsSnapshot
  onChanged: () => void
}) {
  const { t, setLang } = useI18n()
  return (
    <Section title={t('setSecGeneral')}>
      <div className="ai-fields">
        <label className="ai-field">
          <span className="ai-field-label">{t('language')}</span>
          <select
            className="ai-input"
            value={snapshot.language}
            onChange={(e) => {
              const value = e.target.value
              void window.genofficeSettings.setLanguage(value).then(onChanged)
              setLang(value as Parameters<typeof setLang>[0])
            }}
          >
            {LANG_OPTIONS.map((o) => (
              <option key={o.value} value={o.value}>
                {o.label}
              </option>
            ))}
          </select>
        </label>

        <label className="ai-field">
          <span className="ai-field-label">{t('updateChannel')}</span>
          <select
            className="ai-input"
            value={snapshot.updateChannel}
            onChange={(e) => {
              void window.genofficeSettings
                .setUpdateChannel(e.target.value as 'stable' | 'beta')
                .then(onChanged)
            }}
          >
            <option value="stable">{t('channelStable')}</option>
            <option value="beta">{t('channelBeta')}</option>
          </select>
        </label>

        <div className="ai-field">
          <span className="ai-field-label">{t('account')}</span>
          <div className="settings-account-row">
            <span>
              {snapshot.account.loggedIn
                ? snapshot.account.email || t('loggedIn')
                : t('aiDlgNotSignedIn')}
            </span>
            {snapshot.account.loggedIn ? (
              <button
                className="btn btn-secondary btn-sm"
                onClick={() => void window.genofficeSettings.accountLogout().then(onChanged)}
              >
                {t('logout')}
              </button>
            ) : (
              <button
                className="btn btn-secondary btn-sm"
                onClick={() => void window.genofficeSettings.accountLogin()}
              >
                {t('loginGenspark')}
              </button>
            )}
          </div>
        </div>

        <div className="ai-field">
          <span className="ai-field-label">{t('versionLabel')}</span>
          <span className="settings-version">{snapshot.appVersion}</span>
        </div>
      </div>
    </Section>
  )
}
