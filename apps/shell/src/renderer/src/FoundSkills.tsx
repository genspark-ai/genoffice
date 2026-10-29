import { useCallback, useEffect, useRef, useState } from 'react'
import type { StringKey, TFunc } from './locale'
import type {
  FoundSkill,
  ForeignSkillCandidate,
  ForeignSkillSource,
  RecommendedSkill,
} from '../../shared/found-skill'
import { useDismissablePopover } from '@genoffice/ui'
import { SkillCreator, SkillDocFields, SkillDocPreview } from './SkillCreator'
import { draftFieldErrors, parseSkillDoc, type SkillDraft } from './skill-draft'

/**
 * The skills the user installed into GenOffice.
 *
 * Only GenOffice's own folder is listed. A skill gets in one way: the user
 * picks a folder or a zip and presses Import, pulls one in from another agent
 * product's files, or writes one here: the pane edits the SKILL.md of a listed
 * skill, and the AI creator drafts a new one from a short interview. Nothing
 * on this screen runs anything.
 */

const api = () => window.aiOfficeIntegrations

type Notice =
  | { kind: 'done' | 'removed' | 'saved'; name: string }
  | { kind: 'foreign'; count: number }
  | { kind: 'failed'; reason: string }

/** what the pane is showing over the list: the editor, the AI creator, or the foreign importer */
type View =
  | { kind: 'create' }
  | { kind: 'edit'; skill: FoundSkill }
  | { kind: 'ai' }
  | { kind: 'foreign' }
  | null

export function FoundSkills({ t }: { t: TFunc }) {
  const [skills, setSkills] = useState<FoundSkill[] | null>(null)
  const [recommended, setRecommended] = useState<RecommendedSkill[]>([])
  const [busy, setBusy] = useState(false)
  const [notice, setNotice] = useState<Notice | null>(null)
  const [tab, setTab] = useState<Category | null>(null)
  const [section, setSection] = useState<Section | null>(null)
  const [view, setView] = useState<View>(null)
  const [importOpen, setImportOpen] = useState(false)
  const importRef = useRef<HTMLDivElement | null>(null)
  const closeImport = useCallback(() => setImportOpen(false), [])
  useDismissablePopover(importOpen, closeImport, { inside: () => [importRef.current] })

  const load = useCallback(async () => {
    // the method itself is checked, not just the bridge: a preload from before
    // this feature was added goes on answering with a partial object
    const [found, offered] = await Promise.all([
      api()?.listSkills?.(),
      Promise.resolve()
        .then(() => api()?.listRecommendedSkills?.())
        .catch(() => undefined),
    ])
    if (found) setSkills(found)
    if (offered) setRecommended(offered)
    // decided once from the first complete answer: with nothing installed the
    // catalog is the useful screen, and installing from it must not flip the view
    if (found) {
      setSection((cur) => cur ?? (found.length === 0 && offered?.length ? 'recommended' : 'mine'))
    }
  }, [])

  useEffect(() => {
    void load()
  }, [load])

  const doImport = async (run: (() => Promise<FoundSkill | null>) | undefined) => {
    if (busy || !run) return
    setBusy(true)
    try {
      const skill = await run()
      if (!skill) return
      setNotice({ kind: 'done', name: skill.name })
      await load()
    } catch (e) {
      setNotice({ kind: 'failed', reason: e instanceof Error ? e.message : String(e) })
    } finally {
      setBusy(false)
    }
  }

  const doRemove = async (skill: FoundSkill) => {
    if (busy || !window.confirm(t('intgSkillsRemoveConfirm', { name: skill.name }))) return
    setBusy(true)
    try {
      await api()?.removeSkill?.(skill.name)
      setNotice({ kind: 'removed', name: skill.name })
      await load()
    } catch (e) {
      setNotice({ kind: 'failed', reason: e instanceof Error ? e.message : String(e) })
    } finally {
      setBusy(false)
    }
  }

  const doInstall = (skill: RecommendedSkill) =>
    doImport(() => api()?.installRecommendedSkill?.(skill.name) ?? Promise.resolve(null))

  // nothing rendered until the listing answers: an empty section that fills in
  // a moment later reads as "none installed", the one thing it must not say
  if (!skills) return null

  const savedNotice = (skill: { name: string }): void => {
    setNotice({ kind: 'saved', name: skill.name })
    setView(null)
    // a saved skill lands in My skills — show that section, not the catalog the
    // user may have been browsing when they started the editor
    setSection('mine')
    void load()
  }

  const categories = CATEGORY_ORDER.filter((c) => recommended.some((s) => categoryOf(s) === c))
  const current = tab && categories.includes(tab) ? tab : categories[0]
  const shown: Section = section ?? 'mine'
  const sectionTab = (s: Section, label: string) => (
    <button
      key={s}
      type="button"
      role="tab"
      aria-selected={s === shown}
      className={s === shown ? 'set-intg-seg active' : 'set-intg-seg'}
      onClick={() => setSection(s)}
    >
      {label}
    </button>
  )

  return (
    <>
      <h3 className="set-pane-title">{t('setSecSkills')}</h3>
      <div className="set-field-desc set-intg-lead">{t('intgSkillsDesc')}</div>

      {view?.kind === 'ai' && (
        <SkillCreator t={t} onClose={() => setView(null)} onSaved={savedNotice} />
      )}

      {view && (view.kind === 'create' || view.kind === 'edit') && (
        <SkillEditor
          t={t}
          target={view.kind === 'edit' ? view.skill : null}
          onClose={() => setView(null)}
          onSaved={savedNotice}
        />
      )}

      {view?.kind === 'foreign' && (
        <ForeignSkillImport
          t={t}
          onClose={() => setView(null)}
          onRefresh={load}
          onDone={(count) => {
            setNotice({ kind: 'foreign', count })
            setView(null)
            // imported skills land in My skills — the section the notice points at
            setSection('mine')
          }}
        />
      )}

      {!view && (
        <>
          <div className="set-intg-head">
            {categories.length > 0 ? (
              <div className="set-intg-tabs set-intg-sections" role="tablist">
                {sectionTab('mine', `${t('intgSkillsMine')} (${skills.length})`)}
                {sectionTab('recommended', t('intgSkillsRecommended'))}
              </div>
            ) : (
              <div className="set-field-label">{t('intgSkillsMine')}</div>
            )}
            <div className="set-intg-actions">
              <button
                className="set-btn primary"
                disabled={busy}
                onClick={() => setView({ kind: 'ai' })}
              >
                {t('intgSkillsAiCreate')}
              </button>
              <button
                className="set-btn"
                disabled={busy}
                onClick={() => setView({ kind: 'create' })}
              >
                {t('intgSkillsNew')}
              </button>
              <div className="set-intg-import" ref={importRef}>
                <button
                  className="set-btn"
                  disabled={busy}
                  aria-haspopup="menu"
                  aria-expanded={importOpen}
                  onClick={() => setImportOpen((v) => !v)}
                >
                  {t('intgSkillsImport')}
                  <svg width="10" height="10" viewBox="0 0 10 10" aria-hidden>
                    <path
                      d="M2 3.5l3 3 3-3"
                      fill="none"
                      stroke="currentColor"
                      strokeWidth="1.4"
                      strokeLinecap="round"
                      strokeLinejoin="round"
                    />
                  </svg>
                </button>
                {importOpen && (
                  <div className="set-intg-import-menu" role="menu">
                    <button
                      type="button"
                      role="menuitem"
                      className="set-intg-import-item"
                      onClick={() => {
                        setImportOpen(false)
                        void doImport(
                          () =>
                            api()?.importSkillZip?.(t('intgSkillsImportZipTitle')) ??
                            Promise.resolve(null),
                        )
                      }}
                    >
                      {t('intgSkillsImportZip')}
                    </button>
                    <button
                      type="button"
                      role="menuitem"
                      className="set-intg-import-item"
                      onClick={() => {
                        setImportOpen(false)
                        void doImport(
                          () =>
                            api()?.importSkillFolder?.(t('intgSkillsImportFolderTitle')) ??
                            Promise.resolve(null),
                        )
                      }}
                    >
                      {t('intgSkillsImportFolder')}
                    </button>
                    <button
                      type="button"
                      role="menuitem"
                      className="set-intg-import-item"
                      onClick={() => {
                        setImportOpen(false)
                        setView({ kind: 'foreign' })
                      }}
                    >
                      {t('intgSkillsImportForeign')}
                    </button>
                  </div>
                )}
              </div>
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
                ? t('intgSkillsImportFailed', { reason: notice.reason })
                : notice.kind === 'done'
                  ? t('intgSkillsImportDone', { name: notice.name })
                  : notice.kind === 'removed'
                    ? t('intgSkillsRemoved', { name: notice.name })
                    : notice.kind === 'foreign'
                      ? t(notice.count === 1 ? 'intgSkillForeignDoneOne' : 'intgSkillForeignDone', {
                          n: notice.count,
                        })
                      : t('intgSkillsSaved', { name: notice.name })}
            </div>
          )}

          {shown === 'mine' && skills.length === 0 && (
            <div className="set-field-desc set-intg-lead">{t('intgSkillsEmpty')}</div>
          )}

          {shown === 'mine' && skills.length > 0 && (
            <div className="set-intg-list">
              {skills.map((skill) => (
                <SkillRow
                  key={skill.path}
                  t={t}
                  skill={skill}
                  action={
                    <>
                      <button
                        className="set-btn"
                        disabled={busy}
                        onClick={() => setView({ kind: 'edit', skill })}
                      >
                        {t('intgSkillsEdit')}
                      </button>
                      <button
                        className="set-btn"
                        disabled={busy}
                        onClick={() => void doRemove(skill)}
                      >
                        {t('intgSkillsRemove')}
                      </button>
                    </>
                  }
                />
              ))}
            </div>
          )}

          {shown === 'recommended' && categories.length > 0 && (
            <div className="set-intg-list">
              <div className="set-intg-tabs" role="tablist" aria-label={t('intgSkillsRecommended')}>
                {categories.map((c) => (
                  <button
                    key={c}
                    type="button"
                    role="tab"
                    aria-selected={c === current}
                    className={c === current ? 'set-intg-tab active' : 'set-intg-tab'}
                    onClick={() => setTab(c)}
                  >
                    {t(CATEGORY_LABELS[c])}
                  </button>
                ))}
              </div>
              <div role="tabpanel">
                {recommended
                  .filter((s) => categoryOf(s) === current)
                  .map((skill) => (
                    <SkillRow
                      key={skill.path}
                      t={t}
                      skill={skill}
                      action={
                        skill.state === 'installed' ? (
                          <span className="set-intg-state">{t('intgSkillsInstalledTag')}</span>
                        ) : (
                          <button
                            className="set-btn"
                            disabled={busy}
                            onClick={() => void doInstall(skill)}
                          >
                            {skill.state === 'outdated'
                              ? t('intgSkillsUpdate')
                              : t('intgSkillsInstall')}
                          </button>
                        )
                      }
                    />
                  ))}
              </div>
            </div>
          )}
        </>
      )}
    </>
  )
}

/**
 * The create/edit form for one skill: the three fields the main process's
 * writeSkill builds the SKILL.md from, with a live read-only preview of that
 * file beside them. Creating starts empty; editing loads the listed skill's
 * raw text and locks the name — it is the folder on disk, and what the
 * overwrite replaces.
 */
function SkillEditor({
  t,
  target,
  onClose,
  onSaved,
}: {
  t: TFunc
  /** null = create; the listed skill being edited otherwise */
  target: FoundSkill | null
  onClose: () => void
  onSaved: (skill: FoundSkill) => void
}) {
  const [draft, setDraft] = useState<SkillDraft | null>(
    target ? null : { name: '', description: '', body: '' },
  )
  const [error, setError] = useState('')
  const [busy, setBusy] = useState(false)

  useEffect(() => {
    if (!target) return
    let alive = true
    void window.aiOffice
      ?.skillsRawText?.(target.name)
      .then((text) => {
        if (!alive) return
        const doc = parseSkillDoc(text)
        if (!doc) throw new Error(t('intgSkillEditorReadFailed'))
        // the name is the folder the listing showed, not the frontmatter's own
        setDraft({ name: target.name, description: doc.description, body: doc.body })
      })
      .catch((e: unknown) => {
        if (alive) setError(e instanceof Error ? e.message : String(e))
      })
    return () => {
      alive = false
    }
    // t is a fresh closure per parent render; the load must run once per
    // target, not again on every re-render
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [target])

  const save = async (): Promise<void> => {
    if (!draft || busy) return
    setBusy(true)
    try {
      const saved = await window.aiOffice?.skillsSave?.(
        draft.name,
        draft.description,
        draft.body,
        target !== null,
      )
      if (saved) onSaved(saved)
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e))
    } finally {
      setBusy(false)
    }
  }

  const errors = draftFieldErrors(draft ?? { name: '', description: '', body: '' })

  return (
    <div className="set-intg-form">
      <div className="set-intg-head">
        <div className="set-field-label">
          {t(target ? 'intgSkillEditorTitleEdit' : 'intgSkillEditorTitleNew')}
        </div>
        <div className="set-intg-actions">
          <button className="set-btn" onClick={onClose}>
            {t('cancel')}
          </button>
        </div>
      </div>
      {error && <div className="set-field-desc set-intg-warn">{error}</div>}
      {draft ? (
        <div className="set-intg-editor-wrap">
          <div className="set-intg-editor">
            <div className="set-intg-editor-main">
              <SkillDocFields
                t={t}
                draft={draft}
                onChange={setDraft}
                nameLocked={target !== null}
              />
              <div className="set-intg-form-actions">
                <button
                  className="set-btn primary"
                  disabled={busy || errors.name || errors.description}
                  onClick={() => void save()}
                >
                  {t('intgSkillEditorSave')}
                </button>
              </div>
            </div>
            <div className="set-intg-editor-side">
              <SkillDocPreview t={t} draft={draft} />
            </div>
          </div>
        </div>
      ) : (
        !error && <div className="set-field-desc set-intg-lead">{t('intgSkillEditorLoading')}</div>
      )}
    </div>
  )
}

/** the products the foreign importer lists, in a fixed display order */
const FOREIGN_SOURCES: readonly ForeignSkillSource[] = [
  'claude',
  'codex',
  'cursor',
  'windsurf',
  'opencode',
  'pi',
]

/** the per-product headings: brand names, identical in every locale */
const FOREIGN_SOURCE_LABELS: Record<ForeignSkillSource, string> = {
  claude: 'Claude Code',
  codex: 'Codex',
  cursor: 'Cursor',
  windsurf: 'Windsurf',
  opencode: 'opencode',
  pi: 'Pi',
}

/**
 * The importer for prompt files other agent products keep on this machine: a
 * checklist grouped by product, one import per checked row. Failures are
 * reported per-row and stay checked so a retry is one click; a clean run
 * closes the view onto the refreshed My skills list.
 */
function ForeignSkillImport({
  t,
  onClose,
  onRefresh,
  onDone,
}: {
  t: TFunc
  onClose: () => void
  onRefresh: () => Promise<void>
  /** a clean run finished: count skills landed, the pane returns to the list */
  onDone: (count: number) => void
}) {
  const [candidates, setCandidates] = useState<ForeignSkillCandidate[] | null>(null)
  const [checked, setChecked] = useState<ReadonlySet<string>>(new Set())
  const [rowFailures, setRowFailures] = useState<Readonly<Record<string, string>>>({})
  const [error, setError] = useState('')
  const [busy, setBusy] = useState(false)

  useEffect(() => {
    let alive = true
    // a preload from before this feature answers with nothing, which reads
    // the same as a machine with no other products installed
    void (async () => {
      try {
        const found = await window.aiOffice?.skillsForeignList?.()
        if (alive) setCandidates(found ?? [])
      } catch (e) {
        if (alive) setError(e instanceof Error ? e.message : String(e))
      }
    })()
    return () => {
      alive = false
    }
  }, [])

  const toggle = (path: string, on: boolean): void => {
    setChecked((prev) => {
      const next = new Set(prev)
      if (on) next.add(path)
      else next.delete(path)
      return next
    })
  }

  const importSelected = async (): Promise<void> => {
    if (busy || !candidates || checked.size === 0) return
    setBusy(true)
    const failures: Record<string, string> = {}
    let landed = 0
    // one at a time: the imports are folder writes, and a failure in one row
    // must not take the rest of the batch down with it
    for (const item of candidates) {
      if (item.alreadyImported || !checked.has(item.sourcePath)) continue
      try {
        await window.aiOffice?.skillsImportForeign?.(item.sourcePath)
        landed += 1
      } catch (e) {
        failures[item.sourcePath] = e instanceof Error ? e.message : String(e)
      }
    }
    await onRefresh().catch(() => undefined)
    setRowFailures(failures)
    setChecked(new Set(Object.keys(failures)))
    setCandidates((prev) =>
      (prev ?? []).map((c) =>
        checked.has(c.sourcePath) && !failures[c.sourcePath] ? { ...c, alreadyImported: true } : c,
      ),
    )
    setBusy(false)
    if (Object.keys(failures).length === 0) onDone(landed)
  }

  const groups = FOREIGN_SOURCES.map((source) => ({
    source,
    items: (candidates ?? []).filter((c) => c.source === source),
  })).filter((g) => g.items.length > 0)

  return (
    <div className="set-intg-form">
      <div className="set-intg-head">
        <div className="set-field-label">{t('intgSkillForeignTitle')}</div>
        <div className="set-intg-actions">
          <button className="set-btn" onClick={onClose}>
            {t('cancel')}
          </button>
        </div>
      </div>
      <div className="set-field-desc set-intg-lead">{t('intgSkillForeignDesc')}</div>
      {error && <div className="set-field-desc set-intg-warn">{error}</div>}
      {candidates === null && !error && (
        <div className="set-field-desc set-intg-lead">{t('intgSkillForeignLoading')}</div>
      )}
      {candidates !== null && candidates.length === 0 && !error && (
        <div className="set-field-desc set-intg-lead">{t('intgSkillForeignEmpty')}</div>
      )}
      {groups.map((group) => (
        <div className="set-intg-foreign-group" key={group.source}>
          <div className="set-intg-foreign-source">{FOREIGN_SOURCE_LABELS[group.source]}</div>
          {group.items.map((item) => (
            <label
              key={item.sourcePath}
              className={
                item.alreadyImported ? 'set-intg-foreign-row imported' : 'set-intg-foreign-row'
              }
            >
              <input
                type="checkbox"
                className="set-intg-foreign-check"
                checked={checked.has(item.sourcePath)}
                disabled={item.alreadyImported || busy}
                onChange={(e) => toggle(item.sourcePath, e.currentTarget.checked)}
              />
              <span className="set-intg-foreign-text">
                <span className="set-intg-foreign-name">{item.name}</span>
                <span className="set-field-desc">{item.description}</span>
                {rowFailures[item.sourcePath] && (
                  <span className="set-intg-foreign-fail">
                    {t('intgSkillsImportFailed', { reason: rowFailures[item.sourcePath]! })}
                  </span>
                )}
              </span>
              {item.alreadyImported && (
                <span className="set-intg-state">{t('intgSkillForeignImportedTag')}</span>
              )}
            </label>
          ))}
        </div>
      ))}
      {candidates !== null && candidates.length > 0 && (
        <div className="set-intg-form-actions">
          <button
            className="set-btn primary"
            disabled={busy || checked.size === 0}
            onClick={() => void importSelected()}
          >
            {busy
              ? t('intgSkillForeignImporting')
              : t('intgSkillForeignImportSelected', { n: checked.size })}
          </button>
        </div>
      )}
    </div>
  )
}

const CATEGORY_ORDER = [
  'office',
  'business',
  'data',
  'finance',
  'hr',
  'legal',
  'marketing',
  'education',
  'other',
] as const
type Category = (typeof CATEGORY_ORDER)[number]
type Section = 'mine' | 'recommended'

const CATEGORY_LABELS: Record<Category, StringKey> = {
  office: 'intgSkillsCatOffice',
  business: 'intgSkillsCatBusiness',
  data: 'intgSkillsCatData',
  finance: 'intgSkillsCatFinance',
  hr: 'intgSkillsCatHr',
  legal: 'intgSkillsCatLegal',
  marketing: 'intgSkillsCatMarketing',
  education: 'intgSkillsCatEducation',
  other: 'intgSkillsCatOther',
}

function categoryOf(s: RecommendedSkill): Category {
  return (CATEGORY_ORDER as readonly string[]).includes(s.category)
    ? (s.category as Category)
    : 'other'
}

/** One skill, with its instructions behind a disclosure fetched on open. */
function SkillRow({ t, skill, action }: { t: TFunc; skill: FoundSkill; action?: React.ReactNode }) {
  const [body, setBody] = useState<string | null>(null)
  const [reading, setReading] = useState(false)

  const show = async () => {
    if (body !== null || reading) return
    setReading(true)
    try {
      setBody((await api()?.skillBody?.(skill.path)) ?? '')
    } catch {
      setBody('')
    } finally {
      setReading(false)
    }
  }

  return (
    <div className="set-intg-row set-intg-skill">
      <div className="set-field">
        <div className="set-field-text">
          <div className="set-field-stack">
            <div className="set-field-label">{skill.name}</div>
            <div className="set-field-desc">{skill.description}</div>
          </div>
        </div>
        {action && <div className="set-intg-actions">{action}</div>}
      </div>
      <details
        className="set-intg-details"
        onToggle={(e) => {
          if ((e.currentTarget as HTMLDetailsElement).open) void show()
        }}
      >
        <summary>{t('intgSkillsRead')}</summary>
        {body === null ? null : (
          <pre className="set-intg-code set-intg-code-block">
            <code>{body}</code>
          </pre>
        )}
      </details>
    </div>
  )
}
