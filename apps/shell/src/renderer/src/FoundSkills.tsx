import { useCallback, useEffect, useRef, useState } from 'react'
import type { CSSProperties } from 'react'
import { useSlidingIndicator } from './settings-fx'
import type { StringKey, TFunc } from './locale'
import type { FoundSkill, RecommendedSkill } from '../../shared/found-skill'

/**
 * The skills the user installed into GenOffice.
 *
 * Only GenOffice's own folder is listed. A skill gets in one way: the user
 * picks a folder or a zip and presses Import. Nothing on this screen runs
 * anything.
 */

const api = () => window.aiOfficeIntegrations

type Notice = { kind: 'done' | 'removed'; name: string } | { kind: 'failed'; reason: string }

export function FoundSkills({ t }: { t: TFunc }) {
  const [skills, setSkills] = useState<FoundSkill[] | null>(null)
  const [recommended, setRecommended] = useState<RecommendedSkill[]>([])
  const [busy, setBusy] = useState(false)
  const [notice, setNotice] = useState<Notice | null>(null)
  const [tab, setTab] = useState<Category | 'all' | null>(null)
  const [section, setSection] = useState<Section | null>(null)
  const tabsBox = useRef<HTMLDivElement | null>(null)
  const tabsPill = useSlidingIndicator(tabsBox, '.set-intg-seg.active', section)

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

  const categories = CATEGORY_ORDER.filter((c) => recommended.some((s) => categoryOf(s) === c))
  // 'all' sits ahead of the real categories and is the landing view: the point
  // of the catalog is to browse everything, not to guess a category first
  const current: Category | 'all' =
    tab === 'all' || (tab !== null && categories.includes(tab)) ? (tab as Category | 'all') : 'all'
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

      <div className="set-intg-head">
        {categories.length > 0 ? (
          <div ref={tabsBox} className="set-intg-tabs set-intg-sections" role="tablist">
            <span ref={tabsPill} className="fx-indicator" aria-hidden="true" />
            {sectionTab('mine', `${t('intgSkillsMine')} (${skills.length})`)}
            {sectionTab('recommended', t('intgSkillsRecommended'))}
          </div>
        ) : (
          <div className="set-field-label">{t('intgSkillsMine')}</div>
        )}
        <div className="set-intg-actions">
          <button
            className="set-btn"
            disabled={busy}
            onClick={() =>
              void doImport(
                () =>
                  api()?.importSkillZip?.(t('intgSkillsImportZipTitle')) ?? Promise.resolve(null),
              )
            }
          >
            {t('intgSkillsImportZip')}
          </button>
          <button
            className="set-btn"
            disabled={busy}
            onClick={() =>
              void doImport(
                () =>
                  api()?.importSkillFolder?.(t('intgSkillsImportFolderTitle')) ??
                  Promise.resolve(null),
              )
            }
          >
            {t('intgSkillsImportFolder')}
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
            ? t('intgSkillsImportFailed', { reason: notice.reason })
            : notice.kind === 'done'
              ? t('intgSkillsImportDone', { name: notice.name })
              : t('intgSkillsRemoved', { name: notice.name })}
        </div>
      )}

      {shown === 'mine' && skills.length === 0 && (
        <div className="set-field-desc set-intg-lead">{t('intgSkillsEmpty')}</div>
      )}

      {shown === 'mine' && skills.length > 0 && (
        <div className="set-skill-grid">
          {skills.map((skill, i) => (
            <SkillRow
              key={skill.path}
              index={i}
              t={t}
              skill={skill}
              action={
                <button className="set-btn" disabled={busy} onClick={() => void doRemove(skill)}>
                  {t('intgSkillsRemove')}
                </button>
              }
            />
          ))}
        </div>
      )}

      {shown === 'recommended' && categories.length > 0 && (
        <div className="set-skill-browse">
          <div
            className="set-intg-tabs set-skill-chips"
            role="tablist"
            aria-label={t('intgSkillsRecommended')}
          >
            {(['all', ...categories] as const).map((c) => (
              <button
                key={c}
                type="button"
                role="tab"
                aria-selected={c === current}
                className={c === current ? 'set-intg-tab active' : 'set-intg-tab'}
                onClick={() => setTab(c)}
              >
                {c === 'all' ? t('intgSkillsCatAll') : t(CATEGORY_LABELS[c])}
              </button>
            ))}
          </div>
          <div role="tabpanel" className="set-skill-grid">
            {(current === 'all'
              ? recommended
              : recommended.filter((s) => categoryOf(s) === current)
            ).map((skill, i) => (
              <SkillRow
                key={skill.path}
                index={i}
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
                      {skill.state === 'outdated' ? t('intgSkillsUpdate') : t('intgSkillsInstall')}
                    </button>
                  )
                }
              />
            ))}
          </div>
        </div>
      )}
    </>
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
function SkillRow({
  t,
  skill,
  action,
  index = 0,
}: {
  t: TFunc
  skill: FoundSkill
  action?: React.ReactNode
  index?: number
}) {
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
    <div className="set-skill-card" style={{ '--i': index } as CSSProperties}>
      <div className="set-skill-head">
        <span className="set-skill-icon" aria-hidden="true">
          {skill.name.slice(0, 1).toUpperCase()}
        </span>
        <div className="set-skill-text">
          <div className="set-skill-name">{skill.name}</div>
          <div className="set-skill-desc">{skill.description}</div>
        </div>
      </div>
      <div className="set-skill-foot">
        <details
          className="set-skill-details"
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
        {action && <div className="set-skill-action">{action}</div>}
      </div>
    </div>
  )
}
