import { useCallback, useEffect, useState } from 'react'
import type { TFunc } from './locale'
import type { FoundSkill } from '../../shared/found-skill'

/**
 * The skills this machine already has.
 *
 * Two lists, and the difference is the whole point. What sits in GenOffice's own
 * folder is ours to use. What sits in a coding agent's folder is offered as a
 * candidate — a developer's `.claude` directory is full of git and deploy
 * skills that have nothing to do with word processing, so scanning them shows
 * what matched and asks before taking anything.
 *
 * Pressing Import copies the skill folder across. Nothing is copied just
 * because it was found, and nothing on this screen runs anything.
 */

const api = () => window.aiOfficeIntegrations

export function FoundSkills({ t }: { t: TFunc }) {
  const [skills, setSkills] = useState<FoundSkill[] | null>(null)
  const [busy, setBusy] = useState<string | null>(null)
  const [notice, setNotice] = useState<{ name: string; ok: boolean } | null>(null)

  const load = useCallback(async () => {
    // the method itself is checked, not just the bridge: a preload from before
    // this feature was added goes on answering with a partial object
    const found = await api()?.listSkills?.()
    if (found) setSkills(found)
  }, [])

  useEffect(() => {
    void load()
  }, [load])

  const doImport = async (skill: FoundSkill) => {
    if (busy) return
    setBusy(skill.path)
    try {
      await api()?.importSkill?.(skill.path)
      setNotice({ name: skill.name, ok: true })
      await load()
    } catch {
      setNotice({ name: skill.name, ok: false })
    } finally {
      setBusy(null)
    }
  }

  // nothing rendered until the scan answers: an empty section that fills in a
  // moment later reads as "none found", which is the one thing it must not say
  if (!skills) return null

  const installed = skills.filter((s) => s.source === 'genoffice')
  const candidates = skills.filter((s) => s.source === 'agent')

  return (
    <section className="set-intg-part">
      <h5 className="set-intg-sub">{t('intgSkillsTitle')}</h5>
      <div className="set-field-desc set-intg-lead">{t('intgSkillsDesc')}</div>

      {skills.length === 0 && (
        <div className="set-field-desc set-intg-lead">{t('intgSkillsEmpty')}</div>
      )}

      {installed.length > 0 && (
        <div className="set-intg-list">
          <div className="set-field-label">{t('intgSkillsInstalled')}</div>
          {installed.map((skill) => (
            <SkillRow key={skill.path} t={t} skill={skill} />
          ))}
        </div>
      )}

      {candidates.length > 0 && (
        <div className="set-intg-list">
          <div className="set-field-label">{t('intgSkillsCandidates')}</div>
          {candidates.map((skill) => (
            <SkillRow
              key={skill.path}
              t={t}
              skill={skill}
              action={
                <button
                  className="set-btn"
                  disabled={busy === skill.path}
                  onClick={() => void doImport(skill)}
                >
                  {t('intgSkillsImport')}
                </button>
              }
            />
          ))}
        </div>
      )}

      {notice && (
        <div
          className={notice.ok ? 'set-field-desc set-intg-notice' : 'set-field-desc set-intg-warn'}
        >
          {notice.ok
            ? t('intgSkillsImportDone', { name: notice.name })
            : t('intgSkillsImportFailed', { name: notice.name })}
        </div>
      )}
    </section>
  )
}

/**
 * One skill, with its instructions behind a disclosure.
 *
 * Reading them before importing is the point of showing them at all: it is the
 * difference between a skill that fills in a document and one that quietly
 * reorganises a folder. The body is fetched when the disclosure opens rather
 * than up front — a developer with fifty skills in `.claude` should not pay for
 * reading all fifty to look at the list.
 */
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
    <div className="set-intg-row">
      <div className="set-field">
        <div className="set-field-text">
          <div className="set-field-stack">
            <div className="set-field-label">
              {skill.name}
              {skill.agent ? ` · ${skill.agent}` : ''}
            </div>
            <div className="set-field-desc">{skill.description}</div>
            {!skill.relevance.relevant && (
              <div className="set-field-desc set-intg-warn">{t('intgSkillsOtherFormats')}</div>
            )}
          </div>
        </div>
        {action && <div className="set-intg-actions">{action}</div>}
      </div>
      <details
        className="set-intg-details set-intg-cli"
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
