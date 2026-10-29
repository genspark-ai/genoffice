import { Fragment, useCallback, useEffect, useRef, useState } from 'react'
import { createIpcTransport, streamText } from '@genoffice/agent-core'
import type { AgentTransport } from '@genoffice/agent-core'
import type { AiSettings } from '@genoffice/ai-provider'
import { useDismissablePopover } from '@genoffice/ui'
import type { StringKey, TFunc } from './locale'
import { aiSetupNeeded } from './ai-setup-state'
import {
  SKILL_CREATOR_SYSTEM,
  SKILL_PARAMS,
  draftFieldErrors,
  draftFromInterview,
  insertSkillParam,
  interviewToUserMessage,
  skillDocFromReply,
  skillDocText,
  suggestSkillName,
  type SkillDraft,
  type SkillInterview,
} from './skill-draft'

/**
 * The skills pane's AI creator: a four-question interview, one model call, a
 * preview the user edits, then the same save the hand-written editor uses.
 *
 * The model is reached exactly the way the home ask panel reaches it — one
 * tool-less request over the shared `ai:stream` bridge (`streamText` on
 * `createIpcTransport`), so no new channel and no loop is involved. With no
 * model configured (or a call that fails) the interview answers are assembled
 * into the same draft deterministically, so the feature works end-to-end
 * before any AI is set up.
 */

/** the reply is a SKILL.md, not a document: a generous but bounded cap */
const MAX_REPLY_CHARS = 24_000

type Phase = 'interview' | 'generating' | 'preview'

interface Question {
  key: keyof SkillInterview
  label: StringKey
  hint: StringKey
  optional?: boolean
  rows: number
}

const QUESTIONS: Question[] = [
  {
    key: 'purpose',
    label: 'intgSkillCreatorStep1',
    hint: 'intgSkillCreatorStep1Hint',
    rows: 3,
  },
  {
    key: 'trigger',
    label: 'intgSkillCreatorStep2',
    hint: 'intgSkillCreatorStep2Hint',
    rows: 3,
  },
  {
    key: 'steps',
    label: 'intgSkillCreatorStep3',
    hint: 'intgSkillCreatorStep3Hint',
    rows: 6,
  },
  {
    key: 'example',
    label: 'intgSkillCreatorStep4',
    hint: 'intgSkillCreatorStep4Hint',
    optional: true,
    rows: 4,
  },
]

export function SkillCreator({
  t,
  onClose,
  onSaved,
}: {
  t: TFunc
  onClose: () => void
  onSaved: (skill: { name: string }) => void
}) {
  const [answers, setAnswers] = useState<SkillInterview>({
    purpose: '',
    trigger: '',
    steps: '',
    example: '',
  })
  const [step, setStep] = useState(0)
  const [phase, setPhase] = useState<Phase>('interview')
  const [draft, setDraft] = useState<SkillDraft | null>(null)
  const [note, setNote] = useState<StringKey | null>(null)
  const [noteReason, setNoteReason] = useState('')
  const [aiReady, setAiReady] = useState<boolean | null>(null)
  const [saving, setSaving] = useState(false)
  const aliveRef = useRef(true)
  const abortRef = useRef<AbortController | null>(null)

  // the same gate the ask panel's composer uses: some usable provider, or the
  // default Genspark sign-in. Unknown until both answers arrive.
  useEffect(() => {
    let alive = true
    void window.aiOffice
      .aiGetSettings()
      .then((settings) =>
        window.aiOffice.aiGskStatus().then((gsk) => {
          if (alive) setAiReady(!aiSetupNeeded(settings, gsk.loggedIn))
        }),
      )
      .catch(() => {
        if (alive) setAiReady(false)
      })
    return () => {
      alive = false
    }
  }, [])

  // a generation that outlives its pane must not touch unmounted state
  useEffect(
    () => () => {
      aliveRef.current = false
      // a closed pane also cancels its one in-flight call rather than paying for it
      abortRef.current?.abort()
    },
    [],
  )

  const close = (): void => {
    abortRef.current?.abort()
    onClose()
  }

  const fillFromTemplate = (failure: StringKey, reason = ''): void => {
    setDraft(draftFromInterview(answers))
    setNote(failure)
    setNoteReason(reason)
    setPhase('preview')
  }

  const generate = async (): Promise<void> => {
    if (phase !== 'interview' || aiReady === null) return
    if (!aiReady) {
      fillFromTemplate('intgSkillCreatorFallbackNote')
      return
    }
    setPhase('generating')
    const controller = new AbortController()
    abortRef.current = controller
    try {
      const settings = await window.aiOffice.aiGetSettings()
      const transport: AgentTransport = createIpcTransport<AiSettings>({
        onStream: (listener) => window.aiOffice.onAiStream(listener),
        start: (request) => window.aiOffice.aiStream(request),
        cancel: (requestId) => void window.aiOffice.aiStreamCancel(requestId),
        getSettings: () => settings,
        unknownErrorText: () => t('askUnknownError'),
      })
      // never throws: every failure arrives as an outcome
      const outcome = await streamText({
        transport,
        system: SKILL_CREATOR_SYSTEM,
        user: interviewToUserMessage(answers),
        signal: controller.signal,
        maxChars: MAX_REPLY_CHARS,
        extract: (raw) => ({ text: raw }),
      })
      if (!aliveRef.current) return
      if (outcome.status === 'empty') {
        fillFromTemplate('intgSkillCreatorModelFailed', outcome.error)
        return
      }
      const doc = skillDocFromReply(outcome.text)
      if (!doc) {
        fillFromTemplate('intgSkillCreatorBadReply')
        return
      }
      setDraft(doc)
      setPhase('preview')
    } catch (e) {
      if (aliveRef.current) fillFromTemplate('intgSkillCreatorModelFailed', messageOf(e))
    } finally {
      abortRef.current = null
    }
  }

  const save = async (): Promise<void> => {
    if (!draft || saving) return
    setSaving(true)
    try {
      const saved = await window.aiOffice?.skillsSave?.(draft.name, draft.description, draft.body)
      if (!aliveRef.current) return
      if (saved) onSaved(saved)
    } catch (e) {
      if (aliveRef.current) {
        setNote('intgSkillsSaveFailed')
        setNoteReason(messageOf(e))
      }
    } finally {
      if (aliveRef.current) setSaving(false)
    }
  }

  const question = QUESTIONS[step]!
  const answer = answers[question.key]
  const canAdvance = question.optional || answer.trim().length > 0
  const last = step === QUESTIONS.length - 1
  const errors = draftFieldErrors(draft ?? { name: '', description: '', body: '' })

  return (
    <div className="set-intg-form">
      <div className="set-intg-head">
        <div className="set-field-label">{t('intgSkillCreatorTitle')}</div>
        <div className="set-intg-actions">
          <button className="set-btn" onClick={close}>
            {t('cancel')}
          </button>
        </div>
      </div>

      {phase === 'interview' && (
        <>
          <div className="set-field-desc">
            {t('intgSkillCreatorStepOf', { n: step + 1, total: QUESTIONS.length })}
          </div>
          <div className="set-intg-form-field">
            <label className="set-field-label" htmlFor="skill-creator-answer">
              {t(question.label)}
            </label>
            <div className="set-field-desc">{t(question.hint)}</div>
            <textarea
              id="skill-creator-answer"
              className="set-input-multiline set-intg-form-input"
              rows={question.rows}
              value={answer}
              onChange={(e) => setAnswers((prev) => ({ ...prev, [question.key]: e.target.value }))}
            />
          </div>
          <div className="set-intg-form-actions">
            {step > 0 && (
              <button className="set-btn" onClick={() => setStep(step - 1)}>
                {t('intgSkillCreatorBack')}
              </button>
            )}
            <button
              className="set-btn primary"
              disabled={!canAdvance || aiReady === null}
              onClick={() => (last ? void generate() : setStep(step + 1))}
            >
              {last ? t('intgSkillCreatorGenerate') : t('intgSkillCreatorNext')}
            </button>
          </div>
        </>
      )}

      {phase === 'generating' && (
        <div className="set-field-desc set-intg-lead">{t('intgSkillCreatorGenerating')}</div>
      )}

      {phase === 'preview' && draft && (
        <>
          {note && (
            <div className="set-field-desc set-intg-warn">
              {t(note, noteReason ? { reason: noteReason } : undefined)}
            </div>
          )}
          <div className="set-field-desc set-intg-lead">{t('intgSkillCreatorPreviewHint')}</div>
          <SkillDocFields t={t} draft={draft} onChange={setDraft} />
          <div className="set-intg-form-actions">
            <button className="set-btn" onClick={() => close()}>
              {t('cancel')}
            </button>
            <button
              className="set-btn primary"
              disabled={saving || errors.name || errors.description}
              onClick={() => void save()}
            >
              {t('intgSkillEditorSave')}
            </button>
          </div>
        </>
      )}
    </div>
  )
}

/**
 * The three SKILL.md fields the editor and the creator's preview share, laid
 * out as a fixed label column with aligned inputs: under each control its
 * live-validation hint and, where it helps, a soft counter. A folder-safe name
 * and a non-empty description are what the main process's save guard will
 * require, so the pane says so before the click — the invalid state is styled,
 * never a surprise dialog. The body field also carries the insert-parameter
 * menu: the vocabulary the app replaces at send time, spliced into the body
 * at the textarea's caret.
 */
export function SkillDocFields({
  t,
  draft,
  onChange,
  nameLocked,
}: {
  t: TFunc
  draft: SkillDraft
  onChange: (draft: SkillDraft) => void
  /** the edit flow's name is the folder on disk; it cannot change after create */
  nameLocked?: boolean
}) {
  const errors = draftFieldErrors(draft)
  const patch = (part: Partial<SkillDraft>) => onChange({ ...draft, ...part })
  const suggestion = suggestSkillName(draft.description)
  const canSuggest = !nameLocked && draft.description.trim() !== '' && suggestion !== draft.name
  const descLength = draft.description.length
  const descOver = descLength > DESC_SOFT_LIMIT
  const bodyLines = draft.body.length === 0 ? 0 : draft.body.split('\n').length
  const [paramOpen, setParamOpen] = useState(false)
  const [caretAfterInsert, setCaretAfterInsert] = useState<number | null>(null)
  const paramRef = useRef<HTMLDivElement | null>(null)
  const bodyRef = useRef<HTMLTextAreaElement | null>(null)
  const closeParam = useCallback(() => setParamOpen(false), [])
  useDismissablePopover(paramOpen, closeParam, { inside: () => [paramRef.current] })

  // the caret lands right after the inserted token once the textarea carries
  // the new body: React writes the value first, then this effect restores the
  // selection typing would have left
  useEffect(() => {
    if (caretAfterInsert === null) return
    setCaretAfterInsert(null)
    const el = bodyRef.current
    if (!el) return
    el.focus()
    el.setSelectionRange(caretAfterInsert, caretAfterInsert)
  }, [caretAfterInsert])

  /** splice one menu token into the body at the caret the textarea was left at */
  const insertParam = (token: string): void => {
    const el = bodyRef.current
    const next = insertSkillParam(draft.body, token, el?.selectionStart ?? draft.body.length)
    patch({ body: next.body })
    setCaretAfterInsert(next.caret)
  }
  return (
    <div className="set-intg-doc">
      <div className="set-intg-doc-field">
        <label className="set-field-label" htmlFor="skill-doc-name">
          {t('intgSkillEditorName')}
        </label>
        <div className="set-intg-doc-control">
          <div className="set-intg-doc-name">
            <input
              id="skill-doc-name"
              className={
                'set-input set-intg-form-input' + (errors.name ? ' set-input-invalid' : '')
              }
              value={draft.name}
              readOnly={nameLocked}
              spellCheck={false}
              onChange={(e) => patch({ name: e.target.value })}
            />
            <button
              type="button"
              className="set-intg-suggest"
              disabled={!canSuggest}
              onClick={() => patch({ name: suggestion })}
            >
              {t('intgSkillEditorNameSuggest')}
            </button>
          </div>
          <div className="set-intg-doc-hint">
            <span className={errors.name ? 'set-intg-doc-hint-bad' : undefined}>
              {errors.name ? t('intgSkillEditorNameBad') : t('intgSkillEditorNameHint')}
            </span>
          </div>
        </div>
      </div>
      <div className="set-intg-doc-field">
        <label className="set-field-label" htmlFor="skill-doc-desc">
          {t('intgSkillEditorDesc')}
        </label>
        <div className="set-intg-doc-control">
          <input
            id="skill-doc-desc"
            className={
              'set-input set-intg-form-input' + (errors.description ? ' set-input-invalid' : '')
            }
            value={draft.description}
            onChange={(e) => patch({ description: e.target.value })}
          />
          <div className="set-intg-doc-hint">
            <span className={errors.description ? 'set-intg-doc-hint-bad' : undefined}>
              {errors.description
                ? t('intgSkillEditorDescBad')
                : descOver
                  ? t('intgSkillEditorDescLong')
                  : t('intgSkillEditorDescHint')}
            </span>
            <span className={descOver ? 'set-intg-doc-count over' : 'set-intg-doc-count'}>
              {descLength}/{DESC_SOFT_LIMIT}
            </span>
          </div>
        </div>
      </div>
      <div className="set-intg-doc-field">
        <label className="set-field-label" htmlFor="skill-doc-body">
          {t('intgSkillEditorBody')}
        </label>
        <div className="set-intg-doc-control">
          <div className="set-intg-doc-tools set-intg-import" ref={paramRef}>
            <button
              type="button"
              className="set-intg-suggest"
              aria-haspopup="menu"
              aria-expanded={paramOpen}
              onClick={() => setParamOpen((v) => !v)}
            >
              {t('intgSkillEditorInsertParam')}
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
            {paramOpen && (
              <div className="set-intg-import-menu set-intg-param-menu" role="menu">
                {SKILL_PARAMS.map((param) => (
                  <button
                    key={param.token}
                    type="button"
                    role="menuitem"
                    className="set-intg-import-item set-intg-param-item"
                    onClick={() => {
                      setParamOpen(false)
                      insertParam(param.token)
                    }}
                  >
                    <code className="set-intg-param-token">{param.token}</code>
                    <span className="set-intg-param-desc">{t(param.desc)}</span>
                  </button>
                ))}
              </div>
            )}
          </div>
          <textarea
            id="skill-doc-body"
            ref={bodyRef}
            className="set-input-multiline set-intg-form-input"
            rows={14}
            spellCheck={false}
            value={draft.body}
            onChange={(e) => patch({ body: e.target.value })}
          />
          <div className="set-intg-doc-hint">
            <span className="set-intg-doc-hint-lines">
              <span>{t('intgSkillEditorBodyHint')}</span>
              <span>{t('intgSkillEditorBodyArgsHint')}</span>
            </span>
            <span className="set-intg-doc-count">
              {t(bodyLines === 1 ? 'intgSkillEditorLineCountOne' : 'intgSkillEditorLineCount', {
                n: bodyLines,
              })}
            </span>
          </div>
        </div>
      </div>
    </div>
  )
}

/** the description's soft target: the relevance check reads one sentence */
const DESC_SOFT_LIMIT = 160

/** a frontmatter line the preview highlights: `name:` / `description:` */
const PREVIEW_KEY = /^(name|description):(.*)$/

/**
 * The read-only half of the skill editor: the exact SKILL.md Save will write,
 * live as the fields change. The text comes from skillDocText — the renderer
 * twin of the main process's writer — so the preview can never show a file
 * Save would build differently. The frontmatter block sits visually apart
 * from the body: quiet fences, emphasized keys, muted panel.
 */
export function SkillDocPreview({ t, draft }: { t: TFunc; draft: SkillDraft | null }) {
  const text = draft ? skillDocText(draft) : ''
  // the frontmatter through its closing fence; the rest is the body
  const closing = text.indexOf('\n---\n')
  const frontmatter = closing === -1 ? text : text.slice(0, closing + 5)
  const body = closing === -1 ? '' : text.slice(closing + 5)
  return (
    <div className="set-intg-preview">
      <div className="set-intg-preview-title">{t('intgSkillEditorPreviewTitle')}</div>
      <div className="set-field-desc set-intg-preview-hint">{t('intgSkillEditorPreviewHint')}</div>
      <pre className="set-intg-preview-code">
        <code>
          {frontmatter.split('\n').map((line, i) => {
            const key = PREVIEW_KEY.exec(line)
            return (
              <Fragment key={i}>
                {i > 0 ? '\n' : null}
                {key ? (
                  <>
                    <span className="set-intg-preview-key">{key[1]}:</span>
                    {key[2]}
                  </>
                ) : (
                  <span className="set-intg-preview-fence">{line}</span>
                )}
              </Fragment>
            )
          })}
          <span className="set-intg-preview-body">{body}</span>
        </code>
      </pre>
    </div>
  )
}

function messageOf(e: unknown): string {
  return e instanceof Error ? e.message : String(e)
}
