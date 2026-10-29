import React, { useEffect, useId, useLayoutEffect, useMemo, useRef, useState } from 'react'
import type { KeyboardEvent, RefObject } from 'react'
import type { Lang } from '@genoffice/i18n'
import type { AiSkillsMenuEntry } from './AiSkillsMenu'
import { detectSlashQuery, filterSlashSkills } from './ai-slash'
import { AI_SLASH_LABELS } from './strings-ai-slash'
import { useDismissablePopover } from './popover-dismiss'
import './ai-slash.css'

export interface UseAiSlashOptions {
  readonly lang: Lang
  /** called on every open so a fresh install shows up without a reload */
  readonly listSkills: () => Promise<readonly AiSkillsMenuEntry[]>
  /** the composer textarea; the palette anchors above it and tracks its caret */
  readonly inputRef: RefObject<HTMLTextAreaElement | null>
  /** current textarea text — the parent keeps owning the draft state */
  readonly value: string
  /** the token range handed over covers the whole `/query`, slash included */
  readonly onAttachSkill: (name: string, range: { start: number; end: number }) => void
}

export interface AiSlash {
  /** the palette, to render next to the composer; null while closed */
  readonly element: React.JSX.Element | null
  /** call first in the textarea's onKeyDown; true when the palette ate the key */
  readonly onKeyDown: (event: KeyboardEvent<HTMLTextAreaElement>) => boolean
}

/**
 * Slash palette for the AI composer: typing `/` at the start of the draft (or
 * after whitespace) opens a skill picker pinned above the textarea, the same
 * skill list the "+" menu shows; picking one hands the parent the token range
 * to strike from the draft and the skill name to attach. The parent keeps
 * owning the textarea state — this hook only reads the text and the caret, and
 * never edits the draft itself.
 *
 * Feed `onKeyDown` to the textarea before any send/stop handling: it returns
 * true while the palette consumed the key (↑/↓, Enter/Tab pick, Escape), so a
 * pick can never send the half-typed token. A dismissal — Escape, an outside
 * press, a window switch — stays dismissed until the slash token is destroyed
 * and typed again.
 */
export function useAiSlash({
  lang,
  listSkills,
  inputRef,
  value,
  onAttachSkill,
}: UseAiSlashOptions): AiSlash {
  const strings = AI_SLASH_LABELS[lang]
  const [caret, setCaret] = useState(0)
  const [skills, setSkills] = useState<readonly AiSkillsMenuEntry[] | null>(null)
  const [active, setActive] = useState(0)
  const [popStyle, setPopStyle] = useState<React.CSSProperties>({})
  const [suppressedStart, setSuppressedStart] = useState<number | null>(null)
  const popRef = useRef<HTMLDivElement | null>(null)
  const listSkillsRef = useRef(listSkills)
  listSkillsRef.current = listSkills

  const token = useMemo(() => detectSlashQuery(value, caret), [value, caret])
  const open = token !== null && token.start !== suppressedStart
  const query = token?.query ?? ''

  // a dismissal holds while its token lives: the palette comes back only when
  // the token disappears (caret leaves, slash deleted) and is typed anew
  const tokenStart = token?.start ?? null
  useEffect(() => {
    setSuppressedStart((s) => (s !== null && s !== tokenStart ? null : s))
  }, [tokenStart])

  // the caret is read after each draft change (typing) and on the click/keyup
  // that move it without changing the text
  useEffect(() => {
    const ta = inputRef.current
    setCaret(ta ? (ta.selectionStart ?? value.length) : value.length)
  }, [value, inputRef])

  useEffect(() => {
    const ta = inputRef.current
    if (!ta) return
    const read = () => setCaret(ta.selectionStart ?? 0)
    ta.addEventListener('click', read)
    ta.addEventListener('keyup', read)
    return () => {
      ta.removeEventListener('click', read)
      ta.removeEventListener('keyup', read)
    }
  }, [inputRef])

  useEffect(() => {
    if (!open) return
    setSkills(null)
    let alive = true
    listSkillsRef
      .current()
      .then((list) => {
        if (alive) setSkills(list)
      })
      .catch(() => {
        if (alive) setSkills([])
      })
    return () => {
      alive = false
    }
  }, [open])

  // the composer sits at the panel bottom: open upwards, pinned to the textarea
  // (and re-pinned as the textarea auto-grows with the draft)
  useLayoutEffect(() => {
    if (!open) return
    const rect = inputRef.current?.getBoundingClientRect()
    if (rect) setPopStyle({ left: rect.left, bottom: window.innerHeight - rect.top + 6 })
  }, [open, value, inputRef])

  const close = () => setSuppressedStart(token ? token.start : null)
  useDismissablePopover(open, close, { inside: () => [popRef.current, inputRef.current] })

  const rows = useMemo(() => (skills ? filterSlashSkills(skills, query) : []), [skills, query])
  const activeIndex = rows.length === 0 ? 0 : Math.min(active, rows.length - 1)
  useEffect(() => {
    setActive(0)
  }, [open, query])

  const pick = (entry: AiSkillsMenuEntry) => {
    if (!token) return
    setSuppressedStart(token.start)
    onAttachSkill(entry.name, { start: token.start, end: token.end })
  }

  const onKeyDown = (event: KeyboardEvent<HTMLTextAreaElement>): boolean => {
    if (!open || event.nativeEvent.isComposing) return false
    const count = rows.length
    if (event.key === 'ArrowUp' || event.key === 'ArrowDown') {
      if (count === 0) return false
      event.preventDefault()
      const step = event.key === 'ArrowUp' ? count - 1 : 1
      setActive((i) => (Math.min(i, count - 1) + step) % count)
      return true
    }
    if (event.key === 'Enter' || event.key === 'Tab') {
      const row = rows[activeIndex]
      if (!row) return false
      event.preventDefault()
      pick(row)
      return true
    }
    if (event.key === 'Escape') {
      event.preventDefault()
      close()
      return true
    }
    return false
  }

  const uid = useId()
  const element = open ? (
    <div
      ref={popRef}
      className="ai-slash-pop"
      role="listbox"
      aria-label={strings.title}
      aria-activedescendant={rows.length > 0 ? `${uid}-opt-${activeIndex}` : undefined}
      style={popStyle}
    >
      <div className="ai-slash-list">
        {skills && rows.length === 0 && (
          <div className="ai-slash-empty">
            {skills.length > 0 ? strings.noMatch : strings.noSkills}
          </div>
        )}
        {rows.map((s, i) => (
          <button
            key={s.name}
            type="button"
            id={`${uid}-opt-${i}`}
            role="option"
            aria-selected={i === activeIndex}
            className={`ai-slash-skill${i === activeIndex ? ' active' : ''}`}
            onClick={() => pick(s)}
            onMouseDown={(e) => e.preventDefault()}
            title={s.description}
          >
            <span className="ai-slash-avatar" aria-hidden>
              {s.name.slice(0, 1).toUpperCase()}
            </span>
            <span className="ai-slash-skill-text">
              <span className="ai-slash-skill-name">{s.name}</span>
              <span className="ai-slash-skill-desc">{s.description}</span>
            </span>
          </button>
        ))}
      </div>
    </div>
  ) : null

  return { element, onKeyDown }
}
