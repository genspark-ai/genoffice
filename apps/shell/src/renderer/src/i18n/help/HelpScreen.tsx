import { useEffect, useMemo, useRef, useState } from 'react'
import type { KeyboardEvent as ReactKeyboardEvent } from 'react'
import { Markdown } from '@genoffice/ui'
import { HELP_GROUPS, HELP_TOPICS, helpBody, helpImage, searchTopics } from './help-registry'
import { useI18n } from '../../locale'
import { groupTitle, topicTitle } from './help-titles'

/**
 * The in-app manual (issue #1520): sidebar topic list with full-text search
 * over a markdown body per topic. Rendered by the shell renderer when the
 * view is loaded at ?mode=help (the Help tab / F1 / Help menu).
 *
 * Topics ship in zh + en; every other locale falls back to en (the registry
 * makes the gap visible rather than hiding it).
 *
 * Interactive bits: `[label](help://topic-id)` links in topic bodies jump to
 * another topic (the shared Markdown renderer's nav hook); ArrowUp/ArrowDown
 * walk the sidebar, `/` focuses search; the in-page table of contents is built
 * from the body's `##` headings and scrolls to them by index.
 */
export function HelpScreen(): React.ReactElement {
  const { t } = useI18n()
  // the help tab has no locale of its own: it renders in the shell's UI language
  const langTag = document.documentElement.lang
  const zh = langTag.startsWith('zh')
  const [query, setQuery] = useState('')
  const [activeId, setActiveId] = useState(HELP_TOPICS[0]!.id)
  const hits = useMemo(() => searchTopics(query, zh ? 'zh' : 'en'), [query, zh])
  const searchRef = useRef<HTMLInputElement>(null)
  const mainRef = useRef<HTMLElement>(null)

  const active = HELP_TOPICS.find((t) => t.id === activeId) ?? null
  const raw = active ? helpBody(active.id, zh ? 'zh' : 'en') : null
  // the article header already renders the title; drop the body's own `# ` line
  const body = raw === null ? null : raw.replace(/^#\s+[^\n]*\n+/, '')

  /** topics in sidebar order that match the current query */
  const visibleTopics = useMemo(
    () =>
      HELP_TOPICS.filter((t) => hits.has(t.id)).sort(
        (a, b) =>
          HELP_GROUPS.findIndex((g) => g.id === a.group) -
            HELP_GROUPS.findIndex((g) => g.id === b.group) || a.order - b.order,
      ),
    [hits],
  )

  /** `## heading` lines of the active body, in order */
  const toc = useMemo(() => {
    if (body === null) return [] as string[]
    const out: string[] = []
    for (const line of body.split('\n')) {
      const m = /^## (.+)$/.exec(line)
      if (m) out.push(m[1]!.trim())
    }
    return out
  }, [body])

  // reset scroll when the topic changes
  useEffect(() => {
    mainRef.current?.scrollTo({ top: 0 })
  }, [activeId])

  const gotoTopic = (id: string) => {
    if (HELP_TOPICS.some((t) => t.id === id)) setActiveId(id)
  }

  const stepTopic = (dir: 1 | -1) => {
    const i = visibleTopics.findIndex((t) => t.id === activeId)
    const next = visibleTopics[Math.min(Math.max(i + dir, 0), visibleTopics.length - 1)]
    if (next && next.id !== activeId) setActiveId(next.id)
  }

  const onKeyDown = (e: ReactKeyboardEvent<HTMLDivElement>) => {
    // ignore keys typed into the search box itself (except Escape)
    if (e.target instanceof HTMLInputElement) {
      if (e.key === 'Escape') {
        setQuery('')
        e.currentTarget.focus()
      }
      return
    }
    if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
      e.preventDefault()
      stepTopic(e.key === 'ArrowDown' ? 1 : -1)
    } else if (e.key === '/') {
      e.preventDefault()
      searchRef.current?.focus()
    }
  }

  const scrollToHeading = (index: number) => {
    // the TOC is built from `## ` lines; the Markdown renderer emits them as
    // .ai-md-h2 (it renders all heading levels as <p class="ai-md-h…">)
    const headings = mainRef.current?.querySelectorAll('.help-article .ai-md-h2')
    headings?.[index]?.scrollIntoView({ behavior: 'smooth', block: 'start' })
  }

  return (
    <div className="help-root" onKeyDown={onKeyDown} tabIndex={-1}>
      <aside className="help-side">
        <div className="help-search-row">
          <input
            ref={searchRef}
            className="help-search"
            type="search"
            placeholder={t('helpSearchPlaceholder')}
            value={query}
            aria-label={t('helpSearchLabel')}
            onChange={(e) => setQuery(e.target.value)}
          />
          {query && (
            <span className="help-hit-count">
              {t(visibleTopics.length === 1 ? 'helpHitCountOne' : 'helpHitCount', {
                n: visibleTopics.length,
              })}
            </span>
          )}
        </div>
        <nav className="help-nav" aria-label={t('helpContentsLabel')}>
          {HELP_GROUPS.map((g) => {
            const topics = visibleTopics.filter((t) => t.group === g.id)
            if (topics.length === 0) return null
            return (
              <div key={g.id} className="help-group">
                <div className="help-group-title">{groupTitle(g.id, langTag)}</div>
                {topics.map((t) => (
                  <button
                    key={t.id}
                    type="button"
                    className={`help-topic${t.id === activeId ? ' active' : ''}`}
                    onClick={() => setActiveId(t.id)}
                  >
                    {topicTitle(t.id, langTag)}
                  </button>
                ))}
              </div>
            )
          })}
        </nav>
      </aside>
      <main className="help-main" ref={mainRef}>
        {active ? (
          <article className="help-article" key={active.id}>
            <h1>{topicTitle(active.id, langTag)}</h1>
            {toc.length > 1 && (
              <nav className="help-toc" aria-label={t('helpOnThisPage')}>
                <div className="help-toc-title">{t('helpOnThisPage')}</div>
                {toc.map((h, i) => (
                  <button
                    key={i}
                    type="button"
                    className="help-toc-item"
                    onClick={() => scrollToHeading(i)}
                  >
                    {h}
                  </button>
                ))}
              </nav>
            )}
            {body !== null ? (
              <Markdown
                text={body}
                images={{ resolve: helpImage }}
                nav={{
                  scheme: 'help://',
                  onNavigate: (href) => {
                    const id = href.replace(/^help:\/\//, '').replace(/\/$/, '')
                    gotoTopic(id)
                  },
                }}
              />
            ) : (
              <p className="help-missing">{t('helpTopicMissing')}</p>
            )}
          </article>
        ) : null}
      </main>
    </div>
  )
}
