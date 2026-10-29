import { useCallback, useEffect, useMemo, useState } from 'react'
import type { Lang } from '@genoffice/i18n'
import { FileBadge } from './file-badge'
import {
  TEMPLATE_CATEGORIES,
  TEMPLATE_GALLERY_STRINGS,
  type TemplateCategory,
} from './template-gallery-strings'
import type { TemplateInfo } from '../../shared/home-api'

export interface TemplateGalleryProps {
  lang: Lang
  onClose: () => void
}

/** the file-type badge + kind tag both show the extension: language-neutral */
const KIND_LABEL: Record<TemplateInfo['kind'], string> = {
  docx: 'DOCX',
  xlsx: 'XLSX',
  pptx: 'PPTX',
  md: 'MD',
}

interface CategoryGroup {
  category: TemplateCategory
  templates: TemplateInfo[]
}

/** group by category in the skills pane's order; an unknown category sorts last */
function groupTemplates(templates: readonly TemplateInfo[]): CategoryGroup[] {
  const byCategory = new Map<TemplateCategory, TemplateInfo[]>()
  const other: TemplateInfo[] = []
  for (const template of templates) {
    if ((TEMPLATE_CATEGORIES as readonly string[]).includes(template.category)) {
      const key = template.category as TemplateCategory
      const bucket = byCategory.get(key) ?? []
      bucket.push(template)
      byCategory.set(key, bucket)
    } else {
      other.push(template)
    }
  }
  const groups = TEMPLATE_CATEGORIES.filter((category) => byCategory.has(category)).map(
    (category) => ({ category, templates: byCategory.get(category)! }),
  )
  if (other.length > 0) groups.push({ category: 'other', templates: other })
  return groups
}

/** one template card: file badge, title, description, kind tag */
function TemplateCard({
  template,
  busy,
  onClick,
}: {
  template: TemplateInfo
  busy: boolean
  onClick: () => void
}) {
  return (
    <button
      type="button"
      className="tg-card"
      disabled={busy}
      aria-label={template.title}
      onClick={onClick}
    >
      <span className="tg-card-head">
        <FileBadge ext={template.kind} size={40} />
        <span className="tg-kind">{KIND_LABEL[template.kind]}</span>
      </span>
      <span className="tg-card-title">{template.title}</span>
      <span className="tg-card-desc">{template.description}</span>
    </button>
  )
}

/**
 * The template center dialog, mounted by Home. Loading, empty and error
 * states included; picking a template copies the shipped file into the user's
 * default save folder (main process) and opens it in the matching app tab,
 * after which the gallery closes. Its strings live in
 * template-gallery-strings.ts, not the shell's shared dictionary.
 */
export function TemplateGallery({ lang, onClose }: TemplateGalleryProps) {
  const strings = TEMPLATE_GALLERY_STRINGS[lang] ?? TEMPLATE_GALLERY_STRINGS.en
  const [templates, setTemplates] = useState<TemplateInfo[] | null>(null)
  const [loadFailed, setLoadFailed] = useState(false)
  const [creatingId, setCreatingId] = useState<string | null>(null)
  const [createFailed, setCreateFailed] = useState(false)

  const load = useCallback(() => {
    setTemplates(null)
    setLoadFailed(false)
    // a rejected invoke must surface as the error state, not an unhandled
    // rejection — the same policy the settings modal applies to its probes
    window.aiOffice.templatesList().then(
      (list) => setTemplates(list),
      () => setLoadFailed(true),
    )
  }, [])

  useEffect(() => {
    load()
  }, [load])

  useEffect(() => {
    const handler = (e: KeyboardEvent) => {
      if (e.key === 'Escape') onClose()
    }
    window.addEventListener('keydown', handler)
    return () => window.removeEventListener('keydown', handler)
  }, [onClose])

  const groups = useMemo(() => groupTemplates(templates ?? []), [templates])

  const createFromTemplate = useCallback(
    (id: string) => {
      if (creatingId !== null) return
      setCreatingId(id)
      setCreateFailed(false)
      window.aiOffice.newFromTemplate(id).then(
        () => onClose(),
        () => {
          setCreateFailed(true)
          setCreatingId(null)
        },
      )
    },
    [creatingId, onClose],
  )

  return (
    <div
      className="set-overlay"
      onMouseDown={(e) => {
        if (e.target === e.currentTarget) onClose()
      }}
    >
      <div
        className="set-dialog tg-dialog"
        role="dialog"
        aria-modal="true"
        aria-label={strings.title}
      >
        <div className="set-header">
          <div>
            <h2 className="set-title">{strings.title}</h2>
            <p className="tg-subtitle">{strings.subtitle}</p>
          </div>
          <button className="set-close" onClick={onClose} aria-label={strings.close}>
            <svg width="14" height="14" viewBox="0 0 14 14" aria-hidden="true">
              <path
                d="M2 2l10 10M12 2L2 12"
                stroke="currentColor"
                strokeWidth="1.5"
                strokeLinecap="round"
              />
            </svg>
          </button>
        </div>
        <div className="tg-body">
          {createFailed && <div className="tg-error">{strings.createFailed}</div>}
          {loadFailed ? (
            <div className="tg-state">
              <p className="tg-state-text">{strings.error}</p>
              <button type="button" className="tg-state-action" onClick={load}>
                {strings.retry}
              </button>
            </div>
          ) : templates === null ? (
            <div className="tg-state">
              <p className="tg-state-text">{strings.loading}</p>
            </div>
          ) : templates.length === 0 ? (
            <div className="tg-state">
              <p className="tg-state-text">{strings.empty}</p>
            </div>
          ) : (
            groups.map((group) => (
              <section key={group.category} className="tg-group">
                <h3 className="tg-group-title">{strings.categories[group.category]}</h3>
                <div className="tg-grid">
                  {group.templates.map((template) => (
                    <TemplateCard
                      key={template.id}
                      template={template}
                      busy={creatingId !== null}
                      onClick={() => createFromTemplate(template.id)}
                    />
                  ))}
                </div>
              </section>
            ))
          )}
        </div>
      </div>
    </div>
  )
}
