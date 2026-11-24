import { useState } from 'react'
import { getMarkRange } from '@tiptap/core'
import type { Editor } from '@tiptap/core'
import { useI18n, type StringKey } from '../i18n/locale'
import { useModalKeys } from './modal-keys'
import {
  TOP_ANCHOR,
  buildMailto,
  collectLinkTargets,
  ensureHeadingAnchor,
  linkKindOf,
  parseMailto,
  type LinkKind,
} from '../editor/link-targets'

const TABS: Array<{ key: LinkKind; nameKey: StringKey }> = [
  { key: 'web', nameKey: 'ribbonLinkTabWeb' },
  { key: 'doc', nameKey: 'ribbonLinkTabDoc' },
  { key: 'email', nameKey: 'ribbonLinkTabEmail' },
]

/** a place in this document: its top, a heading (by position) or an anchor name */
type Place =
  { kind: 'top' } | { kind: 'heading'; pos: number } | { kind: 'anchor'; name: string } | null

const RECENT_EMAILS_KEY = 'docs.link.recentEmails'
const RECENT_EMAILS_MAX = 8

function recentEmails(): string[] {
  try {
    const parsed: unknown = JSON.parse(localStorage.getItem(RECENT_EMAILS_KEY) ?? '[]')
    return Array.isArray(parsed) ? parsed.filter((v): v is string => typeof v === 'string') : []
  } catch {
    return []
  }
}

function rememberEmail(address: string): void {
  try {
    const next = [address, ...recentEmails().filter((a) => a !== address)]
    localStorage.setItem(RECENT_EMAILS_KEY, JSON.stringify(next.slice(0, RECENT_EMAILS_MAX)))
  } catch {
    /* storage unavailable: the list is a convenience */
  }
}

/** the name of the file a file: address points at, as the text of a link to it */
export function fileNameOfUrl(href: string): string | null {
  if (!/^file:/i.test(href)) return null
  try {
    const name = decodeURIComponent(new URL(href).pathname.split('/').pop() ?? '')
    return name || null
  } catch {
    return null
  }
}

/** Insert Hyperlink dialog, shared by the ribbon and the native application menu */
export function LinkInsertModal({ editor, onClose }: { editor: Editor; onClose: () => void }) {
  const { t } = useI18n()
  const modalKeys = useModalKeys(onClose)
  // Word parity: with the caret on an existing hyperlink the dialog EDITS it
  // — text and address pre-filled, plus Remove Link. Imported links carry the
  // same mark as in-app ones, but there was no way to view, change, or
  // remove any link after creation.
  const [linkAtOpen] = useState(() => {
    const { $from, empty } = editor.state.selection
    const markType = editor.state.schema.marks.link
    if (!markType) return null
    const range = getMarkRange($from, markType)
    if (!range) return null
    // A non-empty selection reaching outside the link is a fresh insert over
    // that selection, not an edit of the link under its endpoint.
    const { from, to } = editor.state.selection
    if (!empty && (from < range.from || to > range.to)) return null
    // Read attrs from the run itself, not the selection: at the link's
    // trailing edge $head.marks() drops inclusive:false marks, so
    // getAttributes('link') comes back empty for a real link.
    const attrs = editor.state.doc
      .nodeAt(range.from)
      ?.marks.find((mark) => mark.type === markType)?.attrs
    if (!attrs) return null
    return {
      from: range.from,
      to: range.to,
      text: editor.state.doc.textBetween(range.from, range.to, ' '),
      href: typeof attrs.href === 'string' ? attrs.href : '',
      tooltip: typeof attrs.tooltip === 'string' ? attrs.tooltip : null,
    }
  })
  // Word parity: selected text pre-populates the display-text field, so a
  // select-then-link flow only needs the address.
  const [selectionAtOpen] = useState(() => {
    const { from, to } = editor.state.selection
    return { from, to, text: from === to ? '' : editor.state.doc.textBetween(from, to, ' ') }
  })
  const [targets] = useState(() => collectLinkTargets(editor.state.doc))
  const hrefAtOpen = linkAtOpen?.href ?? ''
  const kindAtOpen = linkKindOf(hrefAtOpen)
  const [tab, setTab] = useState<LinkKind>(kindAtOpen)
  const [linkText, setLinkText] = useState(linkAtOpen ? linkAtOpen.text : selectionAtOpen.text)
  const [linkUrl, setLinkUrl] = useState(kindAtOpen === 'web' ? hrefAtOpen : '')
  const [place, setPlace] = useState<Place>(() => {
    if (kindAtOpen !== 'doc') return null
    const name = hrefAtOpen.slice(1)
    if (name === TOP_ANCHOR) return { kind: 'top' }
    const heading = targets.headings.find((h) => h.anchors.includes(name))
    return heading ? { kind: 'heading', pos: heading.pos } : { kind: 'anchor', name }
  })
  const [mail, setMail] = useState(
    () => (kindAtOpen === 'email' && parseMailto(hrefAtOpen)) || { address: '', subject: '' },
  )
  const [screenTip, setScreenTip] = useState(linkAtOpen?.tooltip ?? '')
  const [recent] = useState(recentEmails)

  /** an anchor the lists do not show (a bookmark inside an object, a target that is gone) */
  const strayAnchor =
    place?.kind === 'anchor' && !targets.bookmarks.some((b) => b.name === place.name)
      ? place.name
      : null

  /** what the active tab links to, with the text shown when the display text is left blank */
  const target = (): { href: string; label: string; headingPos?: number } | null => {
    if (tab === 'web') {
      const href = linkUrl.trim()
      return href ? { href, label: fileNameOfUrl(href) ?? href } : null
    }
    if (tab === 'email') {
      const href = buildMailto(mail)
      return href ? { href, label: mail.address.trim() } : null
    }
    if (!place) return null
    if (place.kind === 'top') return { href: `#${TOP_ANCHOR}`, label: t('ribbonLinkTopOfDoc') }
    if (place.kind === 'anchor') return { href: `#${place.name}`, label: place.name }
    const heading = targets.headings.find((h) => h.pos === place.pos)
    return heading ? { href: '', label: heading.text, headingPos: heading.pos } : null
  }
  const ready = target() !== null

  const insertLink = () => {
    const picked = target()
    if (!picked || !editor.isEditable) return
    const text = linkText.trim() || picked.label
    const tooltip = screenTip.trim() || null
    let href = picked.href
    const chain = editor.chain().focus()
    if (picked.headingPos !== undefined) {
      const pos = picked.headingPos
      // the heading's bookmark and the link are one undo step; a node attribute
      // change moves no position, so the ranges captured at open still hold
      chain.command(({ tr }) => {
        const current = kindAtOpen === 'doc' ? hrefAtOpen.slice(1) : undefined
        href = `#${ensureHeadingAnchor(tr, pos, current) ?? ''}`
        return href.length > 1
      })
    }
    const attrs = () => ({ href, rId: null, anchor: null, history: null, tooltip })
    if (linkAtOpen) {
      if (text === linkAtOpen.text.trim()) {
        // target-only change: re-mark the existing run so character
        // formatting, comments and inline objects survive
        chain
          .setTextSelection({ from: linkAtOpen.from, to: linkAtOpen.to })
          .command(({ commands }) => commands.setMark('link', attrs()))
      } else {
        chain.deleteRange({ from: linkAtOpen.from, to: linkAtOpen.to }).command(({ commands }) =>
          commands.insertContentAt(linkAtOpen.from, {
            type: 'text',
            text,
            marks: [{ type: 'link', attrs: attrs() }],
          }),
        )
      }
    } else if (selectionAtOpen.text && text === selectionAtOpen.text.trim()) {
      // untouched display text: mark the ORIGINAL selection instead of
      // re-inserting plain text — character formatting, comments and inline
      // objects in the selection survive
      chain
        .setTextSelection({ from: selectionAtOpen.from, to: selectionAtOpen.to })
        .command(({ commands }) => commands.setMark('link', attrs()))
    } else {
      chain.command(({ commands }) =>
        commands.insertContent({ type: 'text', text, marks: [{ type: 'link', attrs: attrs() }] }),
      )
    }
    if (!chain.run()) return
    if (tab === 'email') rememberEmail(mail.address.trim())
    onClose()
  }

  const removeLink = () => {
    if (!linkAtOpen || !editor.isEditable) return
    // Word's Remove Hyperlink: the text stays, only the link goes.
    editor
      .chain()
      .focus()
      .setTextSelection({ from: linkAtOpen.from, to: linkAtOpen.to })
      .unsetMark('link')
      .run()
    onClose()
  }

  const onEnter = (e: { key: string }) => e.key === 'Enter' && insertLink()

  const placeRow = (key: string, label: string, value: Place, level = 0, detail = '') => {
    const selected = JSON.stringify(place) === JSON.stringify(value)
    return (
      <button
        key={key}
        type="button"
        role="option"
        aria-selected={selected}
        className={selected ? 'link-place active' : 'link-place'}
        style={{ paddingInlineStart: 10 + level * 14 }}
        onClick={() => setPlace(value)}
        onDoubleClick={insertLink}
      >
        <span className="link-place-name">{label}</span>
        {detail && <span className="link-place-detail">{detail}</span>}
      </button>
    )
  }

  return (
    <div
      className="modal-backdrop"
      ref={modalKeys.ref}
      onKeyDown={modalKeys.onKeyDown}
      onMouseDown={(e) => e.target === e.currentTarget && onClose()}
    >
      <div className="modal link-dialog">
        <h2>{t(linkAtOpen ? 'ribbonLinkEditTitle' : 'ribbonLinkInsertTitle')}</h2>
        <label>
          {t('ribbonLinkText')}
          <input
            name="linkText"
            value={linkText}
            onChange={(e) => setLinkText(e.target.value)}
            placeholder={t('ribbonLinkTextPh')}
          />
        </label>
        <div className="dlg-tabs" role="tablist">
          {TABS.map(({ key, nameKey }) => (
            <button
              key={key}
              type="button"
              role="tab"
              aria-selected={tab === key}
              className={tab === key ? 'active' : ''}
              onClick={() => setTab(key)}
            >
              {t(nameKey)}
            </button>
          ))}
        </div>
        {tab === 'web' && (
          <div className="link-address-row">
            <label>
              {t('ribbonLinkAddress')}
              <input
                name="linkAddress"
                value={linkUrl}
                onChange={(e) => setLinkUrl(e.target.value)}
                placeholder="https://…"
                onKeyDown={onEnter}
              />
            </label>
            <button
              type="button"
              name="linkSelectFile"
              onClick={() => {
                void window.desktop.pickLinkFile().then((file) => {
                  if (file) setLinkUrl(file.url)
                })
              }}
            >
              {t('ribbonLinkSelect')}
            </button>
          </div>
        )}
        {tab === 'doc' && (
          <div className="link-places" role="listbox" aria-label={t('ribbonLinkTabDoc')}>
            {placeRow('top', t('ribbonLinkTopOfDoc'), { kind: 'top' })}
            {targets.headings.length > 0 && (
              <div className="link-place-group">{t('appNavTabHeadings')}</div>
            )}
            {targets.headings.map((h) =>
              placeRow(`h${h.pos}`, h.text, { kind: 'heading', pos: h.pos }, h.level),
            )}
            {(targets.bookmarks.length > 0 || strayAnchor) && (
              <div className="link-place-group">{t('ribbonBookmark')}</div>
            )}
            {targets.bookmarks.map((b) =>
              placeRow(
                `b${b.pos}-${b.name}`,
                b.name,
                { kind: 'anchor', name: b.name },
                1,
                b.preview,
              ),
            )}
            {strayAnchor &&
              placeRow('stray', strayAnchor, { kind: 'anchor', name: strayAnchor }, 1)}
          </div>
        )}
        {tab === 'email' && (
          <>
            <label>
              {t('ribbonLinkTabEmail')}
              <input
                name="linkEmail"
                type="email"
                list="link-recent-emails"
                value={mail.address}
                onChange={(e) => setMail((m) => ({ ...m, address: e.target.value }))}
                placeholder="name@example.com"
                onKeyDown={onEnter}
              />
            </label>
            <datalist id="link-recent-emails">
              {recent.map((address) => (
                <option key={address} value={address} />
              ))}
            </datalist>
            <label>
              {t('ribbonPropSubject')}
              <input
                name="linkSubject"
                value={mail.subject}
                onChange={(e) => setMail((m) => ({ ...m, subject: e.target.value }))}
                onKeyDown={onEnter}
              />
            </label>
          </>
        )}
        <label>
          {t('ribbonLinkScreenTip')}
          <input
            name="linkScreenTip"
            value={screenTip}
            onChange={(e) => setScreenTip(e.target.value)}
            onKeyDown={onEnter}
          />
        </label>
        <div className="modal-actions">
          {linkAtOpen && (
            <button className="btn-ghost" onClick={removeLink}>
              {t('ribbonLinkRemove')}
            </button>
          )}
          <button className="btn-ghost" onClick={onClose}>
            {t('ribbonCancel')}
          </button>
          <button className="btn-primary" disabled={!ready} onClick={insertLink}>
            {t(linkAtOpen ? 'ribbonApply' : 'ribbonInsert')}
          </button>
        </div>
      </div>
    </div>
  )
}
