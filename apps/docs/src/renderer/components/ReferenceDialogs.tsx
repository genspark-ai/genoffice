/** The Bookmark, Cross-reference and Caption dialogs. */
import { useMemo, useState } from 'react'
import { NodeSelection } from '@tiptap/pm/state'
import type { Editor } from '@tiptap/core'
import {
  CHAPTER_SEPARATORS,
  SEQ_NUMBER_FORMATS,
  captionNumberingOfXml,
  generateCaptionXml,
  type Block,
  type CaptionNumbering,
  type SeqNumberFormat,
} from '@genoffice/docx-engine'
import { captionLabels, collectCaptions } from '../commands/reference-commands'
import {
  captionPreview,
  numberingContextOf,
  renumberAfterInsert,
  textOfXml,
  unreliableSeqLabelsOf,
} from '../editor/field-results'
import {
  AUTO_CAPTION_KINDS,
  autoCaptionRules,
  saveAutoCaptionRules,
  type AutoCaptionKind,
} from '../editor/auto-caption'
import {
  addBookmark,
  collectBookmarks,
  goToBookmark,
  insertCrossReference,
  isParaNumber,
  listRefTargets,
  removeBookmark,
  sortBookmarks,
  validBookmarkName,
  type RefContent,
  type RefNotes,
  type RefTarget,
  type RefType,
} from '../editor/cross-reference'
import { xmlOfNode } from '../editor/toc-update'
import { useI18n, type StringKey } from '../i18n/locale'
import { PromptModal } from './PromptModal'
import { useModalKeys } from './modal-keys'

/** Word's Bookmark dialog: name, list, Hidden bookmarks, Sort by, Add / Delete / Go To */
export function BookmarkModal({ editor, onClose }: { editor: Editor; onClose: () => void }) {
  const { t } = useI18n()
  const keys = useModalKeys(onClose)
  const [name, setName] = useState('')
  const [error, setError] = useState<StringKey | null>(null)
  const [showHidden, setShowHidden] = useState(false)
  const [sortBy, setSortBy] = useState<'name' | 'location'>('name')
  const [, forceRender] = useState(0)
  const bookmarks = sortBookmarks(collectBookmarks(editor.state.doc, showHidden), sortBy)
  const picked = bookmarks.find((b) => b.name === name.trim()) ?? null

  const add = () => {
    const trimmed = name.trim()
    if (!validBookmarkName(trimmed)) return setError('ribbonBookmarkNameInvalid')
    if (!addBookmark(editor, trimmed)) return setError('ribbonBookmarkNoCursor')
    setError(null)
    forceRender((n) => n + 1)
  }

  const remove = () => {
    if (!picked) return
    removeBookmark(editor, picked)
    setName('')
  }

  return (
    <div
      className="modal-backdrop"
      ref={keys.ref}
      onKeyDown={keys.onKeyDown}
      onMouseDown={(e) => e.target === e.currentTarget && onClose()}
    >
      <div className="modal bookmark-dialog" role="dialog" aria-label={t('ribbonBookmark')}>
        <h2>{t('ribbonBookmark')}</h2>
        <input
          autoFocus
          name="bookmark-name"
          value={name}
          placeholder={t('ribbonBookmarkNamePh')}
          aria-label={t('ribbonBookmarkNamePh')}
          onChange={(e) => {
            setName(e.target.value)
            setError(null)
          }}
          onKeyDown={(e) => e.key === 'Enter' && add()}
        />
        {error && <div className="modal-error">{t(error)}</div>}
        <div className="bookmark-list" role="listbox" aria-label={t('ribbonBookmark')}>
          {bookmarks.length === 0 && (
            <div className="bookmark-empty">{t('ribbonBookmarkEmpty')}</div>
          )}
          {bookmarks.map((b) => (
            <button
              type="button"
              key={`${b.name}-${b.pos}`}
              role="option"
              aria-selected={b === picked}
              className={`bookmark-row ${b === picked ? 'on' : ''}`}
              onClick={() => {
                setName(b.name)
                setError(null)
              }}
              onDoubleClick={() => goToBookmark(editor, b)}
            >
              <span className="bookmark-name">{b.name}</span>
              <span className="bookmark-preview">{b.preview}</span>
            </button>
          ))}
        </div>
        <div className="ref-radios" role="radiogroup" aria-label={t('ribbonBookmarkSortBy')}>
          <span>{t('ribbonBookmarkSortBy')}</span>
          {(['name', 'location'] as const).map((by) => (
            <label key={by} className="ref-check">
              <input
                type="radio"
                name="bookmark-sort"
                value={by}
                checked={sortBy === by}
                onChange={() => setSortBy(by)}
              />
              {t(by === 'name' ? 'ribbonStyleDialogName' : 'ribbonBookmarkSortLocation')}
            </label>
          ))}
        </div>
        <label className="ref-check">
          <input
            type="checkbox"
            name="bookmark-hidden"
            checked={showHidden}
            onChange={(e) => setShowHidden(e.target.checked)}
          />
          {t('ribbonBookmarkHidden')}
        </label>
        <div className="modal-actions">
          <button
            name="go-to"
            disabled={!picked}
            onClick={() => picked && goToBookmark(editor, picked)}
          >
            {t('appGoToTitle')}
          </button>
          <button name="delete" disabled={!picked} onClick={remove}>
            {t('appNavDelete')}
          </button>
          <button onClick={onClose}>{t('ribbonClose')}</button>
          <button className="btn-primary" name="add" disabled={!name.trim()} onClick={add}>
            {t('ribbonAdd')}
          </button>
        </div>
      </div>
    </div>
  )
}

const REF_CONTENTS: RefContent[] = ['text', 'page', 'aboveBelow']
// Word's order for a note: its number, its page, where it is, its number as a reference mark
const NOTE_REF_CONTENTS: RefContent[] = ['noteNumber', 'page', 'aboveBelow', 'noteNumberFormatted']
const NUMBERED_REF_CONTENTS: RefContent[] = [
  'page',
  'paraNumber',
  'paraNumberNoContext',
  'paraNumberFullContext',
  'text',
  'aboveBelow',
]

// Word's order for a heading: its number comes after its page
const HEADING_REF_CONTENTS: RefContent[] = [
  'text',
  'page',
  'paraNumber',
  'paraNumberNoContext',
  'paraNumberFullContext',
  'aboveBelow',
]

const refContentsOf = (kind: RefType['kind']): RefContent[] =>
  kind === 'numbered'
    ? NUMBERED_REF_CONTENTS
    : kind === 'heading'
      ? HEADING_REF_CONTENTS
      : kind === 'footnote' || kind === 'endnote'
        ? NOTE_REF_CONTENTS
        : REF_CONTENTS

/** what Word's dialog picks first for the type */
const firstRefContent = (kind: RefType['kind']): RefContent =>
  kind === 'numbered' ? 'paraNumber' : refContentsOf(kind)[0]

const PARA_NUMBER_KEYS: Partial<Record<RefContent, StringKey>> = {
  paraNumber: 'ribbonCrossRefParaNumber',
  paraNumberNoContext: 'ribbonCrossRefParaNumberNoContext',
  paraNumberFullContext: 'ribbonCrossRefParaNumberFullContext',
}

const HEADING_NUMBER_KEYS: Partial<Record<RefContent, StringKey>> = {
  paraNumber: 'ribbonCrossRefHeadingNumber',
  paraNumberNoContext: 'ribbonCrossRefHeadingNumberNoContext',
  paraNumberFullContext: 'ribbonCrossRefHeadingNumberFullContext',
}

const refTypeOf = (value: string): RefType =>
  value === 'numbered' ||
  value === 'bookmark' ||
  value === 'heading' ||
  value === 'footnote' ||
  value === 'endnote'
    ? { kind: value }
    : { kind: 'caption', label: value.slice('caption:'.length) }

/** Word's Cross-reference dialog: what to point at, what the reference shows, and how */
export function CrossRefModal({
  editor,
  blocks,
  notes,
  pagesOf,
  onClose,
}: {
  editor: Editor
  blocks: Block[]
  /** the document's notes as they stand now: a note is listed by its number and the start of its text */
  notes: RefNotes
  pagesOf: (positions: number[]) => number[] | null
  onClose: () => void
}) {
  const { t } = useI18n()
  const keys = useModalKeys(onClose)
  const [typeValue, setTypeValue] = useState('heading')
  const [content, setContent] = useState<RefContent>('text')
  const [hyperlink, setHyperlink] = useState(true)
  const [aboveBelow, setAboveBelow] = useState(false)
  const [separate, setSeparate] = useState(false)
  const [separator, setSeparator] = useState('')
  const [pick, setPick] = useState(0)
  const type = refTypeOf(typeValue)
  const contents = refContentsOf(type.kind)
  const takesAboveBelow = content === 'page' || isParaNumber(content)
  const takesSeparator = content === 'paraNumberFullContext'
  const targets = listRefTargets(editor, blocks, type, notes)
  const target = targets[pick] as RefTarget | undefined
  const labels = [
    ...new Set([
      t('ribbonCaptionEquation'),
      t('ribbonCaptionFigure'),
      t('ribbonCaptionTable'),
      ...captionLabels(editor, blocks),
    ]),
  ]
  const textKey: StringKey =
    type.kind === 'bookmark'
      ? 'ribbonCrossRefBookmarkText'
      : type.kind === 'heading'
        ? 'ribbonCrossRefHeadingText'
        : type.kind === 'numbered'
          ? 'ribbonCrossRefParaText'
          : 'ribbonCrossRefEntireCaption'
  const contentKey = (c: RefContent): StringKey => {
    if (c === 'text') return textKey
    if (c === 'page') return 'ribbonPageNumber'
    if (c === 'aboveBelow') return 'ribbonCrossRefAboveBelow'
    const paraNumber = (type.kind === 'heading' ? HEADING_NUMBER_KEYS : PARA_NUMBER_KEYS)[c]
    if (paraNumber) return paraNumber
    const endnote = type.kind === 'endnote'
    if (c === 'noteNumber')
      return endnote ? 'ribbonCrossRefEndnoteNumber' : 'ribbonCrossRefFootnoteNumber'
    return endnote
      ? 'ribbonCrossRefEndnoteNumberFormatted'
      : 'ribbonCrossRefFootnoteNumberFormatted'
  }

  const insert = () => {
    if (!target) return
    // pages are looked up by top-level block
    const top = (pos: number) => {
      const $pos = editor.state.doc.resolve(pos)
      return $pos.depth > 0 ? $pos.before(1) : pos
    }
    const found =
      content === 'page' ? pagesOf([top(target.pos), top(editor.state.selection.from)]) : null
    const done = insertCrossReference(editor, blocks, {
      target,
      content,
      hyperlink,
      includeAboveBelow: takesAboveBelow && aboveBelow,
      heading: type.kind === 'heading',
      ...(takesSeparator && separate && separator ? { separator } : {}),
      pages: found && found.length === 2 ? { target: found[0], here: found[1] } : null,
      words: {
        above: t('ribbonCrossRefAbove'),
        below: t('ribbonCrossRefBelow'),
        onPage: (n) => t('ribbonCrossRefOnPage', { n }),
      },
    })
    if (done) onClose()
  }

  return (
    <div
      className="modal-backdrop"
      ref={keys.ref}
      onKeyDown={keys.onKeyDown}
      onMouseDown={(e) => e.target === e.currentTarget && onClose()}
    >
      <div className="modal crossref-dialog" role="dialog" aria-label={t('ribbonCrossRef')}>
        <h2>{t('ribbonCrossRef')}</h2>
        <div className="modal-row">
          <label>
            {t('ribbonCrossRefType')}
            <select
              name="ref-type"
              value={typeValue}
              onChange={(e) => {
                const next = refTypeOf(e.target.value).kind
                setTypeValue(e.target.value)
                setPick(0)
                // another set of choices opens on the one Word picks first
                if (refContentsOf(next) !== contents) setContent(firstRefContent(next))
              }}
            >
              <option value="numbered">{t('ribbonCrossRefNumbered')}</option>
              <option value="heading">{t('appGoToHeading')}</option>
              <option value="bookmark">{t('ribbonBookmark')}</option>
              <option value="footnote">{t('appGoToFootnote')}</option>
              <option value="endnote">{t('appGoToEndnote')}</option>
              {labels.map((l) => (
                <option key={l} value={`caption:${l}`}>
                  {l}
                </option>
              ))}
            </select>
          </label>
          <label>
            {t('ribbonCrossRefTo')}
            <select
              name="ref-content"
              value={content}
              onChange={(e) => setContent(e.target.value as RefContent)}
            >
              {contents.map((c) => (
                <option key={c} value={c}>
                  {t(contentKey(c))}
                </option>
              ))}
            </select>
          </label>
        </div>
        <label className="ref-check">
          <input
            type="checkbox"
            name="ref-hyperlink"
            checked={hyperlink}
            onChange={(e) => setHyperlink(e.target.checked)}
          />
          {t('ribbonCrossRefHyperlink')}
        </label>
        <label className="ref-check">
          <input
            type="checkbox"
            name="ref-above-below"
            checked={takesAboveBelow && aboveBelow}
            disabled={!takesAboveBelow}
            onChange={(e) => setAboveBelow(e.target.checked)}
          />
          {t('ribbonCrossRefIncludeAboveBelow')}
        </label>
        <div className="ref-separator">
          <label className="ref-check">
            <input
              type="checkbox"
              name="ref-separator"
              checked={takesSeparator && separate}
              disabled={!takesSeparator}
              onChange={(e) => setSeparate(e.target.checked)}
            />
            {t('ribbonCrossRefSeparate')}
          </label>
          <input
            type="text"
            name="ref-separator-text"
            aria-label={t('ribbonCrossRefSeparate')}
            value={takesSeparator && separate ? separator : ''}
            disabled={!takesSeparator || !separate}
            onChange={(e) => setSeparator(e.target.value)}
          />
        </div>
        <div className="bookmark-list" role="listbox" aria-label={t('ribbonCrossRef')}>
          {targets.length === 0 && (
            <div className="bookmark-empty">
              {t(type.kind === 'bookmark' ? 'ribbonCrossRefEmpty' : 'appGoToNotFound')}
            </div>
          )}
          {targets.map((item, i) => (
            <button
              type="button"
              key={`${item.pos}-${item.bookmark ?? ''}`}
              role="option"
              aria-selected={i === pick}
              className={`bookmark-row ${i === pick ? 'on' : ''}`}
              style={item.level ? { paddingLeft: 10 + (item.level - 1) * 12 } : undefined}
              onClick={() => setPick(i)}
              onDoubleClick={insert}
            >
              {type.kind === 'bookmark' && <span className="bookmark-name">{item.bookmark}</span>}
              <span className="bookmark-preview">{item.text}</span>
            </button>
          ))}
        </div>
        <div className="modal-actions">
          <button onClick={onClose}>{t('ribbonCancel')}</button>
          <button className="btn-primary" name="insert" disabled={!target} onClick={insert}>
            {t('ribbonInsert')}
          </button>
        </div>
      </div>
    </div>
  )
}

const CAPTION_LABEL_KEYS = [
  'ribbonCaptionFigure',
  'ribbonCaptionTable',
  'ribbonCaptionEquation',
] as const

/** A caption goes below (or above) a selected picture or object, as in Word; typing over it would delete it. */
export function insertCaptionBlock(
  editor: Editor,
  attrs: Record<string, unknown>,
  position: 'above' | 'below' = 'below',
): void {
  const { selection } = editor.state
  const node = { type: 'docProtected', attrs } as never
  if (!(selection instanceof NodeSelection)) {
    editor.chain().focus().insertContent(node).run()
    return
  }
  // climb to the first ancestor that takes a caption block: past the paragraph of
  // an in-line picture, and past the table when the picture sits in a cell
  const above = position === 'above'
  const $edge = above ? selection.$from : selection.$to
  const type = editor.schema.nodes.docProtected
  let at = above ? selection.from : selection.to
  for (let depth = $edge.depth; depth >= 0; depth--) {
    const index = above ? $edge.index(depth) : $edge.indexAfter(depth)
    if ($edge.node(depth).canReplaceWith(index, index, type)) break
    if (depth > 0) at = above ? $edge.before(depth) : $edge.after(depth)
  }
  editor.chain().focus().insertContentAt(at, node).run()
}

const CUSTOM_LABELS_KEY = 'docs.captionLabels'

function storedLabels(): string[] {
  try {
    const list: unknown = JSON.parse(localStorage.getItem(CUSTOM_LABELS_KEY) ?? '[]')
    return Array.isArray(list) ? list.filter((l): l is string => typeof l === 'string') : []
  } catch {
    return []
  }
}

const FORMAT_SAMPLES: Record<SeqNumberFormat, string> = {
  ARABIC: '1, 2, 3, …',
  alphabetic: 'a, b, c, …',
  ALPHABETIC: 'A, B, C, …',
  roman: 'i, ii, iii, …',
  ROMAN: 'I, II, III, …',
}

const AUTO_CAPTION_NAMES: Record<AutoCaptionKind, StringKey> = {
  table: 'ribbonCaptionTable',
  equation: 'ribbonCaptionEquation',
}

/** Word's AutoCaption: which inserted objects get a caption by themselves, and what it looks like */
function AutoCaptionModal({ labels, onClose }: { labels: string[]; onClose: () => void }) {
  const { t } = useI18n()
  const keys = useModalKeys(onClose)
  const [rules, setRules] = useState(autoCaptionRules)
  const [picked, setPicked] = useState<AutoCaptionKind>('table')
  const rule = rules[picked]
  const set = (kind: AutoCaptionKind, patch: Partial<typeof rule>) =>
    setRules({ ...rules, [kind]: { ...rules[kind], ...patch } })
  const offered = [...new Set([...labels, rule.label])]
  return (
    <div
      className="modal-backdrop"
      ref={keys.ref}
      onKeyDown={keys.onKeyDown}
      onMouseDown={(e) => e.target === e.currentTarget && onClose()}
    >
      <div className="modal caption-dialog" role="dialog" aria-label={t('ribbonAutoCaptionTitle')}>
        <h2>{t('ribbonAutoCaptionTitle')}</h2>
        <p className="modal-desc">{t('ribbonAutoCaptionWhen')}</p>
        <div className="auto-caption-list" role="listbox">
          {AUTO_CAPTION_KINDS.map((kind) => (
            <label
              key={kind}
              className={`ref-check${picked === kind ? ' picked' : ''}`}
              onMouseDown={() => setPicked(kind)}
            >
              <input
                type="checkbox"
                name={`auto-caption-${kind}`}
                checked={rules[kind].on}
                onChange={(e) => {
                  setPicked(kind)
                  set(kind, { on: e.target.checked })
                }}
              />
              {t(AUTO_CAPTION_NAMES[kind])}
            </label>
          ))}
        </div>
        <div className="modal-row">
          <label>
            {t('ribbonAutoCaptionUseLabel')}
            <select
              name="auto-caption-label"
              value={rule.label}
              onChange={(e) => set(picked, { label: e.target.value })}
            >
              {offered.map((l) => (
                <option key={l} value={l}>
                  {l}
                </option>
              ))}
            </select>
          </label>
          <label>
            {t('ribbonPosition')}
            <select
              name="auto-caption-position"
              value={rule.position}
              onChange={(e) => set(picked, { position: e.target.value as 'above' | 'below' })}
            >
              <option value="above">{t('ribbonCaptionAbove')}</option>
              <option value="below">{t('ribbonCaptionBelow')}</option>
            </select>
          </label>
        </div>
        <div className="modal-row">
          <label>
            {t('ribbonNumbering')}
            <select
              name="auto-caption-format"
              value={rule.format}
              onChange={(e) => set(picked, { format: e.target.value as SeqNumberFormat })}
            >
              {SEQ_NUMBER_FORMATS.map((f) => (
                <option key={f} value={f}>
                  {FORMAT_SAMPLES[f]}
                </option>
              ))}
            </select>
          </label>
        </div>
        <div className="modal-actions">
          <button className="btn-ghost" onClick={onClose}>
            {t('ribbonCancel')}
          </button>
          <button
            className="btn-primary"
            name="ok"
            onClick={() => {
              saveAutoCaptionRules(rules)
              onClose()
            }}
          >
            {t('ribbonOk')}
          </button>
        </div>
      </div>
    </div>
  )
}

/** Word's Caption Numbering: the number format of a label's captions, and the chapter number in front of it */
function CaptionNumberingModal({
  numbering,
  onSubmit,
  onClose,
}: {
  numbering: CaptionNumbering
  onSubmit: (numbering: CaptionNumbering) => void
  onClose: () => void
}) {
  const { t } = useI18n()
  const keys = useModalKeys(onClose)
  const [format, setFormat] = useState(numbering.format)
  const [on, setOn] = useState(numbering.chapter !== null)
  // kept while the box is unticked, so ticking it again brings the same choices back
  const [shown, setShown] = useState(numbering.chapter ?? { style: 1, sep: 'hyphen' })
  const chapter = on ? shown : null
  return (
    <div
      className="modal-backdrop"
      ref={keys.ref}
      onKeyDown={keys.onKeyDown}
      onMouseDown={(e) => e.target === e.currentTarget && onClose()}
    >
      <div
        className="modal caption-dialog"
        role="dialog"
        aria-label={t('ribbonCaptionNumberingTitle')}
      >
        <h2>{t('ribbonCaptionNumberingTitle')}</h2>
        <div className="modal-row">
          <label>
            {t('appNumberFormat')}
            <select
              name="caption-numbering-format"
              value={format}
              onChange={(e) => setFormat(e.target.value as SeqNumberFormat)}
            >
              {SEQ_NUMBER_FORMATS.map((f) => (
                <option key={f} value={f}>
                  {FORMAT_SAMPLES[f]}
                </option>
              ))}
            </select>
          </label>
        </div>
        <label className="ref-check">
          <input
            type="checkbox"
            name="caption-chapter"
            checked={chapter !== null}
            onChange={(e) => setOn(e.target.checked)}
          />
          {t('appPgNumIncludeChapter')}
        </label>
        <div className="modal-row">
          <label>
            {t('appPgNumChapterStyle')}
            <select
              name="caption-chapter-style"
              disabled={!chapter}
              value={shown.style}
              onChange={(e) => setShown({ ...shown, style: Number(e.target.value) })}
            >
              {[1, 2, 3, 4, 5, 6, 7, 8, 9].map((level) => (
                <option key={level} value={level}>
                  {t('ribbonStyleHeadingN', { n: level })}
                </option>
              ))}
            </select>
          </label>
          <label>
            {t('appPgNumSeparator')}
            <select
              name="caption-chapter-sep"
              disabled={!chapter}
              value={shown.sep}
              onChange={(e) => setShown({ ...shown, sep: e.target.value })}
            >
              {Object.entries(CHAPTER_SEPARATORS).map(([name, glyph]) => (
                <option key={name} value={name}>
                  {glyph}
                </option>
              ))}
            </select>
          </label>
        </div>
        <div className="modal-actions">
          <button className="btn-ghost" onClick={onClose}>
            {t('ribbonCancel')}
          </button>
          <button className="btn-primary" name="ok" onClick={() => onSubmit({ format, chapter })}>
            {t('ribbonOk')}
          </button>
        </div>
      </div>
    </div>
  )
}

/** Caption dialog: label, position, number format and an optional description, numbered by existing SEQ fields */
export function CaptionModal({
  editor,
  blocks,
  onClose,
}: {
  editor: Editor
  blocks: Block[]
  onClose: () => void
}) {
  const { t } = useI18n()
  const keys = useModalKeys(onClose)
  const builtIn = CAPTION_LABEL_KEYS.map((k) => t(k))
  const [custom, setCustom] = useState<string[]>(storedLabels)
  const [label, setLabel] = useState<string>(() => t(CAPTION_LABEL_KEYS[0]))
  const [text, setText] = useState('')
  const [position, setPosition] = useState<'above' | 'below'>('below')
  const [picked, setPicked] = useState<Record<string, CaptionNumbering>>({})
  const [excludeLabel, setExcludeLabel] = useState(false)
  const [naming, setNaming] = useState(false)
  const [auto, setAuto] = useState(false)
  const [formatting, setFormatting] = useState(false)
  const captions = collectCaptions(editor, blocks)
  const labels = [...new Set([...builtIn, ...captions.map((c) => c.label), ...custom])]
  const objectSelected = editor.state.selection instanceof NodeSelection

  const saveCustom = (list: string[]) => {
    setCustom(list)
    localStorage.setItem(CUSTOM_LABELS_KEY, JSON.stringify(list))
  }

  /** the numbering set here, else the one the label's last caption already has */
  const numbering = ((): CaptionNumbering => {
    if (picked[label]) return picked[label]
    const last = captions.filter((c) => c.label === label).pop()
    const node = last ? editor.state.doc.nodeAt(last.pos) : null
    const used = node ? captionNumberingOfXml(xmlOfNode(node as never, blocks), label) : null
    return used?.numbering ?? { format: 'ARABIC', chapter: null }
  })()
  const { format, chapter } = numbering
  const setNumbering = (next: CaptionNumbering) => setPicked({ ...picked, [label]: next })

  const blockOf = (tail: string) => {
    // a label this editor cannot number reliably keeps Word's own update for the new caption
    const dirty = unreliableSeqLabelsOf(blocks, editor.schema, numberingContextOf(editor)).has(
      label,
    )
    const number = captions.filter((c) => c.label === label).length + 1
    const xml = generateCaptionXml(label, number, tail, { format, chapter, excludeLabel, dirty })
    return {
      docxIndex: null,
      blockType: 'passthrough',
      label: t('ribbonCaption'),
      genXml: xml,
      fieldDisplay: { kind: 'text', left: textOfXml(xml) },
    }
  }
  /** Word applies a label's numbering to the captions it already has */
  const applied = picked[label] ? { label, ...picked[label] } : undefined
  const shown = useMemo(
    () => captionPreview(editor, blocks, blockOf(''), applied),
    // eslint-disable-next-line react-hooks/exhaustive-deps -- the document does not change while the dialog is open
    [label, format, chapter?.style, chapter?.sep, excludeLabel],
  )

  const insert = () => {
    insertCaptionBlock(editor, blockOf(text.trim()), objectSelected ? position : 'below')
    renumberAfterInsert(editor, blocks, applied)
    onClose()
  }

  if (formatting)
    return (
      <CaptionNumberingModal
        numbering={numbering}
        onSubmit={(next) => {
          setNumbering(next)
          setFormatting(false)
        }}
        onClose={() => setFormatting(false)}
      />
    )

  if (auto) return <AutoCaptionModal labels={labels} onClose={() => setAuto(false)} />

  if (naming)
    return (
      <PromptModal
        title={t('ribbonCaptionNewLabel')}
        placeholder={t('ribbonCaptionLabel')}
        onSubmit={(value) => {
          const name = value.trim().replace(/"/g, '')
          if (name && !labels.includes(name)) saveCustom([...custom, name])
          if (name) setLabel(name)
          setNaming(false)
        }}
        onClose={() => setNaming(false)}
      />
    )

  return (
    <div
      className="modal-backdrop"
      ref={keys.ref}
      onKeyDown={keys.onKeyDown}
      onMouseDown={(e) => e.target === e.currentTarget && onClose()}
    >
      <div
        className="modal caption-dialog"
        role="dialog"
        aria-label={t('ribbonCaptionInsertTitle')}
      >
        <h2>{t('ribbonCaptionInsertTitle')}</h2>
        <label>
          {t('ribbonCaption')}
          <input
            name="caption-text"
            value={text}
            onChange={(e) => setText(e.target.value)}
            placeholder={t('ribbonCaptionPh', {
              label: excludeLabel ? '' : label,
              n: excludeLabel ? shown : shown.slice(label.length).trim(),
            }).trim()}
            onKeyDown={(e) => e.key === 'Enter' && insert()}
          />
        </label>
        <div className="modal-row">
          <label>
            {t('ribbonCaptionLabel')}
            <select name="caption-label" value={label} onChange={(e) => setLabel(e.target.value)}>
              {labels.map((l) => (
                <option key={l} value={l}>
                  {l}
                </option>
              ))}
            </select>
          </label>
          <label>
            {t('ribbonPosition')}
            {/* Word offers a side only for a selected item; at the cursor the caption goes where the cursor is */}
            <select
              name="caption-position"
              value={position}
              disabled={!objectSelected}
              onChange={(e) => setPosition(e.target.value as 'above' | 'below')}
            >
              <option value="above">{t('ribbonCaptionAbove')}</option>
              <option value="below">{t('ribbonCaptionBelow')}</option>
            </select>
          </label>
        </div>
        <div className="modal-row">
          <label>
            {t('ribbonNumbering')}
            <select
              name="caption-format"
              value={format}
              onChange={(e) =>
                setNumbering({ ...numbering, format: e.target.value as SeqNumberFormat })
              }
            >
              {SEQ_NUMBER_FORMATS.map((f) => (
                <option key={f} value={f}>
                  {FORMAT_SAMPLES[f]}
                </option>
              ))}
            </select>
          </label>
          <div className="modal-actions caption-numbering-action">
            <button name="caption-numbering" onClick={() => setFormatting(true)}>
              {t('ribbonCaptionNumberingFormat')}
            </button>
          </div>
        </div>
        <label className="ref-check">
          <input
            type="checkbox"
            name="caption-exclude-label"
            checked={excludeLabel}
            onChange={(e) => setExcludeLabel(e.target.checked)}
          />
          {t('ribbonCaptionExcludeLabel')}
        </label>
        <div className="modal-actions caption-label-actions">
          <button name="new-label" onClick={() => setNaming(true)}>
            {t('ribbonCaptionNewLabel')}
          </button>
          <button
            name="delete-label"
            disabled={!custom.includes(label)}
            onClick={() => {
              saveCustom(custom.filter((l) => l !== label))
              setLabel(builtIn[0])
            }}
          >
            {t('ribbonCaptionDeleteLabel')}
          </button>
          <button name="auto-caption" onClick={() => setAuto(true)}>
            {t('ribbonCaptionAuto')}
          </button>
        </div>
        <div className="modal-actions">
          <button className="btn-ghost" onClick={onClose}>
            {t('ribbonCancel')}
          </button>
          <button className="btn-primary" name="ok" onClick={insert}>
            {t('ribbonInsert')}
          </button>
        </div>
      </div>
    </div>
  )
}
