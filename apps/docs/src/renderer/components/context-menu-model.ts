import type { ActionCommandId } from '../../shared/commands'
import type { RevisionRange } from '../editor/revisions'
import type { StringKey } from '../i18n/locale'

/** one row of a right-click section: an item running a command, or a submenu of such items */
export interface CtxEntry {
  id: string
  labelKey?: StringKey
  /** text shown as it is (a spelling suggestion) */
  label?: string
  ellipsis?: boolean
  command?: ActionCommandId
  payload?: string
  disabled?: boolean
  strong?: boolean
  children?: CtxEntry[]
}
export type CtxRow = CtxEntry | 'sep'

export interface SpellTarget {
  word: string
  suggestions: string[]
  /** where the word is; null when only the browser's marker knows */
  pos: number | null
  /** the browser has the word marked, so its own replacement keeps the other marks of the paragraph */
  marked: boolean
}

const SUGGESTIONS_SHOWN = 5

/** Word's menu on a misspelled word: suggestions, Ignore, Ignore All, Add to Dictionary, AutoCorrect, Spelling… */
export function spellRows(target: SpellTarget, canEdit: boolean): CtxRow[] {
  const { word, pos, marked } = target
  const offered = target.suggestions.slice(0, SUGGESTIONS_SHOWN)
  const pick = (to: string) => JSON.stringify({ pos, word, to, marked })
  const suggestions: CtxEntry[] = offered.map((to) => ({
    id: 'spell-suggestion',
    label: to,
    strong: true,
    command: 'spell-replace',
    payload: pick(to),
    disabled: !canEdit,
  }))
  return [
    ...(suggestions.length
      ? suggestions
      : [{ id: 'spell-none', labelKey: 'appSpellNoSuggestions' as const, disabled: true }]),
    'sep',
    {
      id: 'spell-ignore-once',
      labelKey: 'appSpellingIgnoreOnce',
      command: 'spell-ignore-once',
      payload: JSON.stringify({ pos, word }),
      disabled: pos === null,
    },
    {
      id: 'spell-ignore-all',
      labelKey: 'appSpellIgnoreAll',
      command: 'spell-ignore-all',
      payload: word,
    },
    {
      id: 'spell-add',
      labelKey: 'appSpellAddToDictionary',
      command: 'spell-add-word',
      payload: word,
    },
    {
      id: 'spell-autocorrect',
      labelKey: 'appAcTabAutoCorrect',
      disabled: !canEdit || offered.length === 0,
      children: offered.map((to) => ({
        id: 'spell-autocorrect-pair',
        label: to,
        command: 'spell-autocorrect',
        payload: pick(to),
      })),
    },
    {
      id: 'spell-dialog',
      labelKey: 'ribbonSpellcheckBtn',
      ellipsis: true,
      command: 'tools-spelling',
      payload: pos === null ? undefined : String(pos),
    },
  ]
}

export interface CommentTarget {
  id: string | null
  /** the thread is resolved */
  done: boolean
}

/** Word's comment items: in the text menu they stay listed, greyed, off a comment; a balloon gets them alone */
export function commentRows(comment: CommentTarget): CtxRow[] {
  const id = comment.id ?? undefined
  const off = !comment.id
  return [
    {
      id: 'comment-reply',
      labelKey: 'appCtxReplyComment',
      command: 'comment-reply',
      payload: id,
      disabled: off || comment.done,
    },
    {
      id: 'comment-delete',
      labelKey: 'appCtxDeleteComment',
      command: 'delete-comment',
      payload: id,
      disabled: off,
    },
    comment.done
      ? {
          id: 'comment-reopen',
          labelKey: 'appCtxReopenComment',
          command: 'comment-reopen',
          payload: id,
          disabled: off,
        }
      : {
          id: 'comment-resolve',
          labelKey: 'appCtxResolveComment',
          command: 'comment-resolve',
          payload: id,
          disabled: off,
        },
  ]
}

export interface NoteTarget {
  kind: 'footnote' | 'endnote'
  id: string
  /** the right-click landed on the note's text, or on its reference mark in the body */
  at: 'text' | 'ref'
}

/** Word's note items: Go to Footnote / Endnote, Note Options…, Convert to Endnote / Footnote */
export function noteRows(note: NoteTarget, canEdit: boolean): CtxRow[] {
  const foot = note.kind === 'footnote'
  return [
    {
      id: 'note-goto',
      labelKey: foot ? 'appCtxGoToFootnote' : 'appCtxGoToEndnote',
      command: 'note-goto',
      // from the text to the mark, from the mark to the text
      payload: `${note.kind}:${note.at === 'text' ? 'ref' : 'text'}:${note.id}`,
    },
    {
      id: 'note-options',
      labelKey: 'appCtxNoteOptions',
      command: 'insert-footnote-dialog',
      disabled: !canEdit,
    },
    {
      id: 'note-convert',
      labelKey: foot ? 'appCtxConvertToEndnote' : 'appCtxConvertToFootnote',
      command: 'note-convert',
      payload: `${note.kind}:${note.id}`,
      disabled: !canEdit,
    },
  ]
}

const INSERTIONS: ReadonlySet<RevisionRange['kind']> = new Set([
  'ins',
  'paraMarkIns',
  'rowIns',
  'cellIns',
  'blockIns',
])
const DELETIONS: ReadonlySet<RevisionRange['kind']> = new Set([
  'del',
  'paraMarkDel',
  'rowDel',
  'cellDel',
  'blockDel',
])

/** Word names the change: Accept Insertion / Reject Deletion; other kinds keep Accept Change */
export function revisionLabelKeys(kind: RevisionRange['kind']): {
  accept: StringKey
  reject: StringKey
} {
  if (INSERTIONS.has(kind))
    return { accept: 'appCtxAcceptInsertion', reject: 'appCtxRejectInsertion' }
  if (DELETIONS.has(kind)) return { accept: 'appCtxAcceptDeletion', reject: 'appCtxRejectDeletion' }
  return { accept: 'appCtxAcceptChange', reject: 'appCtxRejectChange' }
}

/** the heading under the pointer: where it is and whether its body is folded or can fold */
export interface HeadingFoldTarget {
  pos: number
  collapsed: boolean
  foldable: boolean
}

/** Word's Expand/Collapse submenu on a heading */
export function headingFoldRows(target: HeadingFoldTarget): CtxRow[] {
  return [
    {
      id: 'heading-fold',
      labelKey: 'appHeadingExpandCollapse',
      children: [
        {
          id: 'heading-expand',
          labelKey: 'appHeadingExpand',
          command: 'heading-collapse',
          payload: `expand@${target.pos}`,
          disabled: !target.collapsed,
        },
        {
          id: 'heading-collapse',
          labelKey: 'appHeadingCollapse',
          command: 'heading-collapse',
          payload: `collapse@${target.pos}`,
          disabled: target.collapsed || !target.foldable,
        },
        {
          id: 'heading-expand-all',
          labelKey: 'appHeadingExpandAll',
          command: 'heading-collapse',
          payload: 'expand-all',
        },
        {
          id: 'heading-collapse-all',
          labelKey: 'appHeadingCollapseAll',
          command: 'heading-collapse',
          payload: 'collapse-all',
        },
      ],
    },
  ]
}
