/**
 * Live search while the Find dialog is in replace mode.
 *
 * Univer's FindReplaceModel only researches on typing while the replace
 * panel is hidden; with it revealed, typing merely stores
 * `inputtingFindString` and stops the session, so Replace / Replace All stay
 * disabled until the user clicks Find. Expanding first and typing afterwards
 * therefore looks broken. This mirrors what the Find button does (commit the
 * input as findString, then start) whenever the input or a search option
 * changes in replace mode, so the buttons light up as soon as there is a hit.
 */
import type { IFindReplaceState } from '@univerjs/find-replace'
import { debounceTime, distinctUntilChanged, map, type Observable } from 'rxjs'

/// Must exceed Univer's 200ms stateUpdates$ throttle: its trailing emission
/// runs _stopSearching for replace-mode typing and would kill a session we
/// started earlier.
export const REPLACE_AUTOSEARCH_DELAY_MS = 300

export interface ReplaceAutoSearchService {
  readonly state$: Observable<IFindReplaceState>
  changeFindString(findString: string): void
  find(): void
}

type SearchOptions = Pick<
  IFindReplaceState,
  'caseSensitive' | 'matchesTheWholeCell' | 'findBy' | 'findScope' | 'findDirection'
>

const OPTION_KEYS: readonly (keyof SearchOptions)[] = [
  'caseSensitive',
  'matchesTheWholeCell',
  'findBy',
  'findScope',
  'findDirection',
]

/** The slice of dialog state that decides whether to search. */
export type SearchInputs = SearchOptions &
  Pick<IFindReplaceState, 'revealed' | 'replaceRevealed' | 'findString' | 'inputtingFindString'>

const INPUT_KEYS: readonly (keyof SearchInputs)[] = [
  'revealed',
  'replaceRevealed',
  'findString',
  'inputtingFindString',
  ...OPTION_KEYS,
]

export function pickSearchInputs(state: IFindReplaceState): SearchInputs {
  return {
    revealed: state.revealed,
    replaceRevealed: state.replaceRevealed,
    findString: state.findString,
    inputtingFindString: state.inputtingFindString,
    caseSensitive: state.caseSensitive,
    matchesTheWholeCell: state.matchesTheWholeCell,
    findBy: state.findBy,
    findScope: state.findScope,
    findDirection: state.findDirection,
  }
}

export function sameSearchInputs(a: SearchInputs, b: SearchInputs): boolean {
  return INPUT_KEYS.every((key) => a[key] === b[key])
}

export type AutoSearchAction = 'commit-and-search' | 'search' | null

/**
 * Decides whether a debounced state snapshot needs a search. Option toggles
 * in replace mode also stop the session (findString stays equal to the
 * input), so they are detected by diffing against the previous snapshot.
 * A finished Replace All stops the session too, but changes neither the
 * strings nor the options, so it does not trigger a re-search (and its
 * "no match" toast).
 */
export function planAutoSearch(prev: SearchInputs | null, next: SearchInputs): AutoSearchAction {
  if (!next.revealed || !next.replaceRevealed) return null
  if (next.inputtingFindString.trim() === '') return null
  if (next.inputtingFindString !== next.findString) return 'commit-and-search'
  if (prev && OPTION_KEYS.some((key) => prev[key] !== next[key])) return 'search'
  return null
}

export function installReplaceAutoSearch(
  service: ReplaceAutoSearchService,
  delayMs = REPLACE_AUTOSEARCH_DELAY_MS,
): { dispose(): void } {
  let prev: SearchInputs | null = null
  // Only the search inputs feed the debounce: a running scan (notably the
  // lazy file-backed one) keeps bumping matchesCount, and letting those
  // emissions reset the timer would starve later typing.
  const sub = service.state$
    .pipe(map(pickSearchInputs), distinctUntilChanged(sameSearchInputs), debounceTime(delayMs))
    .subscribe((inputs) => {
      const action = planAutoSearch(prev, inputs)
      prev = inputs
      if (!action) return
      if (action === 'commit-and-search') service.changeFindString(inputs.inputtingFindString)
      service.find()
    })
  return { dispose: () => sub.unsubscribe() }
}
