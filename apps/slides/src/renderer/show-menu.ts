/**
 * Right-click menu of the slide show window, in PowerPoint's item order. Pure so the disabled
 * states can be unit-tested without mounting the show.
 */
import type { CtxItem } from './components/ContextMenu'
import type { TFunc } from './i18n/locale'
import type { ShowScreen } from './show-keys'

export interface ShowMenuState {
  /** Current play-order position + total length */
  pos: number
  count: number
  /** Whether the "end of show" screen is up */
  ended: boolean
  /** Whether the show still has an in-page animation step on the current slide */
  pending: boolean
  hasLastViewed: boolean
  screen: ShowScreen
}

export interface ShowMenuActions {
  next: () => void
  prev: () => void
  lastViewed: () => void
  seeAll: () => void
  setScreen: (screen: Exclude<ShowScreen, 'none'>) => void
  end: () => void
}

export function buildShowMenu(
  t: TFunc,
  s: ShowMenuState,
  a: ShowMenuActions,
): Array<CtxItem | null> {
  return [
    {
      label: t('paneShowMenuNext'),
      disabled: s.ended || (s.pos >= s.count - 1 && !s.pending),
      onClick: a.next,
    },
    { label: t('paneShowMenuPrev'), disabled: !s.ended && s.pos <= 0, onClick: a.prev },
    { label: t('paneShowMenuLastViewed'), disabled: !s.hasLastViewed, onClick: a.lastViewed },
    { label: t('paneShowMenuSeeAll'), onClick: a.seeAll },
    null,
    {
      label: t('paneShowMenuScreen'),
      sub: [
        {
          label: t('paneShowMenuBlack'),
          hint: 'B',
          checked: s.screen === 'black',
          onClick: () => a.setScreen('black'),
        },
        {
          label: t('paneShowMenuWhite'),
          hint: 'W',
          checked: s.screen === 'white',
          onClick: () => a.setScreen('white'),
        },
      ],
    },
    null,
    { label: t('paneShowMenuEnd'), onClick: a.end },
  ]
}
