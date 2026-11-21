import { useState } from 'react'
import type { Editor } from '@tiptap/core'
import { useI18n, type StringKey } from '../i18n/locale'
import { effectiveAlignFrame, type AlignFrame, type AlignMode } from '../editor/object-align'
import { selectedObjectPositions } from '../editor/object-multi-select'
import { IconAlignObjects, IconCaret } from './icons'
import { useRunCommand } from './ribbon-command'
import { BIG, toggleDropdown, type SetDropdown } from './ribbon-tabs'

/** Word's Align menu, in its order; `rule` starts a new block */
const ALIGN_ITEMS: Array<{
  mode: AlignMode
  labelKey: StringKey
  rule?: true
  distribute?: true
}> = [
  { mode: 'left', labelKey: 'ribbonObjAlignLeft' },
  { mode: 'center', labelKey: 'ribbonObjAlignCenter' },
  { mode: 'right', labelKey: 'ribbonObjAlignRight' },
  { mode: 'top', labelKey: 'ribbonObjAlignTop', rule: true },
  { mode: 'middle', labelKey: 'ribbonObjAlignMiddle' },
  { mode: 'bottom', labelKey: 'ribbonObjAlignBottom' },
  { mode: 'distribute-h', labelKey: 'ribbonObjDistributeH', rule: true, distribute: true },
  { mode: 'distribute-v', labelKey: 'ribbonObjDistributeV', distribute: true },
]

const ALIGN_FRAME_ITEMS: Array<{ frame: AlignFrame; labelKey: StringKey }> = [
  { frame: 'page', labelKey: 'ribbonObjAlignToPage' },
  { frame: 'margin', labelKey: 'ribbonObjAlignToMargin' },
  { frame: 'objects', labelKey: 'ribbonObjAlignSelected' },
]

/**
 * Arrange > Align, the same on Layout and on the Shape Format and Picture
 * Format tabs: `look` is the button the tab draws it as.
 */
export function AlignObjectsMenu({
  editor,
  hasDoc,
  dropdown,
  setDropdown,
  look,
  opensEmpty = false,
}: {
  editor: Editor
  hasDoc: boolean
  dropdown: string | null
  setDropdown: SetDropdown
  look: 'big' | 'tool'
  /** Word's Layout and Picture Format keep the button live with nothing to line up, its items greyed */
  opensEmpty?: boolean
}) {
  const { t } = useI18n()
  const runCommand = useRunCommand()
  const objectCount = selectedObjectPositions(editor.state).length
  // Word: several objects line up against each other unless the user picked the page or the margin
  const [alignFrame, setAlignFrame] = useState<AlignFrame | null>(null)
  const alignAgainst = effectiveAlignFrame(alignFrame ?? 'objects', objectCount)
  const canDistribute = alignAgainst !== 'objects' || objectCount > 2
  const disabled = !hasDoc || (objectCount === 0 && !opensEmpty)
  return (
    <div className="rb-split-wrap">
      {look === 'big' ? (
        <button
          className="rb-big rb-align-objects"
          disabled={disabled}
          data-tip={t('ribbonAlignObjects')}
          onClick={() => toggleDropdown(setDropdown, 'arrange-align')}
        >
          <span className="rb-big-icon">
            <IconAlignObjects size={BIG} />
            <IconCaret />
          </span>
          <span>{t('ribbonAlignObjects')}</span>
        </button>
      ) : (
        <button
          className="table-tool-button rb-align-objects"
          disabled={disabled}
          data-tip={t('ribbonAlignObjects')}
          aria-label={t('ribbonAlignObjects')}
          aria-haspopup="menu"
          aria-expanded={dropdown === 'arrange-align'}
          onClick={() => toggleDropdown(setDropdown, 'arrange-align')}
        >
          <IconAlignObjects size={17} />
          <IconCaret />
        </button>
      )}
      {dropdown === 'arrange-align' && (
        <div data-rb-panel="" className="layout-menu">
          {ALIGN_ITEMS.flatMap((item) => [
            item.rule && <div key={`${item.mode}-rule`} className="rb-menu-sep" />,
            <button
              key={item.mode}
              disabled={objectCount === 0 || (item.distribute && !canDistribute)}
              onClick={() => {
                runCommand('arrange-align', `${item.mode}:${alignAgainst}`)
                setDropdown(() => null)
              }}
            >
              <b>{t(item.labelKey)}</b>
            </button>,
          ])}
          <div className="rb-menu-sep" />
          {ALIGN_FRAME_ITEMS.map((item) => (
            <button
              key={item.frame}
              role="menuitemradio"
              aria-checked={alignAgainst === item.frame}
              className={alignAgainst === item.frame ? 'active' : ''}
              disabled={item.frame === 'objects' && objectCount < 2}
              // command-exempt: what Align lines up against is the dropdown's own setting
              onClick={() => setAlignFrame(item.frame)}
            >
              <b>{t(item.labelKey)}</b>
            </button>
          ))}
        </div>
      )}
    </div>
  )
}
