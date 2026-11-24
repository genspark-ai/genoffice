import type { ReactNode } from 'react'
import type { Editor } from '@tiptap/core'
import { WRAP_OPTIONS } from './ContextMenu'
import { getActiveSubEditor } from '../editor/active-editor'
import { arrangeStateOf } from '../editor/arrange-target'
import { canGroupObjects, canUngroupObjects } from '../editor/object-group'
import { useI18n, type StringKey } from '../i18n/locale'
import type { ActionCommandId } from '../../shared/commands'
import { NONE } from '../commands/menu-commands'
import { RbCommand, useRunCommand } from './ribbon-command'
import { AlignObjectsMenu } from './ribbon-align-menu'
import { RbGroupMenu } from './ribbon-group-menu'
import {
  IconCaret,
  IconPosition,
  IconWrapText,
  IconBringForward,
  IconSendBackward,
  IconSelectionPane,
  IconGroupObjects,
  IconRotateRight,
} from './icons'
import { BIG, toggleDropdown, type SetDropdown } from './ribbon-tabs'

/** Word's Arrange group, the same on Layout and on the tab of the selected object */
export function ArrangeGroup({
  editor,
  hasDoc,
  dropdown,
  setDropdown,
}: {
  editor: Editor
  hasDoc: boolean
  dropdown: string | null
  setDropdown: SetDropdown
}) {
  const { t } = useI18n()
  const runCommand = useRunCommand()
  const pick = (id: ActionCommandId, payload?: string) => {
    runCommand(id, payload)
    setDropdown(() => null)
  }
  const state = arrangeStateOf(editor, getActiveSubEditor())
  const canWrap = hasDoc && state.canWrap
  const canPosition = hasDoc && state.canPosition
  const canOrder = hasDoc && state.canOrder
  const canGroup = hasDoc && canGroupObjects(editor.state)
  const canUngroup = hasDoc && canUngroupObjects(editor.state)
  const currentWrap = state.wrap
  // Word's two stacking buttons each open the step and the jump to the end
  const order = (
    name: 'arrange-forward' | 'arrange-backward',
    step: [StringKey, string],
    end: [StringKey, string],
    icon: ReactNode,
  ) => (
    <div className="rb-split-wrap">
      <button
        className="rb-big"
        disabled={!canOrder}
        data-tip={t(step[0])}
        onClick={() => toggleDropdown(setDropdown, name)}
      >
        <span className="rb-big-icon">
          {icon}
          <IconCaret />
        </span>
        <span>{t(step[0])}</span>
      </button>
      {dropdown === name && (
        <div data-rb-panel="" className="layout-menu">
          {[step, end].map(([label, payload]) => (
            <button key={payload} onClick={() => pick('arrange-order', payload)}>
              <b>{t(label)}</b>
            </button>
          ))}
        </div>
      )}
    </div>
  )

  return (
    <div className="ribbon-group" data-menu={1001}>
      <RbGroupMenu name={t('ribbonGroupArrange')}>
        <IconPosition size={BIG} />
      </RbGroupMenu>
      <div className="ribbon-group-items">
        <div className="rb-split-wrap">
          <button
            className="rb-big"
            disabled={!canPosition}
            data-tip={t('ribbonPosition')}
            onClick={() => toggleDropdown(setDropdown, 'arrange-pos')}
          >
            <span className="rb-big-icon">
              <IconPosition size={BIG} />
              <IconCaret />
            </span>
            <span>{t('ribbonPosition')}</span>
          </button>
          {dropdown === 'arrange-pos' && (
            <div data-rb-panel="" className="layout-menu">
              <button
                className={currentWrap ? '' : 'active'}
                onClick={() => pick('picture-wrap', NONE)}
              >
                <b>{t('appWrapInline')}</b>
              </button>
              <div className="pos-grid">
                {(['top', 'center', 'bottom'] as const).map((v) =>
                  (['left', 'center', 'right'] as const).map((h) => (
                    <button
                      key={`${v}-${h}`}
                      className={`pos-cell ph-${h} pv-${v}${
                        state.posH === h && state.posV === v ? ' active' : ''
                      }`}
                      data-tip={t('ribbonPosition')}
                      aria-label={t('ribbonPosition')}
                      onClick={() => pick('picture-position', `${h},${v}`)}
                    >
                      <span className="pos-dot" />
                    </button>
                  )),
                )}
              </div>
            </div>
          )}
        </div>
        <div className="rb-split-wrap">
          <button
            className="rb-big"
            disabled={!canWrap}
            data-tip={t('ribbonWrapText')}
            onClick={() => toggleDropdown(setDropdown, 'arrange-wrap')}
          >
            <span className="rb-big-icon">
              <IconWrapText size={BIG} />
              <IconCaret />
            </span>
            <span>{t('ribbonWrapText')}</span>
          </button>
          {dropdown === 'arrange-wrap' && (
            <div data-rb-panel="" className="layout-menu">
              {WRAP_OPTIONS.map((opt) => (
                <button
                  key={String(opt.value)}
                  className={currentWrap === opt.value ? 'active' : ''}
                  onClick={() => pick('picture-wrap', opt.value ?? NONE)}
                >
                  <b>{t(opt.labelKey)}</b>
                </button>
              ))}
            </div>
          )}
        </div>
        {order(
          'arrange-forward',
          ['appBringForward', 'forward'],
          ['appBringToFront', 'front'],
          <IconBringForward size={BIG} />,
        )}
        {order(
          'arrange-backward',
          ['appSendBackward', 'backward'],
          ['appSendToBack', 'back'],
          <IconSendBackward size={BIG} />,
        )}
        <RbCommand cmd="format-selection-pane" look="big">
          <IconSelectionPane size={BIG} />
        </RbCommand>
        <AlignObjectsMenu
          editor={editor}
          hasDoc={hasDoc}
          dropdown={dropdown}
          setDropdown={setDropdown}
          look="big"
          opensEmpty
        />
        <div className="rb-split-wrap">
          <button
            className="rb-big"
            data-icon={1}
            disabled={!canGroup && !canUngroup}
            data-tip={t('ribbonArrangeGroup')}
            onClick={() => toggleDropdown(setDropdown, 'arrange-group')}
          >
            <span className="rb-big-icon">
              <IconGroupObjects size={BIG} />
              <IconCaret />
            </span>
            <span>{t('ribbonArrangeGroup')}</span>
          </button>
          {dropdown === 'arrange-group' && (
            <div data-rb-panel="" className="layout-menu">
              <button disabled={!canGroup} onClick={() => pick('arrange-group', 'group')}>
                <b>{t('ribbonArrangeGroup')}</b>
              </button>
              <button disabled={!canUngroup} onClick={() => pick('arrange-group', 'ungroup')}>
                <b>{t('ribbonArrangeUngroup')}</b>
              </button>
            </div>
          )}
        </div>
        <div className="rb-split-wrap">
          <button
            className="rb-big"
            data-icon={1}
            disabled={!hasDoc || !(state.canRotate || state.canFlip)}
            data-tip={t('ribbonArrangeRotate')}
            onClick={() => toggleDropdown(setDropdown, 'arrange-rotate')}
          >
            <span className="rb-big-icon">
              <IconRotateRight size={BIG} />
              <IconCaret />
            </span>
            <span>{t('ribbonArrangeRotate')}</span>
          </button>
          {dropdown === 'arrange-rotate' && (
            <div data-rb-panel="" className="layout-menu">
              <button disabled={!state.canRotate} onClick={() => pick('picture-rotate', '90')}>
                {t('ribbonRotateRight')}
              </button>
              <button disabled={!state.canRotate} onClick={() => pick('picture-rotate', '-90')}>
                {t('ribbonRotateLeft')}
              </button>
              <button disabled={!state.canFlip} onClick={() => pick('picture-flip', 'v')}>
                {t('ribbonFlipV')}
              </button>
              <button disabled={!state.canFlip} onClick={() => pick('picture-flip', 'h')}>
                {t('ribbonFlipH')}
              </button>
            </div>
          )}
        </div>
      </div>
      <div className="ribbon-group-label">{t('ribbonGroupArrange')}</div>
    </div>
  )
}
