import type { Editor } from '@tiptap/core'
import { type TextboxDisplay } from '@genoffice/docx-engine'
import { WRAP_OPTIONS } from './ContextMenu'
import { shapeWrapOf } from '../editor/floating-z-order'
import { canGroupObjects, canUngroupObjects } from '../editor/object-group'
import { useI18n } from '../i18n/locale'
import type { ActionCommandId } from '../../shared/commands'
import { NONE } from '../commands/menu-commands'
import { RbCommand, useRunCommand } from './ribbon-command'
import { AlignObjectsMenu } from './ribbon-align-menu'
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
  // enabled when a floating object (image/textbox) is selected
  const protAttrs = editor.getAttributes('docProtected')
  const isImage = protAttrs?.blockType === 'image' && !!protAttrs.imageDataUrl
  const isFloatingBox =
    Array.isArray(protAttrs?.textboxes) && (protAttrs.textboxes as unknown[]).length > 0
  const canWrap = hasDoc && (isImage || isFloatingBox)
  // position presets write imagePosH/imagePosV (and clear the drag offsets):
  // original images patch through imagePatchOf, newly inserted ones save the
  // same margin-relative align pair via NewImage.posH/posV
  const canPosition = hasDoc && isImage
  const canGroup = hasDoc && canGroupObjects(editor.state)
  const canUngroup = hasDoc && canUngroupObjects(editor.state)
  const boxes = protAttrs?.textboxes as TextboxDisplay[] | undefined
  const currentWrap = boxes?.length
    ? shapeWrapOf(boxes[0])
    : ((protAttrs?.imageWrap as string | null) ?? null)

  return (
    <div className="ribbon-group">
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
                        protAttrs?.imagePosH === h && protAttrs?.imagePosV === v ? ' active' : ''
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
        <button
          className="rb-big"
          disabled={!canWrap}
          data-tip={t('appBringForward')}
          onClick={() => runCommand('arrange-order', 'forward')}
        >
          <span className="rb-big-icon">
            <IconBringForward size={BIG} />
          </span>
          <span>{t('appBringForward')}</span>
        </button>
        <button
          className="rb-big"
          disabled={!canWrap}
          data-tip={t('appSendBackward')}
          onClick={() => runCommand('arrange-order', 'backward')}
        >
          <span className="rb-big-icon">
            <IconSendBackward size={BIG} />
          </span>
          <span>{t('appSendBackward')}</span>
        </button>
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
            disabled={!hasDoc || !isImage}
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
              <button onClick={() => pick('picture-rotate', '90')}>{t('ribbonRotateRight')}</button>
              <button onClick={() => pick('picture-rotate', '-90')}>{t('ribbonRotateLeft')}</button>
              <button onClick={() => pick('picture-flip', 'v')}>{t('ribbonFlipV')}</button>
              <button onClick={() => pick('picture-flip', 'h')}>{t('ribbonFlipH')}</button>
            </div>
          )}
        </div>
      </div>
      <div className="ribbon-group-label">{t('ribbonGroupArrange')}</div>
    </div>
  )
}
