/**
 * Picture/shape right-click dialogs: Edit Alt Text, Size and Properties,
 * Hyperlink. Each hands a ShapeEditChanges back to the app.
 */
import { useState } from 'react'

import type { WorkbookVisualObject } from '../shared/desktop-api'
import { useI18n } from './i18n/locale'
import { useModalDialog } from './modal-dialog'
import type { VisualDialogKind } from './visual-arrange-actions'
import { EMU_PER_PIXEL, type ShapeEditChanges } from './WorkbookVisuals'

const EMU_PER_CM = 360_000
const pxToCm = (px: number): number => (px * EMU_PER_PIXEL) / EMU_PER_CM
const cmToPx = (cm: number): number => (cm * EMU_PER_CM) / EMU_PER_PIXEL
const format = (value: number): string => (Math.round(value * 100) / 100).toString()

/// Size and Properties commits size/rotation and placement in one apply so
/// they land as a single undo step.
export interface SizeDialogResult {
  size?: { width: number; height: number; rotation: number }
  editAs?: NonNullable<WorkbookVisualObject['editAs']>
}

export interface VisualDialogProps {
  readonly kind: VisualDialogKind
  readonly visual: WorkbookVisualObject
  /// Unzoomed frame in sheet pixels (the true, unrotated frame).
  readonly frame: { readonly width: number; readonly height: number } | null
  readonly onApply: (changes: ShapeEditChanges | SizeDialogResult) => void
  readonly onClose: () => void
}

export function VisualDialog(props: VisualDialogProps): React.JSX.Element {
  if (props.kind === 'alt-text') return <AltTextDialog {...props} />
  if (props.kind === 'hyperlink') return <HyperlinkDialog {...props} />
  return <SizePropertiesDialog {...props} />
}

function Frame({
  title,
  onClose,
  onApply,
  children,
}: {
  readonly title: string
  readonly onClose: () => void
  readonly onApply: () => void
  readonly children: React.ReactNode
}): React.JSX.Element {
  const { t } = useI18n()
  const modal = useModalDialog(onClose)
  return (
    <div className="dialog-backdrop" onClick={onClose}>
      <div
        {...modal}
        className="format-cells-dialog visual-dialog"
        role="dialog"
        aria-label={title}
        onClick={(event) => event.stopPropagation()}
        onKeyDown={(event) => {
          modal.onKeyDown(event)
          if (
            event.key === 'Enter' &&
            !(event.target instanceof HTMLButtonElement) &&
            !(event.target instanceof HTMLTextAreaElement)
          ) {
            event.preventDefault()
            onApply()
          }
        }}
      >
        <header>{title}</header>
        <section className="dialog-body">{children}</section>
        <div className="dialog-actions">
          <button className="secondary" onClick={onClose}>
            {t('dlgCancel')}
          </button>
          <button className="primary-action" onClick={onApply}>
            {t('dlgOk')}
          </button>
        </div>
      </div>
    </div>
  )
}

function AltTextDialog({ visual, onApply, onClose }: VisualDialogProps): React.JSX.Element {
  const { t } = useI18n()
  const [text, setText] = useState(visual.altText ?? '')
  const apply = (): void => {
    if (text.trim() !== (visual.altText ?? '')) onApply({ altText: text.trim() })
    onClose()
  }
  return (
    <Frame title={t('dlgAltTextTitle')} onClose={onClose} onApply={apply}>
      <p className="dialog-note">{t('dlgAltTextHint')}</p>
      <textarea
        className="visual-dialog-textarea"
        autoFocus
        rows={4}
        value={text}
        onChange={(event) => setText(event.target.value)}
      />
    </Frame>
  )
}

function HyperlinkDialog({ visual, onApply, onClose }: VisualDialogProps): React.JSX.Element {
  const { t } = useI18n()
  const [address, setAddress] = useState(visual.hyperlink ?? '')
  const apply = (): void => {
    const next = address.trim()
    if (next !== (visual.hyperlink ?? '')) onApply({ hyperlink: next })
    onClose()
  }
  return (
    <Frame title={t('dlgHyperlinkTitle')} onClose={onClose} onApply={apply}>
      <label className="visual-dialog-field">
        <span>{t('dlgHyperlinkAddress')}</span>
        <input
          type="url"
          autoFocus
          value={address}
          placeholder="https://"
          onChange={(event) => setAddress(event.target.value)}
        />
      </label>
      {visual.hyperlink && (
        <button
          type="button"
          className="secondary visual-dialog-remove"
          onClick={() => {
            onApply({ hyperlink: '' })
            onClose()
          }}
        >
          {t('dlgHyperlinkRemove')}
        </button>
      )}
    </Frame>
  )
}

function SizePropertiesDialog({
  visual,
  frame,
  onApply,
  onClose,
}: VisualDialogProps): React.JSX.Element {
  const { t } = useI18n()
  const canResize = frame !== null && visual.kind !== 'chart'
  const [width, setWidth] = useState(format(pxToCm(frame?.width ?? 0)))
  const [height, setHeight] = useState(format(pxToCm(frame?.height ?? 0)))
  const [lockAspect, setLockAspect] = useState(visual.kind === 'image')
  const [rotation, setRotation] = useState(format(visual.rotation ?? 0))
  // Session objects save as two-cell anchors; file objects without editAs
  // are one-cell/absolute anchors, which carry no placement mode.
  const initialEditAs = visual.editAs ?? (visual.id.startsWith('added-') ? 'twoCell' : undefined)
  const [editAs, setEditAs] = useState(initialEditAs)
  const [error, setError] = useState(false)
  const ratio = frame && frame.height > 0 ? frame.width / frame.height : 1

  const onWidth = (value: string): void => {
    setWidth(value)
    setError(false)
    const cm = Number(value)
    if (lockAspect && Number.isFinite(cm) && cm > 0) setHeight(format(cm / ratio))
  }
  const onHeight = (value: string): void => {
    setHeight(value)
    setError(false)
    const cm = Number(value)
    if (lockAspect && Number.isFinite(cm) && cm > 0) setWidth(format(cm * ratio))
  }

  const apply = (): void => {
    const nextWidth = cmToPx(Number(width))
    const nextHeight = cmToPx(Number(height))
    const nextRotation = Number(rotation)
    if (
      (canResize && (!Number.isFinite(nextWidth) || nextWidth < 1)) ||
      (canResize && (!Number.isFinite(nextHeight) || nextHeight < 1)) ||
      !Number.isFinite(nextRotation)
    ) {
      setError(true)
      return
    }
    const result: SizeDialogResult = {}
    if (editAs !== undefined && editAs !== initialEditAs) result.editAs = editAs
    if (canResize && frame) {
      const sizeChanged =
        Math.abs(nextWidth - frame.width) > 0.5 || Math.abs(nextHeight - frame.height) > 0.5
      if (sizeChanged || nextRotation !== (visual.rotation ?? 0)) {
        result.size = { width: nextWidth, height: nextHeight, rotation: nextRotation }
      }
    }
    if (result.size || result.editAs) onApply(result)
    onClose()
  }

  const placement = (value: NonNullable<WorkbookVisualObject['editAs']>, label: string) => (
    <label className="dialog-check">
      <input
        type="radio"
        name="visual-edit-as"
        checked={editAs === value}
        onChange={() => setEditAs(value)}
      />
      {label}
    </label>
  )

  return (
    <Frame title={t('dlgSizeTitle')} onClose={onClose} onApply={apply}>
      {canResize && (
        <fieldset className="visual-dialog-size">
          <legend>{t('dlgSizeSection')}</legend>
          <label className="visual-dialog-field">
            <span>{t('dlgSizeHeight')}</span>
            <input
              type="number"
              step="0.01"
              min="0.01"
              autoFocus
              value={height}
              onChange={(event) => onHeight(event.target.value)}
            />
            <span>cm</span>
          </label>
          <label className="visual-dialog-field">
            <span>{t('dlgSizeWidth')}</span>
            <input
              type="number"
              step="0.01"
              min="0.01"
              value={width}
              onChange={(event) => onWidth(event.target.value)}
            />
            <span>cm</span>
          </label>
          <label className="visual-dialog-field">
            <span>{t('dlgSizeRotation')}</span>
            <input
              type="number"
              step="1"
              value={rotation}
              onChange={(event) => {
                setRotation(event.target.value)
                setError(false)
              }}
            />
            <span>°</span>
          </label>
          <label className="dialog-check">
            <input
              type="checkbox"
              checked={lockAspect}
              onChange={(event) => setLockAspect(event.target.checked)}
            />
            {t('dlgSizeLockAspect')}
          </label>
        </fieldset>
      )}
      {editAs !== undefined && (
        <fieldset className="visual-dialog-placement">
          <legend>{t('dlgSizeProperties')}</legend>
          {placement('twoCell', t('dlgSizeMoveAndSize'))}
          {placement('oneCell', t('dlgSizeMoveOnly'))}
          {placement('absolute', t('dlgSizeFixed'))}
        </fieldset>
      )}
      {error && (
        <p className="dialog-note dialog-error" role="alert">
          {t('dlgSizeInvalid')}
        </p>
      )}
    </Frame>
  )
}
