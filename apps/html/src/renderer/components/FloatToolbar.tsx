import { useRef, useState } from 'react'
import type { Ref } from 'react'
import { useDismissablePopover } from '@genoffice/ui'
import { useI18n } from '../i18n/locale'
import type { ComputedSnapshot } from '../preview/inspector-protocol'
import { ColorPop } from './ColorField'
import { backgroundPickStyles } from '../document/background-style'
import {
  flipImage,
  imageAlignOf,
  imageAlignStyles,
  parseImageTransform,
  rotateImage,
  transformStyles,
} from '../document/image-style'
import {
  IconAlignCenter,
  IconAlignLeft,
  IconAlignRight,
  IconCopy,
  IconCrop,
  IconFlipH,
  IconFlipV,
  IconMoveDown,
  IconMoveUp,
  IconRemoveBg,
  IconReplacePicture,
  IconRotateLeft,
  IconRotateRight,
  IconSliders,
  IconSparkle,
  IconTrash,
} from './icons'

export type StyleEdit = (styles: Record<string, string | null>) => void

interface Props {
  barRef?: Ref<HTMLDivElement>
  left: number
  top: number
  computed: ComputedSnapshot
  /** live pokes not yet written to the source; newer than the snapshot, which round-trips through the frame */
  pending: Record<string, string | null>
  /** the element has one editable text run */
  canEditText: boolean
  /** the element can carry an AI instruction (structural nodes cannot) */
  canAskAi: boolean
  onStyle: StyleEdit
  onEditText: () => void
  onReplaceImage: () => void
  onCropImage: () => void
  onCutoutImage: () => void
  onAskAi: () => void
  tag: string
  onMove: (dir: -1 | 1) => void
  onDuplicate: () => void
  onDelete: () => void
  panelOpen: boolean
  onTogglePanel: () => void
}

const FONT_STEP = 2
const MIN_FONT = 8

export function fontSizeNumber(computed: ComputedSnapshot): number {
  const n = Number(computed.fontSize)
  return Number.isFinite(n) && n > 0 ? n : 16
}

const stop = (e: React.MouseEvent) => e.preventDefault()

/** middle of the bar for text-bearing elements: B / I, font size, colours, text-align */
function TextTools(p: Props) {
  const { t } = useI18n()
  const [colorOpen, setColorOpen] = useState(false)
  const [colorTarget, setColorTarget] = useState<'text' | 'background'>('text')
  const colorRef = useRef<HTMLSpanElement>(null)
  const popRef = useRef<HTMLDivElement>(null)
  useDismissablePopover(colorOpen, () => setColorOpen(false), {
    inside: () => [colorRef.current, popRef.current],
  })
  const bold = Number(p.computed.fontWeight) >= 600
  const italic = p.computed.fontStyle === 'italic'
  const size = fontSizeNumber(p.computed)
  return (
    <>
      <button
        type="button"
        className={`hx-float-btn${bold ? ' on' : ''}`}
        data-tip={t('fmtBold')}
        aria-label={t('fmtBold')}
        aria-pressed={bold}
        onMouseDown={stop}
        onClick={() => p.onStyle({ 'font-weight': bold ? '400' : '700' })}
      >
        <b>B</b>
      </button>
      <button
        type="button"
        className={`hx-float-btn${italic ? ' on' : ''}`}
        data-tip={t('fmtItalic')}
        aria-label={t('fmtItalic')}
        aria-pressed={italic}
        onMouseDown={stop}
        onClick={() => p.onStyle({ 'font-style': italic ? 'normal' : 'italic' })}
      >
        <i>I</i>
      </button>
      <span className="hx-float-sep" />
      <button
        type="button"
        className="hx-float-btn"
        data-tip={t('fontSizeDec')}
        aria-label={t('fontSizeDec')}
        onMouseDown={stop}
        onClick={() => p.onStyle({ 'font-size': `${Math.max(MIN_FONT, size - FONT_STEP)}px` })}
      >
        −
      </button>
      <span className="hx-float-size" aria-label={t('fontSize')}>
        {size}
      </span>
      <button
        type="button"
        className="hx-float-btn"
        data-tip={t('fontSizeInc')}
        aria-label={t('fontSizeInc')}
        onMouseDown={stop}
        onClick={() => p.onStyle({ 'font-size': `${size + FONT_STEP}px` })}
      >
        +
      </button>
      <span className="hx-float-sep" />
      <span className="hx-float-pop-host" ref={colorRef}>
        <button
          type="button"
          className={`hx-float-btn${colorOpen ? ' on' : ''}`}
          data-tip={t('color')}
          aria-label={t('color')}
          aria-expanded={colorOpen}
          onMouseDown={stop}
          onClick={() => setColorOpen((v) => !v)}
        >
          <span className="hx-float-dot" style={{ background: p.computed.color || '#000' }} />
        </button>
        {colorOpen && (
          <ColorPop
            popRef={popRef}
            track={`${p.left}:${p.top}`}
            anchor={colorRef.current}
            value={colorTarget === 'text' ? p.computed.color : p.computed.background}
            none={colorTarget === 'background'}
            onPick={(hex) =>
              colorTarget === 'text'
                ? hex && p.onStyle({ color: hex })
                : p.onStyle(backgroundPickStyles(hex, p.computed.backgroundImage))
            }
            onClose={() => setColorOpen(false)}
          >
            <div className="hx-color-tabs" role="tablist">
              {(['text', 'background'] as const).map((target) => (
                <button
                  key={target}
                  type="button"
                  role="tab"
                  className={colorTarget === target ? 'on' : ''}
                  aria-selected={colorTarget === target}
                  onClick={() => setColorTarget(target)}
                >
                  {t(target === 'text' ? 'textLabel' : 'background')}
                </button>
              ))}
            </div>
          </ColorPop>
        )}
      </span>
      <span className="hx-float-sep" />
      {(
        [
          ['left', 'alignLeft', IconAlignLeft],
          ['center', 'alignCenter', IconAlignCenter],
          ['right', 'alignRight', IconAlignRight],
        ] as const
      ).map(([value, key, Icon]) => (
        <button
          key={value}
          type="button"
          className={`hx-float-btn${p.computed.textAlign === value ? ' on' : ''}`}
          data-tip={t(key)}
          aria-label={t(key)}
          aria-pressed={p.computed.textAlign === value}
          onMouseDown={stop}
          onClick={() => p.onStyle({ 'text-align': value })}
        >
          <Icon size={16} />
        </button>
      ))}
      <span className="hx-float-sep" />
    </>
  )
}

/** middle of the bar for <img>: replace, mirror, quarter-turn, block alignment */
function ImageTools(p: Props) {
  const { t } = useI18n()
  // rapid clicks must compose on what was just written, not on the snapshot still in flight
  const inline = (prop: string, snapshot: string) =>
    prop in p.pending ? (p.pending[prop] ?? '') : snapshot
  const current = parseImageTransform(inline('transform', p.computed.transform))
  const align = imageAlignOf(
    inline('margin-left', p.computed.marginLeft),
    inline('margin-right', p.computed.marginRight),
  )
  const buttons = [
    ['imageFlipH', IconFlipH, () => p.onStyle(transformStyles(flipImage(current, 'h')))],
    ['imageFlipV', IconFlipV, () => p.onStyle(transformStyles(flipImage(current, 'v')))],
    ['imageRotateLeft', IconRotateLeft, () => p.onStyle(transformStyles(rotateImage(current, -1)))],
    [
      'imageRotateRight',
      IconRotateRight,
      () => p.onStyle(transformStyles(rotateImage(current, 1))),
    ],
  ] as const
  return (
    <>
      <button
        type="button"
        className="hx-float-btn"
        data-tip={t('imageReplace')}
        aria-label={t('imageReplace')}
        onMouseDown={stop}
        onClick={p.onReplaceImage}
      >
        <IconReplacePicture size={18} />
      </button>
      <button
        type="button"
        className="hx-float-btn"
        data-tip={t('imageCrop')}
        aria-label={t('imageCrop')}
        onMouseDown={stop}
        onClick={p.onCropImage}
      >
        <IconCrop size={16} />
      </button>
      <button
        type="button"
        className="hx-float-btn"
        data-tip={t('imageRemoveBg')}
        aria-label={t('imageRemoveBg')}
        onMouseDown={stop}
        onClick={p.onCutoutImage}
      >
        <IconRemoveBg size={16} />
      </button>
      <span className="hx-float-sep" />
      {buttons.map(([key, Icon, onClick]) => (
        <button
          key={key}
          type="button"
          className="hx-float-btn"
          data-tip={t(key)}
          aria-label={t(key)}
          onMouseDown={stop}
          onClick={onClick}
        >
          <Icon size={16} />
        </button>
      ))}
      <span className="hx-float-sep" />
      {(
        [
          ['left', 'alignLeft', IconAlignLeft],
          ['center', 'alignCenter', IconAlignCenter],
          ['right', 'alignRight', IconAlignRight],
        ] as const
      ).map(([value, key, Icon]) => (
        <button
          key={value}
          type="button"
          className={`hx-float-btn${align === value ? ' on' : ''}`}
          data-tip={t(key)}
          aria-label={t(key)}
          aria-pressed={align === value}
          onMouseDown={stop}
          onClick={() => p.onStyle(imageAlignStyles(value))}
        >
          <Icon size={16} />
        </button>
      ))}
      <span className="hx-float-sep" />
    </>
  )
}

/** floating toolbar over the selected preview element (MaxGen / AI Design style) */
export function FloatToolbar(p: Props) {
  const { t } = useI18n()

  return (
    <div
      ref={p.barRef}
      className="hx-float"
      style={{ left: p.left, top: p.top }}
      role="toolbar"
      aria-label={t('elementToolbar')}
      onMouseDown={(e) => e.stopPropagation()}
    >
      {p.tag === 'img' ? <ImageTools {...p} /> : <TextTools {...p} />}
      {p.canEditText && (
        <button
          type="button"
          className="hx-float-btn text"
          onMouseDown={stop}
          onClick={p.onEditText}
        >
          {t('editText')}
        </button>
      )}
      <button
        type="button"
        className="hx-float-btn"
        data-tip={t('elementMoveUp')}
        aria-label={t('elementMoveUp')}
        onMouseDown={stop}
        onClick={() => p.onMove(-1)}
      >
        <IconMoveUp size={16} />
      </button>
      <button
        type="button"
        className="hx-float-btn"
        data-tip={t('elementMoveDown')}
        aria-label={t('elementMoveDown')}
        onMouseDown={stop}
        onClick={() => p.onMove(1)}
      >
        <IconMoveDown size={16} />
      </button>
      <button
        type="button"
        className="hx-float-btn"
        data-tip={t('elementDuplicate')}
        aria-label={t('elementDuplicate')}
        onMouseDown={stop}
        onClick={p.onDuplicate}
      >
        <IconCopy size={16} />
      </button>
      <button
        type="button"
        className="hx-float-btn danger"
        data-tip={t('elementDelete')}
        aria-label={t('elementDelete')}
        onMouseDown={stop}
        onClick={p.onDelete}
      >
        <IconTrash size={16} />
      </button>
      <span className="hx-float-sep" />
      <button
        type="button"
        className={`hx-float-btn${p.panelOpen ? ' on' : ''}`}
        data-tip={t('stylePanel')}
        aria-label={t('stylePanel')}
        aria-pressed={p.panelOpen}
        onMouseDown={stop}
        onClick={p.onTogglePanel}
      >
        <IconSliders size={16} />
      </button>
      {p.canAskAi && (
        <>
          <span className="hx-float-sep" />
          <button
            type="button"
            className="hx-float-btn text ask"
            data-tip={t('aiAskTitle')}
            onMouseDown={stop}
            onClick={p.onAskAi}
          >
            <IconSparkle size={15} />
            {t('aiAskBtn')}
          </button>
        </>
      )}
    </div>
  )
}
