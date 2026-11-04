import React, { useRef } from 'react'
import { HIGHLIGHTER_ALPHA, isPenKind, type InkPenKind, type InkTool } from '../ink'
import { colorsFor, newInkPreset, widthsFor, type InkPreset } from '../ink-presets'
import { useI18n } from '../i18n/locale'
import { BIG, Group, RbCaret, closeSiblingPanels, type RibbonPanelKey } from './ribbon-shared'
import {
  IconCursor,
  IconEraser,
  IconHighlighterPen,
  IconPalette,
  IconPen,
  IconPencilInk,
  IconPlus,
  IconTrash,
} from './icons'

/** which Draw-tab popup is open: one pen slot's properties, or the Add menu */
export type PenPanel = `slot:${string}` | 'add'

interface DrawTabProps {
  hasDoc: boolean
  tool: InkTool
  onTool: (tool: InkTool) => void
  presets: InkPreset[]
  activeId: string
  onPresets: (presets: InkPreset[]) => void
  onActivate: (id: string) => void
  inkCount: number
  onClearAll: () => void
  panel: PenPanel | null
  setPanel: (panel: PenPanel | null) => void
  closePanels: (keep: RibbonPanelKey[]) => void
}

function PenGlyph({ kind, size }: { kind: InkPenKind; size: number }) {
  if (kind === 'highlighter') return <IconHighlighterPen size={size} />
  if (kind === 'pencil') return <IconPencilInk size={size} />
  return <IconPen size={size} />
}

/**
 * Draw tab, same model as apps/docs: a row of remembered pens, each with its own
 * kind / color / thickness. Clicking a slot arms that pen, clicking the armed
 * slot opens its properties; Add appends a pen, highlighter or pencil.
 */
export function RibbonDrawTab({
  hasDoc,
  tool,
  onTool,
  presets,
  activeId,
  onPresets,
  onActivate,
  inkCount,
  onClearAll,
  panel,
  setPanel,
  closePanels,
}: DrawTabProps) {
  const { t } = useI18n()
  const colorInputs = useRef(new Map<string, HTMLInputElement>())
  const drawing = isPenKind(tool)
  const kindLabel = (kind: InkPenKind) =>
    kind === 'highlighter'
      ? t('ribbonHighlighter')
      : kind === 'pencil'
        ? t('ribbonPencil')
        : t('ribbonPen')
  const update = (id: string, patch: Partial<InkPreset>) =>
    onPresets(presets.map((p) => (p.id === id ? { ...p, ...patch } : p)))
  const remove = (id: string) => {
    const rest = presets.filter((p) => p.id !== id)
    onPresets(rest)
    if (id === activeId) {
      onActivate(rest[0]!.id)
      if (drawing) onTool(rest[0]!.kind)
    }
    setPanel(null)
  }
  const add = (kind: InkPenKind) => {
    const preset = newInkPreset(kind)
    onPresets([...presets, preset])
    onActivate(preset.id)
    onTool(kind)
    setPanel(null)
  }
  const pickSlot = (preset: InkPreset) => {
    if (drawing && preset.id === activeId) {
      setPanel(panel === `slot:${preset.id}` ? null : `slot:${preset.id}`)
      return
    }
    setPanel(null)
    onActivate(preset.id)
    onTool(preset.kind)
  }
  const onTriggerDown = (e: React.MouseEvent<HTMLElement>) => {
    e.stopPropagation()
    closeSiblingPanels(e, closePanels, 'pen')
  }

  return (
    <>
      <Group label={t('ribbonGroupDrawTools')}>
        <button
          className={`rb-big ${tool === 'select' ? 'active' : ''}`}
          disabled={!hasDoc}
          data-tip={t('ribbonSelectTip')}
          onClick={() => {
            setPanel(null)
            onTool('select')
          }}
        >
          <span className="rb-big-icon">
            <IconCursor size={BIG} />
          </span>
          <span>{t('ribbonGroupSelect')}</span>
        </button>
        <button
          className={`rb-big ${tool === 'eraser' ? 'active' : ''}`}
          disabled={!hasDoc}
          data-tip={t('ribbonEraserTip')}
          onClick={() => {
            setPanel(null)
            onTool(tool === 'eraser' ? 'select' : 'eraser')
          }}
        >
          <span className="rb-big-icon">
            <IconEraser size={BIG} />
          </span>
          <span>{t('ribbonEraser')}</span>
        </button>
      </Group>
      <div className="ribbon-sep" />
      <div className="ribbon-group">
        <div className="ribbon-group-items ink-pens">
          {presets.map((preset) => {
            const armed = drawing && preset.id === activeId
            const open = panel === `slot:${preset.id}`
            const widths = widthsFor(preset.kind)
            const colors = colorsFor(preset.kind)
            const alpha = preset.kind === 'highlighter' ? HIGHLIGHTER_ALPHA : 1
            return (
              <div key={preset.id} className="rb-drop-wrap">
                <button
                  className={`ink-pen-slot ${armed ? 'active' : ''}`}
                  disabled={!hasDoc}
                  data-tip={`${kindLabel(preset.kind)} · #${preset.color}`}
                  aria-label={kindLabel(preset.kind)}
                  onMouseDown={onTriggerDown}
                  onClick={() => pickSlot(preset)}
                >
                  <span className="ink-pen-glyph" style={{ color: `#${preset.color}` }}>
                    <PenGlyph kind={preset.kind} size={28} />
                  </span>
                  <span
                    className="ink-pen-bar"
                    style={{
                      background: `#${preset.color}`,
                      height: Math.max(
                        2,
                        Math.min(8, preset.width / (preset.kind === 'highlighter' ? 4 : 1)),
                      ),
                      opacity: alpha,
                    }}
                  />
                  {armed && <RbCaret />}
                </button>
                {/* lives outside the popup: the native picker blurs the window, which
                    dismisses the popup, and an unmounted input would tear the picker down */}
                <input
                  ref={(el) => {
                    if (el) colorInputs.current.set(preset.id, el)
                    else colorInputs.current.delete(preset.id)
                  }}
                  type="color"
                  className="ink-pen-color-input"
                  tabIndex={-1}
                  aria-hidden="true"
                  value={`#${preset.color}`}
                  onChange={(e) =>
                    update(preset.id, { color: e.target.value.slice(1).toUpperCase() })
                  }
                />
                {open && (
                  <div className="rb-drop ink-pen-pop" onMouseDown={(e) => e.stopPropagation()}>
                    <div className="ink-pen-preview">
                      <span
                        style={{
                          background: `#${preset.color}`,
                          height: preset.width,
                          opacity: alpha,
                        }}
                      />
                    </div>
                    <div className="ink-pen-section">{t('ribbonInkThickness')}</div>
                    <div className="ink-widths">
                      {widths.map((w, i) => (
                        <button
                          key={w}
                          className={`ink-width ${preset.width === w ? 'active' : ''}`}
                          data-tip={t('ribbonInkWidthTip', { w })}
                          aria-label={t('ribbonInkWidthTip', { w })}
                          onClick={() => update(preset.id, { width: w })}
                        >
                          <span
                            className="ink-width-dot"
                            style={{
                              width: 4 + i * 3,
                              height: 4 + i * 3,
                              background: `#${preset.color}`,
                            }}
                          />
                        </button>
                      ))}
                    </div>
                    <div className="ink-pen-section">{t('ribbonInkColors')}</div>
                    <div className="ink-pen-grid">
                      {colors.map((hex) => (
                        <button
                          key={hex}
                          className={`ink-swatch ${preset.color === hex ? 'active' : ''}`}
                          style={{ background: `#${hex}` }}
                          data-tip={`#${hex}`}
                          aria-label={`#${hex}`}
                          onClick={() => update(preset.id, { color: hex })}
                        />
                      ))}
                    </div>
                    <button
                      className="ink-pen-item ink-pen-more"
                      onClick={() => colorInputs.current.get(preset.id)?.click()}
                    >
                      <IconPalette size={16} />
                      {t('ribbonMoreColors')}
                    </button>
                    <button
                      className="ink-pen-item ink-pen-delete"
                      disabled={presets.length <= 1}
                      onClick={() => remove(preset.id)}
                    >
                      <IconTrash size={16} />
                      {t('ribbonInkDelete')}
                    </button>
                  </div>
                )}
              </div>
            )
          })}
          <div className="rb-drop-wrap">
            <button
              className={`rb-big ink-pen-add ${panel === 'add' ? 'active' : ''}`}
              disabled={!hasDoc}
              data-tip={t('ribbonInkAddTip')}
              onMouseDown={onTriggerDown}
              onClick={() => setPanel(panel === 'add' ? null : 'add')}
            >
              <span className="rb-big-icon">
                <IconPlus size={BIG} />
                <RbCaret />
              </span>
              <span>{t('ribbonInkAdd')}</span>
            </button>
            {panel === 'add' && (
              <div className="rb-drop ink-pen-pop" onMouseDown={(e) => e.stopPropagation()}>
                {(['pen', 'highlighter', 'pencil'] as const).map((kind) => (
                  <button key={kind} className="ink-pen-item" onClick={() => add(kind)}>
                    <PenGlyph kind={kind} size={18} />
                    {kindLabel(kind)}
                  </button>
                ))}
              </div>
            )}
          </div>
        </div>
        <div className="ribbon-group-label">{t('ribbonGroupPens')}</div>
      </div>
      <div className="ribbon-sep" />
      <Group label={t('ribbonGroupClear')}>
        <button
          className="rb-big"
          disabled={!hasDoc || inkCount === 0}
          data-tip={t('ribbonEraseAllTip')}
          onClick={onClearAll}
        >
          <span className="rb-big-icon">
            <IconEraser size={BIG} />
          </span>
          <span>{t('ribbonEraseAll')}</span>
        </button>
      </Group>
    </>
  )
}
