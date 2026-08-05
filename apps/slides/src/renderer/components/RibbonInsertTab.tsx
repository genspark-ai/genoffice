/** Insert tab of the slides ribbon. Extracted from Ribbon.tsx. */
import type { InsertKind } from '../../shared/ipc'
import {
  CHART_GALLERY,
  ICON_COLORS,
  ICON_GALLERY,
  SHAPE_GALLERY,
  SMARTART_GALLERY,
  WORDART_PRESETS,
} from '../insert-presets'
import { ChartKindThumb } from './ChartTypeDialog'
import { ShapePreview, SmartArtPreview } from './gallery-previews'
import {
  Icon3d,
  IconAudio,
  IconChart,
  IconComment,
  IconDateTime,
  IconEquation,
  IconFooter,
  IconIconLib,
  IconLink,
  IconNewSlide,
  IconPageNumber,
  IconPicture,
  IconScreenRec,
  IconShapes,
  IconSmartArt,
  IconTable,
  IconTextBox,
  IconVideo,
  IconWordArt,
  IconZoomJump,
} from './icons'
import { saveEditSelection } from '../TextEditOverlay'
import { BIG, Group, RbCaret, closeSiblingPanels, type RibbonTabCtx } from './ribbon-shared'

export function RibbonInsertTab({ rb }: { rb: RibbonTabCtx }) {
  const {
    closePanels,
    currentSlide,
    editing,
    hasDoc,
    hasSelection,
    layouts,
    onAddSlide,
    onAddSlideWithLayout,
    onInsert,
    onInsertChart,
    onInsertField,
    onInsertIcon,
    onInsertImage,
    onInsertMedia,
    onInsertModel3d,
    onInsertSmartArt,
    onInsertTable,
    onInsertWordArt,
    onInsertZoom,
    onNewComment,
    onOpenEquation,
    onOpenHeaderFooter,
    onOpenLink,
    onToggleScreenRecord,
    recording,
    slideCount,
    dropBig,
    iconColor,
    layoutOpen,
    setIconColor,
    setInsertDrop,
    setLayoutOpen,
    setTableCustom,
    setTableHover,
    setTableOpen,
    t,
    tableCustom,
    tableHover,
    tableOpen,
  } = rb
  return (
    <>
      <Group label={t('ribbonGroupSlides')}>
        <div className="rb-drop-wrap">
          <button
            className="rb-big"
            disabled={!hasDoc}
            onClick={onAddSlide}
            title={t('ribbonNewSlideTip')}
          >
            <span className="rb-big-icon">
              <IconNewSlide size={BIG} />
              <span
                className={`rb-caret-hit${layoutOpen ? ' active' : ''}`}
                title={t('ribbonChooseLayout')}
                onMouseDown={(e) => {
                  e.stopPropagation()
                  closeSiblingPanels(e, closePanels, 'layout')
                }}
                onClick={(e) => {
                  e.stopPropagation()
                  if (hasDoc) setLayoutOpen((v) => !v)
                }}
              >
                <RbCaret />
              </span>
            </span>
            <span>{t('ribbonNewSlide')}</span>
          </button>
          {layoutOpen && (
            <div className="rb-drop rb-layout-panel" onMouseDown={(e) => e.stopPropagation()}>
              <div className="rb-drop-title">{t('ribbonChooseLayout')}</div>
              <div className="rb-layout-list">
                {(layouts ?? []).map((lay) => (
                  <button
                    key={lay.path}
                    className="rb-layout-item"
                    onClick={() => {
                      setLayoutOpen(false)
                      onAddSlideWithLayout(lay.path)
                    }}
                    title={lay.name}
                  >
                    <div className="rb-layout-preview">
                      {lay.placeholders.map((ph, i) => {
                        const S = 120,
                          H = 68
                        const EMU = 9144000
                        const EMU_H = 5143500
                        const x = Math.round((ph.x / EMU) * S)
                        const y = Math.round((ph.y / EMU_H) * H)
                        const w = Math.max(8, Math.round((ph.cx / EMU) * S))
                        const h = Math.max(6, Math.round((ph.cy / EMU_H) * H))
                        return (
                          <div
                            key={i}
                            className="rb-layout-ph"
                            style={{ left: x, top: y, width: w, height: h }}
                          >
                            <span>
                              {ph.type === 'title' || ph.type === 'ctrTitle'
                                ? 'T'
                                : ph.type === 'body' || ph.type === 'obj'
                                  ? '≡'
                                  : ''}
                            </span>
                          </div>
                        )
                      })}
                    </div>
                    <div className="rb-layout-name">{lay.name}</div>
                  </button>
                ))}
                {(layouts ?? []).length === 0 && (
                  <div className="rb-layout-empty">{t('ribbonNoLayouts')}</div>
                )}
              </div>
            </div>
          )}
        </div>
      </Group>
      <div className="ribbon-sep" />
      <Group label={t('ribbonGroupTable')}>
        <div className="rb-drop-wrap">
          <button
            className={`rb-big ${tableOpen ? 'active' : ''}`}
            disabled={!hasDoc}
            title={t('ribbonInsertTableTip')}
            onMouseDown={(e) => {
              e.stopPropagation()
              closeSiblingPanels(e, closePanels, 'table')
            }}
            onClick={() => setTableOpen((v) => !v)}
          >
            <span className="rb-big-icon">
              <IconTable size={BIG} />
            </span>
            <span>{t('ribbonGroupTable')}</span>
          </button>
          {tableOpen && (
            <div className="rb-drop rb-table-picker" onMouseDown={(e) => e.stopPropagation()}>
              <div className="rb-table-picker-label">
                {tableHover.r > 0
                  ? t('ribbonTableSize', { r: tableHover.r, c: tableHover.c })
                  : t('ribbonTablePickerHint')}
              </div>
              <div className="rb-table-grid" onMouseLeave={() => setTableHover({ r: 0, c: 0 })}>
                {Array.from({ length: 8 }, (_, r) =>
                  Array.from({ length: 10 }, (_, c) => (
                    <button
                      key={`${r}-${c}`}
                      className={`rb-table-cell ${r < tableHover.r && c < tableHover.c ? 'on' : ''}`}
                      onMouseEnter={() => setTableHover({ r: r + 1, c: c + 1 })}
                      onClick={() => {
                        setTableOpen(false)
                        onInsertTable(r + 1, c + 1)
                      }}
                    />
                  )),
                )}
              </div>
              <div className="rb-table-custom">
                <input
                  type="number"
                  min={1}
                  max={50}
                  value={tableCustom.r}
                  onChange={(e) =>
                    setTableCustom((v) => ({
                      ...v,
                      r: Math.max(1, Math.min(50, Number(e.target.value) || 1)),
                    }))
                  }
                />
                <span>×</span>
                <input
                  type="number"
                  min={1}
                  max={50}
                  value={tableCustom.c}
                  onChange={(e) =>
                    setTableCustom((v) => ({
                      ...v,
                      c: Math.max(1, Math.min(50, Number(e.target.value) || 1)),
                    }))
                  }
                />
                <button
                  className="rb-table-custom-ok"
                  onClick={() => {
                    setTableOpen(false)
                    onInsertTable(tableCustom.r, tableCustom.c)
                  }}
                >
                  {t('paneOk')}
                </button>
              </div>
            </div>
          )}
        </div>
      </Group>
      <div className="ribbon-sep" />
      <Group label={t('ribbonGroupImages')}>
        <button
          className="rb-big"
          disabled={!hasDoc}
          onClick={onInsertImage}
          title={t('ribbonPictureTip')}
        >
          <span className="rb-big-icon">
            <IconPicture size={BIG} />
          </span>
          <span>{t('ribbonPicture')}</span>
        </button>
        {dropBig(
          'icons',
          <IconIconLib size={BIG} />,
          t('ribbonIcons'),
          t('ribbonIconsTip'),
          <div className="rb-icon-gallery">
            <div className="rb-icon-colors">
              {ICON_COLORS.map((c) => (
                <button
                  key={c}
                  className={`rb-swatch ${iconColor === c ? 'on' : ''}`}
                  style={{ background: c }}
                  title={c}
                  onClick={() => setIconColor(c)}
                />
              ))}
            </div>
            <div className="rb-icon-grid">
              {ICON_GALLERY.map((def) => (
                <button
                  key={def.name}
                  className="rb-icon-cell"
                  title={def.name}
                  onClick={() => {
                    setInsertDrop(null)
                    onInsertIcon(def, iconColor)
                  }}
                >
                  <svg
                    viewBox="0 0 24 24"
                    width={22}
                    height={22}
                    fill="none"
                    stroke={iconColor}
                    strokeWidth={1.5}
                    strokeLinecap="round"
                    strokeLinejoin="round"
                    dangerouslySetInnerHTML={{ __html: def.body }}
                  />
                </button>
              ))}
            </div>
          </div>,
        )}
        <button
          className="rb-big"
          disabled={!hasDoc}
          onClick={onInsertModel3d}
          title={t('ribbon3dModelTip')}
        >
          <span className="rb-big-icon">
            <Icon3d size={BIG} />
          </span>
          <span>{t('ribbon3dModel')}</span>
        </button>
      </Group>
      <div className="ribbon-sep" />
      <Group label={t('ribbonGroupIllustrations')}>
        {dropBig(
          'shapes',
          <IconShapes size={BIG} />,
          t('ribbonShapes'),
          t('ribbonShapesTip'),
          <div className="rb-shape-gallery">
            {SHAPE_GALLERY.map((group) => (
              <div key={group.group}>
                <div className="rb-drop-title">{group.group}</div>
                <div className="rb-shape-grid">
                  {group.shapes.map((s) => (
                    <button
                      key={s.prst}
                      className="rb-shape-cell"
                      title={s.label}
                      onClick={() => {
                        setInsertDrop(null)
                        onInsert(s.prst as InsertKind)
                      }}
                    >
                      <ShapePreview prst={s.prst} size={18} />
                    </button>
                  ))}
                </div>
              </div>
            ))}
          </div>,
        )}
        {dropBig(
          'smartart',
          <IconSmartArt size={BIG} />,
          'SmartArt',
          t('ribbonSmartArtTip'),
          <div className="rb-menu rb-menu-wide">
            {SMARTART_GALLERY.map((def) => (
              <button
                key={def.layout}
                onClick={() => {
                  setInsertDrop(null)
                  onInsertSmartArt(def)
                }}
              >
                <span className="rb-menu-glyph rb-menu-thumb">
                  <SmartArtPreview layout={def.layout} width={40} />
                </span>
                {def.label}
              </button>
            ))}
          </div>,
        )}
        {dropBig(
          'chart',
          <IconChart size={BIG} />,
          t('ribbonChart'),
          t('ribbonChartTip'),
          <div className="rb-menu rb-menu-wide">
            {CHART_GALLERY.map((def) => (
              <button
                key={def.kind}
                onClick={() => {
                  setInsertDrop(null)
                  onInsertChart(def.kind)
                }}
              >
                <span className="rb-menu-glyph rb-menu-thumb">
                  <ChartKindThumb kind={def.kind} width={40} />
                </span>
                {def.label}
              </button>
            ))}
          </div>,
        )}
      </Group>
      <div className="ribbon-sep" />
      <Group label={t('ribbonGroupLinks')}>
        <button
          className="rb-big"
          data-keep-edit=""
          disabled={!hasDoc || (!hasSelection && !editing)}
          onMouseDown={(e) => {
            // While editing, keep the editor selection alive so the link applies to it (run-level)
            if (editing) {
              e.preventDefault()
              saveEditSelection()
            }
          }}
          onClick={onOpenLink}
          title={hasSelection || editing ? t('ribbonLinkTip') : t('ribbonLinkTipDisabled')}
        >
          <span className="rb-big-icon">
            <IconLink size={BIG} />
          </span>
          <span>{t('ribbonLink')}</span>
        </button>
        {dropBig(
          'zoom',
          <IconZoomJump size={BIG} />,
          t('ribbonZoomJump'),
          t('ribbonZoomJumpTip'),
          <div className="rb-menu rb-menu-scroll">
            {Array.from({ length: slideCount }, (_, i) => (
              <button
                key={i}
                disabled={i === currentSlide}
                onClick={() => {
                  setInsertDrop(null)
                  onInsertZoom(i)
                }}
              >
                {t('ribbonZoomJumpItem', { n: i + 1 })}
                {i === currentSlide ? t('ribbonCurrentSlideSuffix') : ''}
              </button>
            ))}
          </div>,
        )}
      </Group>
      <div className="ribbon-sep" />
      <Group label={t('ribbonGroupComments')}>
        <button
          className="rb-big"
          disabled={!hasDoc}
          onClick={onNewComment}
          title={t('ribbonNewCommentTip')}
        >
          <span className="rb-big-icon">
            <IconComment size={BIG} />
          </span>
          <span>{t('ribbonComment')}</span>
        </button>
      </Group>
      <div className="ribbon-sep" />
      <Group label={t('ribbonShapeGroupText')}>
        <button
          className="rb-big"
          disabled={!hasDoc}
          onClick={() => onInsert('textbox')}
          title={t('ribbonInsertTextBoxTip')}
        >
          <span className="rb-big-icon">
            <IconTextBox size={BIG} />
          </span>
          <span>{t('ribbonTextBox')}</span>
        </button>
        {dropBig(
          'wordart',
          <IconWordArt size={BIG} />,
          t('ribbonWordArt'),
          t('ribbonWordArtTip'),
          <div className="rb-wordart-grid">
            {WORDART_PRESETS.map((p) => (
              <button
                key={p.id}
                className="rb-wordart-cell"
                title={t('ribbonWordArtCellTip')}
                style={{
                  color: p.fill,
                  WebkitTextStroke: p.outline ? `1px ${p.outline.color}` : undefined,
                  fontWeight: p.bold ? 800 : 400,
                  fontStyle: p.italic ? 'italic' : undefined,
                }}
                onClick={() => {
                  setInsertDrop(null)
                  onInsertWordArt(p)
                }}
              >
                A
              </button>
            ))}
          </div>,
        )}
        <button
          className="rb-big"
          disabled={!hasDoc}
          onClick={onOpenHeaderFooter}
          title={t('ribbonHeaderFooterTip')}
        >
          <span className="rb-big-icon">
            <IconFooter size={BIG} />
          </span>
          <span>{t('ribbonHeaderFooter')}</span>
        </button>
        <button
          className="rb-big"
          disabled={!hasDoc}
          onClick={() => onInsertField('datetime')}
          title={t('ribbonDateTimeTip')}
        >
          <span className="rb-big-icon">
            <IconDateTime size={BIG} />
          </span>
          <span>{t('ribbonDateTime')}</span>
        </button>
        <button
          className="rb-big"
          disabled={!hasDoc}
          onClick={() => onInsertField('slidenum')}
          title={t('ribbonSlideNumberTip')}
        >
          <span className="rb-big-icon">
            <IconPageNumber size={BIG} />
          </span>
          <span>{t('ribbonSlideNumber')}</span>
        </button>
      </Group>
      <div className="ribbon-sep" />
      <Group label={t('ribbonGroupSymbols')}>
        <button
          className="rb-big"
          disabled={!hasDoc}
          onClick={onOpenEquation}
          title={t('ribbonEquationTip')}
        >
          <span className="rb-big-icon">
            <IconEquation size={BIG} />
          </span>
          <span>{t('ribbonEquation')}</span>
        </button>
      </Group>
      <div className="ribbon-sep" />
      <Group label={t('ribbonGroupMedia')}>
        <button
          className="rb-big"
          disabled={!hasDoc}
          onClick={() => onInsertMedia('video')}
          title={t('ribbonVideoTip')}
        >
          <span className="rb-big-icon">
            <IconVideo size={BIG} />
          </span>
          <span>{t('ribbonVideo')}</span>
        </button>
        <button
          className="rb-big"
          disabled={!hasDoc}
          onClick={() => onInsertMedia('audio')}
          title={t('ribbonAudioTip')}
        >
          <span className="rb-big-icon">
            <IconAudio size={BIG} />
          </span>
          <span>{t('ribbonAudio')}</span>
        </button>
        <button
          className={`rb-big ${recording ? 'active rb-recording' : ''}`}
          disabled={!hasDoc}
          onClick={onToggleScreenRecord}
          title={recording ? t('ribbonStopRecTip') : t('ribbonScreenRecTip')}
        >
          <span className="rb-big-icon">
            <IconScreenRec size={BIG} />
          </span>
          <span>{recording ? t('ribbonStopRec') : t('ribbonScreenRec')}</span>
        </button>
      </Group>
    </>
  )
}
