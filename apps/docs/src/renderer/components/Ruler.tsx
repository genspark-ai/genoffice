import { useRef, useState } from 'react'
import type { MouseEvent as ReactMouseEvent } from 'react'
import type { Editor } from '@tiptap/core'
import type { SectionSettings, TabStop } from '@genoffice/docx-engine'
import { t, type StringKey } from '../i18n/locale'

const twipsToPx = (twips: number) => (twips / 1440) * 96

/** Horizontal ruler above the page: inch numbers, gray margin zones, tab stops. */
export function Ruler({
  section,
  editor,
  onTabStopsChange,
}: {
  section: SectionSettings
  editor: Editor | null
  onTabStopsChange: (stops: TabStop[] | null) => void
}) {
  const width = twipsToPx(section.pageWidth)
  const marginLeft = twipsToPx(section.marginLeft)
  const marginRight = twipsToPx(section.marginRight)
  const inches = Math.floor(section.pageWidth / 1440)
  const ticks: number[] = []
  for (let i = 1; i <= inches; i++) ticks.push(i)

  // Default Word tab interval: 0.5in = 720 twips
  const DEFAULT_TAB_TWIPS = 720

  // Tab stop type cycling (Word: click ruler button to cycle L/C/R/Decimal/Bar)
  const [nextTabType, setNextTabType] = useState<TabStop['val']>('left')
  const TAB_TYPE_CYCLE: TabStop['val'][] = ['left', 'center', 'right', 'decimal']
  const TAB_TYPE_LABELS: Record<string, string> = {
    left: 'L',
    center: '⊥',
    right: '⌐',
    decimal: '.',
    bar: '|',
  }
  const TAB_TYPE_NAME_KEYS: Record<TabStop['val'], StringKey> = {
    left: 'appTabLeft',
    center: 'appTabCenter',
    right: 'appTabRight',
    decimal: 'appTabDecimal',
    bar: 'appTabBar',
    clear: 'appTabBar',
  }

  // Get current tab stops from focused paragraph
  const currentTabStops = (): TabStop[] => {
    if (!editor) return []
    const attrs = editor.isActive('docHeading')
      ? editor.getAttributes('docHeading')
      : editor.isActive('docListItem')
        ? editor.getAttributes('docListItem')
        : editor.getAttributes('docParagraph')
    const raw = attrs?.tabStops as string | null
    if (!raw) return []
    try {
      const parsed = JSON.parse(raw)
      return Array.isArray(parsed) ? parsed : []
    } catch {
      return []
    }
  }

  const stops = currentTabStops()

  // Drag state
  const dragRef = useRef<{ stopIndex: number; startX: number; origPos: number } | null>(null)

  // Click on ruler: add tab stop at position, skip margin zones
  const handleRulerClick = (e: ReactMouseEvent<HTMLDivElement>) => {
    const rect = (e.currentTarget as HTMLElement).getBoundingClientRect()
    const x = e.clientX - rect.left
    if (x < marginLeft || x > width - marginRight) return
    const posTwips = Math.round((x / width) * section.pageWidth)
    // snap to nearest 60 twips (~0.04in)
    const snapped = Math.round(posTwips / 60) * 60
    const existing = stops.filter((s) => Math.abs(s.pos - snapped) > 60)
    const newStop: TabStop = { pos: snapped, val: nextTabType }
    const newStops = [...existing, newStop].sort((a, b) => a.pos - b.pos)
    onTabStopsChange(newStops.length > 0 ? newStops : null)
  }

  // Drag tab stop to new position or drop outside to delete
  const handleTabMouseDown = (e: ReactMouseEvent<HTMLSpanElement>, stopIndex: number) => {
    e.stopPropagation()
    e.preventDefault()
    const rect = (e.currentTarget.closest('.ruler') as HTMLElement).getBoundingClientRect()
    dragRef.current = { stopIndex, startX: e.clientX, origPos: stops[stopIndex].pos }

    const onMouseMove = (ev: MouseEvent) => {
      if (!dragRef.current) return
      const x = ev.clientX - rect.left
      const posTwips = Math.round((x / width) * section.pageWidth)
      const snapped = Math.round(posTwips / 60) * 60
      // visual only update via CSS custom property (no state update for perf)
      const marker = document.querySelector(
        `[data-ruler-stop="${stopIndex}"]`,
      ) as HTMLElement | null
      if (marker) marker.style.left = `${twipsToPx(snapped)}px`
    }

    const onMouseUp = (ev: MouseEvent) => {
      document.removeEventListener('mousemove', onMouseMove)
      document.removeEventListener('mouseup', onMouseUp)
      if (!dragRef.current) return
      const x = ev.clientX - rect.left
      // drop outside the content area: delete the stop
      if (x < marginLeft || x > width - marginRight) {
        const newStops = stops.filter((_, i) => i !== dragRef.current!.stopIndex)
        onTabStopsChange(newStops.length > 0 ? newStops : null)
      } else {
        const posTwips = Math.round((x / width) * section.pageWidth)
        const snapped = Math.round(posTwips / 60) * 60
        const newStops = stops
          .map((s, i) => (i === dragRef.current!.stopIndex ? { ...s, pos: snapped } : s))
          .sort((a, b) => a.pos - b.pos)
        onTabStopsChange(newStops)
      }
      dragRef.current = null
    }

    document.addEventListener('mousemove', onMouseMove)
    document.addEventListener('mouseup', onMouseUp)
  }

  // Default tab stop markers (light gray) when no custom stops defined
  const defaultStops: number[] = []
  if (stops.length === 0) {
    const contentWidth = section.pageWidth - section.marginLeft - section.marginRight
    for (let pos = DEFAULT_TAB_TWIPS; pos < contentWidth; pos += DEFAULT_TAB_TWIPS) {
      defaultStops.push(section.marginLeft + pos)
    }
  }

  return (
    <div className="ruler" style={{ width }} onClick={handleRulerClick}>
      {/* Tab type selector button at far left */}
      <button
        className="ruler-tab-type"
        data-tip={t('appTabTypeTip', { type: t(TAB_TYPE_NAME_KEYS[nextTabType]) })}
        onClick={(e) => {
          e.stopPropagation()
          const idx = TAB_TYPE_CYCLE.indexOf(nextTabType)
          setNextTabType(TAB_TYPE_CYCLE[(idx + 1) % TAB_TYPE_CYCLE.length])
        }}
      >
        {TAB_TYPE_LABELS[nextTabType]}
      </button>

      <div className="ruler-zone" style={{ left: 0, width: marginLeft }} />
      <div className="ruler-zone" style={{ left: width - marginRight, width: marginRight }} />

      {ticks.map((i) => (
        <span key={i} className="ruler-num" style={{ left: twipsToPx(i * 1440) }}>
          {i}
        </span>
      ))}

      {/* Default tab stop guides (light, no interaction) */}
      {defaultStops.map((posTwips) => (
        <span
          key={`def-${posTwips}`}
          className="ruler-tab-default"
          style={{ left: twipsToPx(posTwips) }}
        />
      ))}

      {/* Custom tab stops (interactive) */}
      {stops.map((stop, i) => (
        <span
          key={`${stop.pos}-${i}`}
          data-ruler-stop={i}
          className={`ruler-tab ruler-tab-${stop.val}`}
          style={{ left: twipsToPx(stop.pos) }}
          data-tip={
            t('appTabStopTitle', {
              type: t(TAB_TYPE_NAME_KEYS[stop.val]),
              pos: Math.round((stop.pos / 144) * 10) / 10,
            }) + (stop.leader ? t('appTabLeader', { leader: stop.leader }) : '')
          }
          onMouseDown={(e) => handleTabMouseDown(e, i)}
        >
          {TAB_TYPE_LABELS[stop.val]}
        </span>
      ))}
    </div>
  )
}
