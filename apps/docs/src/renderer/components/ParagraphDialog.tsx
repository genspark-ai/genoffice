import { useState } from 'react'
import type { Editor } from '@tiptap/core'
import type { StyleInfo } from '@genoffice/docx-engine'
import { Dropdown } from '@genoffice/ui'
import { useI18n, type StringKey } from '../i18n/locale'
import { setParaAttrs, activeParaAttrs } from './ribbon-tabs'
import { setSelectionAlign } from '../editor/direction'
import { runUiOps } from '../ai/ops'
import { useModalKeys } from './modal-keys'
import { LengthInput } from './LengthInput'
import {
  firstLineFromSpecial,
  pickSpecial,
  specialFromFirstLine,
  type SpecialIndent,
} from './paragraph-special-indent'
import {
  lineSpacingAttrs,
  lineSpacingFromAttrs,
  pickLineSpacing,
  type LineSpacingChoice,
  type LineSpacingState,
} from './paragraph-line-spacing'

const PT_PER_TWIP = 1 / 20

type AlignValue = 'left' | 'center' | 'right' | 'justify'

/** visual alignment values; setSelectionAlign resolves them per paragraph direction */
const ALIGN_OPTIONS: Array<{ key: AlignValue; nameKey: StringKey }> = [
  { key: 'left', nameKey: 'appAlignLeft' },
  { key: 'center', nameKey: 'appAlignCenter' },
  { key: 'right', nameKey: 'appAlignRight' },
  { key: 'justify', nameKey: 'appAlignJustify' },
]

const LINE_SPACING_CHOICES: Array<{ key: LineSpacingChoice; nameKey: StringKey }> = [
  { key: 'single', nameKey: 'appLineSingle' },
  { key: 'oneHalf', nameKey: 'appLine15' },
  { key: 'double', nameKey: 'appLineDouble' },
  { key: 'atLeast', nameKey: 'appLineAtLeast' },
  { key: 'exact', nameKey: 'appLineExactly' },
  { key: 'multiple', nameKey: 'appLineMultiple' },
]

const SPECIAL_OPTIONS: Array<{ key: SpecialIndent; nameKey: StringKey }> = [
  { key: 'none', nameKey: 'appParaSpecialNone' },
  { key: 'firstLine', nameKey: 'appParaSpecialFirstLine' },
  { key: 'hanging', nameKey: 'appParaSpecialHanging' },
]

/** Letter width: the indent bound when no section is known */
const DEFAULT_PAGE_WIDTH = 12240

type Tab = 'indents' | 'breaks'

const FLAG_KEYS = ['widowControl', 'keepNext', 'keepLines', 'suppressLineNumbers'] as const
type FlagKey = (typeof FLAG_KEYS)[number]
const FLAG_LABELS: Record<FlagKey, StringKey> = {
  widowControl: 'appParaWidowControl',
  keepNext: 'appParaKeepWithNext',
  keepLines: 'appParaKeepLines',
  suppressLineNumbers: 'appParaSuppressLineNumbers',
}

/** Word shows the effective value: direct pPr, else the style chain, else the Word default */
function resolvedFlag(
  attrs: Record<string, unknown>,
  key: FlagKey | 'pageBreakBefore' | 'contextualSpacing',
  style: StyleInfo | undefined,
): boolean {
  const direct = attrs[key]
  if (typeof direct === 'boolean' && (key !== 'pageBreakBefore' || direct)) return direct
  const fromStyle = style?.display?.[key]
  if (typeof fromStyle === 'boolean') return fromStyle
  return key === 'widowControl'
}

/** outline level shown for the paragraph at the caret (0 = body text); a style-bound
 *  heading level and a list item are displayed but not editable here, like Word */
function outlineLevelOf(editor: Editor): { level: number; editable: boolean } {
  if (editor.isActive('docListItem')) return { level: 0, editable: false }
  if (!editor.isActive('docHeading')) return { level: 0, editable: true }
  const a = editor.getAttributes('docHeading')
  const level = Math.min(Math.max(Number(a.level) || 1, 1), 9)
  return { level, editable: a.outlineOnly === true }
}

export function ParagraphDialog({
  editor,
  onClose,
  styles,
  pageWidth = DEFAULT_PAGE_WIDTH,
}: {
  editor: Editor
  onClose: () => void
  /** ParsedDoc.styles: resolves the style-inherited state of the checkboxes */
  styles?: Map<string, StyleInfo>
  /** twips; Word lets indents go negative down to the page width */
  pageWidth?: number
}) {
  const { t } = useI18n()
  const modalKeys = useModalKeys(onClose)
  const attrs = activeParaAttrs(editor)
  const [tab, setTab] = useState<Tab>('indents')
  const paraStyle =
    (typeof attrs.styleId === 'string' && styles?.get(attrs.styleId)) ||
    (styles && [...styles.values()].find((s) => s.type === 'paragraph' && s.isDefault)) ||
    undefined
  const initFlags = () => {
    const out = {} as Record<FlagKey | 'pageBreakBefore' | 'contextualSpacing', boolean>
    for (const k of [...FLAG_KEYS, 'pageBreakBefore', 'contextualSpacing'] as const)
      out[k] = resolvedFlag(attrs, k, paraStyle)
    return out
  }
  const [flags, setFlags] = useState(initFlags)
  const [initialFlags] = useState(flags)
  const initOutline = outlineLevelOf(editor)
  const [outline, setOutline] = useState(initOutline.level)
  // unset align means "start": visually left in LTR, right in RTL (same as the ribbon)
  const [align, setAlign] = useState<AlignValue>(
    (attrs.align as AlignValue | null) ?? (attrs.bidi === true ? 'right' : 'left'),
  )
  const [line, setLine] = useState<LineSpacingState>(() => lineSpacingFromAttrs(attrs))
  const twipsToPt = (v: unknown) => Math.round((Number(v) || 0) * PT_PER_TWIP)
  const [indentLeft, setIndentLeft] = useState(Number(attrs.indentLeft) || 0)
  const [indentRight, setIndentRight] = useState(Number(attrs.indentRight) || 0)
  const [special, setSpecial] = useState(() => specialFromFirstLine(attrs.indentFirstLine))
  const [spaceBefore, setSpaceBefore] = useState(twipsToPt(attrs.spaceBefore))
  const [spaceAfter, setSpaceAfter] = useState(twipsToPt(attrs.spaceAfter))

  const apply = () => {
    if (!editor.isEditable) {
      onClose()
      return
    }
    const ptToTwips = (pt: number) => (pt > 0 ? Math.round(pt / PT_PER_TWIP) : null)
    const spacing = lineSpacingAttrs(line)
    // align goes through setSelectionAlign so each paragraph resolves the
    // visual value against its own direction (null = start side)
    setSelectionAlign(editor, align)
    // only touched checkboxes become direct pPr: an untouched one keeps showing the style through
    const changed: Record<string, unknown> = {}
    for (const k of [...FLAG_KEYS, 'contextualSpacing'] as const)
      if (flags[k] !== initialFlags[k]) changed[k] = flags[k]
    if (flags.pageBreakBefore !== initialFlags.pageBreakBefore)
      changed.pageBreakBefore = flags.pageBreakBefore
    setParaAttrs(editor, {
      ...spacing,
      indentLeft: indentLeft !== 0 ? indentLeft : null,
      indentRight: indentRight !== 0 ? indentRight : null,
      indentFirstLine: firstLineFromSpecial(special),
      spaceBefore: ptToTwips(spaceBefore),
      spaceAfter: ptToTwips(spaceAfter),
      ...changed,
    })
    if (initOutline.editable && outline !== initOutline.level)
      runUiOps(editor, [{ op: 'setOutlineLevel', target: { scope: 'selection' }, level: outline }])
    onClose()
  }

  const check = (key: FlagKey | 'pageBreakBefore' | 'contextualSpacing', label: string) => (
    <label className="font-check">
      <input
        type="checkbox"
        checked={flags[key]}
        onChange={(e) => setFlags((f) => ({ ...f, [key]: e.target.checked }))}
      />
      {label}
    </label>
  )

  const indentInput = (
    label: string,
    value: number,
    set: (v: number) => void,
    min = -pageWidth,
  ) => (
    <label>
      {label}
      <span className="para-num">
        <LengthInput
          value={value}
          min={min}
          max={pageWidth}
          ariaLabel={label}
          live
          onCommit={(twips) => set(twips ?? 0)}
        />
      </span>
    </label>
  )

  const numInput = (label: string, value: number, set: (v: number) => void) => (
    <label>
      {label}
      <span className="para-num">
        <input
          type="number"
          min={0}
          max={400}
          value={value}
          onChange={(e) => set(Math.max(0, Number(e.target.value) || 0))}
        />
        <span className="para-unit">pt</span>
      </span>
    </label>
  )

  return (
    <div
      className="modal-backdrop"
      ref={modalKeys.ref}
      onKeyDown={modalKeys.onKeyDown}
      onMouseDown={(e) => e.target === e.currentTarget && onClose()}
    >
      <div className="modal">
        <h2>{t('appParagraph')}</h2>
        <div className="dlg-tabs" role="tablist">
          {(['indents', 'breaks'] as const).map((k) => (
            <button
              key={k}
              type="button"
              role="tab"
              aria-selected={tab === k}
              className={tab === k ? 'active' : ''}
              onClick={() => setTab(k)}
            >
              {t(k === 'indents' ? 'appParaTabIndents' : 'appParaTabBreaks')}
            </button>
          ))}
        </div>
        {tab === 'breaks' ? (
          <div className="para-checks">
            {FLAG_KEYS.map((k) => check(k, t(FLAG_LABELS[k])))}
            {check('pageBreakBefore', t('appParaPageBreakBefore'))}
          </div>
        ) : (
          <>
            <div className="font-dialog-row">
              <label>
                {t('appAlignment')}
                <Dropdown
                  value={align}
                  ariaLabel={t('appAlignment')}
                  options={ALIGN_OPTIONS.map((o) => ({ value: o.key, label: t(o.nameKey) }))}
                  onPick={setAlign}
                />
              </label>
              <label>
                {t('appLineSpacingLabel')}
                <Dropdown
                  value={line.choice}
                  ariaLabel={t('appLineSpacingLabel')}
                  options={LINE_SPACING_CHOICES.map((c) => ({ value: c.key, label: t(c.nameKey) }))}
                  onPick={(choice) => setLine((prev) => pickLineSpacing(prev, choice))}
                />
              </label>
              <label>
                {t('appParaOutlineLevel')}
                <Dropdown
                  value={String(outline)}
                  ariaLabel={t('appParaOutlineLevel')}
                  disabled={!initOutline.editable}
                  options={[
                    { value: '0', label: t('appParaOutlineBody') },
                    ...Array.from({ length: 9 }, (_, i) => ({
                      value: String(i + 1),
                      label: t('appParaOutlineLevelN', { n: i + 1 }),
                    })),
                  ]}
                  onPick={(v) => setOutline(Number(v))}
                />
              </label>
            </div>
            <div className="font-dialog-row">
              {line.choice === 'exact' || line.choice === 'atLeast' ? (
                <label>
                  {t('appLineValue')}
                  <span className="para-num">
                    <input
                      type="number"
                      min={1}
                      max={1584}
                      step={0.5}
                      value={line.pt}
                      onChange={(e) =>
                        setLine({ ...line, pt: Math.max(0, Number(e.target.value) || 0) })
                      }
                    />
                    <span className="para-unit">pt</span>
                  </span>
                </label>
              ) : (
                <label>
                  {t('appLineValue')}
                  <span className="para-num">
                    <input
                      type="number"
                      min={0.06}
                      max={132}
                      step={0.01}
                      disabled={line.choice !== 'multiple'}
                      value={line.choice === 'multiple' ? line.multiple : ''}
                      onChange={(e) =>
                        setLine({
                          ...line,
                          multiple:
                            Math.round(Math.max(0.06, Number(e.target.value) || 1) * 100) / 100,
                        })
                      }
                    />
                    <span className="para-unit">×</span>
                  </span>
                </label>
              )}
            </div>
            <div className="font-dialog-row">
              {indentInput(t('appIndentLeft'), indentLeft, setIndentLeft)}
              {indentInput(t('appIndentRight'), indentRight, setIndentRight)}
            </div>
            <div className="font-dialog-row">
              <label>
                {t('appParaSpecial')}
                <Dropdown
                  value={special.special}
                  ariaLabel={t('appParaSpecial')}
                  options={SPECIAL_OPTIONS.map((o) => ({ value: o.key, label: t(o.nameKey) }))}
                  onPick={(kind) => setSpecial((prev) => pickSpecial(prev, kind))}
                />
              </label>
              <label>
                {t('appParaBy')}
                <span className="para-num">
                  <LengthInput
                    value={special.special === 'none' ? null : special.by}
                    min={0}
                    max={pageWidth}
                    disabled={special.special === 'none'}
                    ariaLabel={t('appParaBy')}
                    live
                    onCommit={(twips) => setSpecial((prev) => ({ ...prev, by: twips ?? 0 }))}
                  />
                </span>
              </label>
            </div>
            <div className="font-dialog-row">
              {numInput(t('appSpaceBefore'), spaceBefore, setSpaceBefore)}
              {numInput(t('appSpaceAfter'), spaceAfter, setSpaceAfter)}
            </div>
            <div className="para-checks">
              {check('contextualSpacing', t('appParaNoSpaceSameStyle'))}
            </div>
          </>
        )}
        <div className="modal-actions">
          <button className="btn-ghost" onClick={onClose}>
            {t('appCancel')}
          </button>
          <button className="btn-primary" onClick={apply}>
            {t('appOk')}
          </button>
        </div>
      </div>
    </div>
  )
}
