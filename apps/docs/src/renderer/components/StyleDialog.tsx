import { useMemo, useState, type CSSProperties } from 'react'
import type { StyleInfo, StyleParaProps, StyleRunProps, StyleUpsert } from '@genoffice/docx-engine'
import { useI18n } from '../i18n/locale'
import { fontFamiliesFor } from '../font-list'
import { styleLabel, type StyleMap } from '../style-gallery'
import { useModalKeys } from './modal-keys'

type Align = NonNullable<StyleParaProps['align']>
const ALIGNS: Align[] = ['left', 'center', 'right', 'justify']
const LINE_SPACINGS = [1, 1.15, 1.5, 2, 2.5, 3]

/** The formatting the dialog edits; `undefined` = the style does not set it */
export interface StyleFormValues {
  name: string
  type: 'paragraph' | 'character'
  basedOn: string | null
  font: string
  sizePt: number | null
  bold: boolean
  italic: boolean
  /** hex without '#'; empty = automatic */
  color: string
  align: Align | null
  spaceBeforePt: number | null
  spaceAfterPt: number | null
  lineSpacing: number | null
}

type Field = keyof StyleFormValues

const toPt = (twips: number | undefined): number | null =>
  twips === undefined ? null : Math.round((twips / 20) * 10) / 10

/** initial form state from a style's resolved display (modify) */
export function formValuesOf(info: StyleInfo): StyleFormValues {
  const d = info.display ?? {}
  return {
    name: info.name,
    type: info.type === 'character' ? 'character' : 'paragraph',
    basedOn: info.basedOn ?? null,
    font: d.fontAscii ?? d.font ?? '',
    sizePt: d.sizeHalfPoints !== undefined ? d.sizeHalfPoints / 2 : null,
    bold: !!d.bold,
    italic: !!d.italic,
    color: d.color && d.color !== 'auto' ? d.color.toUpperCase() : '',
    align: d.align && d.align !== 'distribute' ? d.align : null,
    spaceBeforePt: toPt(d.spaceBeforeTwips),
    spaceAfterPt: toPt(d.spaceAfterTwips),
    lineSpacing: d.lineSpacing ?? null,
  }
}

/** Word writes the id as the name without spaces; ids must stay ASCII and unique */
export function styleIdFor(name: string, taken: (id: string) => boolean): string {
  const base = name.replace(/[^A-Za-z0-9_-]/g, '') || 'Style'
  let id = base
  for (let n = 1; taken(id); n++) id = `${base}${n}`
  return id
}

/** the w:style patch for the fields the user changed (or every field for a new style) */
export function upsertFromForm(
  styleId: string,
  values: StyleFormValues,
  touched: ReadonlySet<Field> | 'all',
  creating: boolean,
): StyleUpsert {
  const has = (f: Field) => touched === 'all' || touched.has(f)
  const up: StyleUpsert = { styleId }
  if (creating) up.type = values.type
  if (has('name')) up.name = values.name.trim()
  if (has('basedOn')) up.basedOn = values.basedOn
  const rPr: StyleRunProps = {}
  if (has('font')) rPr.font = values.font.trim() || null
  if (has('sizePt')) rPr.sizeHalfPoints = values.sizePt ? Math.round(values.sizePt * 2) : null
  if (has('bold')) rPr.bold = values.bold
  if (has('italic')) rPr.italic = values.italic
  if (has('color')) rPr.color = values.color || null
  if (Object.keys(rPr).length) up.rPr = rPr
  if (values.type === 'paragraph') {
    const pPr: StyleParaProps = {}
    if (has('align')) pPr.align = values.align
    if (has('spaceBeforePt'))
      pPr.spaceBeforeTwips =
        values.spaceBeforePt === null ? null : Math.round(values.spaceBeforePt * 20)
    if (has('spaceAfterPt'))
      pPr.spaceAfterTwips =
        values.spaceAfterPt === null ? null : Math.round(values.spaceAfterPt * 20)
    if (has('lineSpacing')) pPr.lineSpacing = values.lineSpacing
    if (Object.keys(pPr).length) up.pPr = pPr
  }
  return up
}

interface Props {
  /** null = new style */
  target: StyleInfo | null
  initial: StyleFormValues
  styles: StyleMap
  /** definitions queued for the save that the live map may not show yet */
  pending: Readonly<Record<string, StyleUpsert>>
  onSubmit: (up: StyleUpsert) => void
  onClose: () => void
}

/** Word's Modify Style / New Style dialog, reduced to the formatting the pane cards show */
export function StyleDialog({ target, initial, styles, pending, onSubmit, onClose }: Props) {
  const { t, lang } = useI18n()
  const modalKeys = useModalKeys(onClose)
  const [values, setValues] = useState(initial)
  const [touched, setTouched] = useState<Set<Field>>(() => new Set())
  const set = <F extends Field>(f: F, v: StyleFormValues[F]) => {
    setValues((prev) => ({ ...prev, [f]: v }))
    setTouched((prev) => new Set(prev).add(f))
  }
  const creating = target === null
  const fonts = useMemo(() => {
    const list = [...fontFamiliesFor(lang)]
    if (values.font && !list.includes(values.font)) list.unshift(values.font)
    return list
  }, [lang, values.font])
  const bases = useMemo(
    () =>
      [...styles.values()].filter(
        (s) => s.type === values.type && !s.linkedCharShell && s.styleId !== target?.styleId,
      ),
    [styles, values.type, target],
  )
  const name = values.name.trim()
  const lower = name.toLowerCase()
  const nameTaken =
    !!name &&
    ([...styles.values()].some(
      (s) => s.styleId !== target?.styleId && s.name.toLowerCase() === lower,
    ) ||
      Object.values(pending).some(
        (up) =>
          up.styleId !== target?.styleId &&
          (up.name ?? styles.get(up.styleId)?.name ?? up.styleId).toLowerCase() === lower,
      ))
  const canSubmit = !!name && !nameTaken
  const submit = () => {
    if (!canSubmit) return
    const id = creating
      ? styleIdFor(name, (candidate) => styles.has(candidate) || candidate in pending)
      : target.styleId
    onSubmit(upsertFromForm(id, values, creating ? 'all' : touched, creating))
  }
  const preview: CSSProperties = {
    fontFamily: values.font ? `"${values.font}"` : undefined,
    fontSize: values.sizePt ? `${Math.min(28, Math.max(9, values.sizePt))}px` : undefined,
    fontWeight: values.bold ? 700 : 400,
    fontStyle: values.italic ? 'italic' : 'normal',
    color: values.color ? `#${values.color}` : undefined,
    textAlign: values.align ?? undefined,
  }
  const num = (v: string): number | null => (v.trim() === '' ? null : Math.max(0, Number(v) || 0))

  return (
    <div
      className="modal-backdrop"
      ref={modalKeys.ref}
      onKeyDown={modalKeys.onKeyDown}
      onMouseDown={(e) => e.target === e.currentTarget && onClose()}
    >
      <div className="modal style-dialog">
        <h2>{t(creating ? 'ribbonStyleDialogNew' : 'ribbonStyleDialogModify')}</h2>
        <div className="font-dialog-row">
          <label>
            {t('ribbonStyleDialogName')}
            <input
              value={values.name}
              autoFocus
              onChange={(e) => set('name', e.target.value)}
              onKeyDown={(e) => e.key === 'Enter' && submit()}
            />
          </label>
          <label>
            {t('ribbonStyleDialogType')}
            <select
              value={values.type}
              disabled={!creating}
              onChange={(e) => {
                set('type', e.target.value as StyleFormValues['type'])
                set('basedOn', null)
              }}
            >
              <option value="paragraph">{t('ribbonStyleTypeParagraph')}</option>
              <option value="character">{t('ribbonStyleTypeCharacter')}</option>
            </select>
          </label>
        </div>
        {nameTaken && <div className="style-dialog-error">{t('ribbonStyleDialogNameTaken')}</div>}
        <div className="font-dialog-row">
          <label>
            {t('ribbonStyleDialogBasedOn')}
            <select
              value={values.basedOn ?? ''}
              onChange={(e) => set('basedOn', e.target.value || null)}
            >
              <option value="">{t('ribbonStyleDialogNoBase')}</option>
              {bases.map((s) => (
                <option key={s.styleId} value={s.styleId}>
                  {styleLabel(s, t)}
                </option>
              ))}
            </select>
          </label>
        </div>
        <div className="font-dialog-row">
          <label>
            {t('appFontTabFont')}
            <select value={values.font} onChange={(e) => set('font', e.target.value)}>
              <option value="">{'\u2014'}</option>
              {fonts.map((f) => (
                <option key={f} value={f}>
                  {f}
                </option>
              ))}
            </select>
          </label>
          <label className="style-dialog-num">
            {t('appFontSizeLabel')}
            <input
              type="number"
              min={1}
              max={1638}
              step={0.5}
              value={values.sizePt ?? ''}
              onChange={(e) => set('sizePt', num(e.target.value))}
            />
          </label>
          <label className="font-check">
            <input
              type="checkbox"
              checked={values.bold}
              onChange={(e) => set('bold', e.target.checked)}
            />
            {t('appFontBold')}
          </label>
          <label className="font-check">
            <input
              type="checkbox"
              checked={values.italic}
              onChange={(e) => set('italic', e.target.checked)}
            />
            {t('appFontItalic')}
          </label>
          <label className="style-dialog-color">
            {t('appFontColor')}
            <input
              type="color"
              className="font-color-input"
              value={`#${values.color || '000000'}`}
              onChange={(e) => set('color', e.target.value.slice(1).toUpperCase())}
            />
          </label>
        </div>
        {values.type === 'paragraph' && (
          <div className="font-dialog-row">
            <label>
              {t('appAlignment')}
              <select
                value={values.align ?? ''}
                onChange={(e) => set('align', (e.target.value || null) as Align | null)}
              >
                <option value="">{'\u2014'}</option>
                {ALIGNS.map((a) => (
                  <option key={a} value={a}>
                    {t(
                      a === 'left'
                        ? 'appAlignLeft'
                        : a === 'center'
                          ? 'appAlignCenter'
                          : a === 'right'
                            ? 'appAlignRight'
                            : 'appAlignJustify',
                    )}
                  </option>
                ))}
              </select>
            </label>
            <label className="style-dialog-num">
              {t('appSpaceBefore')}
              <input
                type="number"
                min={0}
                step={1}
                value={values.spaceBeforePt ?? ''}
                onChange={(e) => set('spaceBeforePt', num(e.target.value))}
              />
            </label>
            <label className="style-dialog-num">
              {t('appSpaceAfter')}
              <input
                type="number"
                min={0}
                step={1}
                value={values.spaceAfterPt ?? ''}
                onChange={(e) => set('spaceAfterPt', num(e.target.value))}
              />
            </label>
            <label className="style-dialog-num">
              {t('appLineSpacingLabel')}
              <select
                value={values.lineSpacing ?? ''}
                onChange={(e) =>
                  set('lineSpacing', e.target.value === '' ? null : Number(e.target.value))
                }
              >
                <option value="">{'\u2014'}</option>
                {LINE_SPACINGS.map((v) => (
                  <option key={v} value={v}>
                    {v}
                  </option>
                ))}
              </select>
            </label>
          </div>
        )}
        <div className="style-dialog-preview" style={preview}>
          AaBbCcYyZz
        </div>
        <div className="modal-actions">
          <button className="btn-ghost" onClick={onClose}>
            {t('appCancel')}
          </button>
          <button className="btn-primary" disabled={!canSubmit} onClick={submit}>
            {t('appOk')}
          </button>
        </div>
      </div>
    </div>
  )
}
