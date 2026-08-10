import { useI18n } from './i18n/locale'

/// Slicer: a visual filter control bound to one pivot dimension field. When
/// members are clicked, App writes the unselected members as the pivot's hidden
/// entries and recomputes (pivot-engine's applyPivotSlicer), matching how Excel
/// slicers drive pivotField hidden items.
///
/// TODO(OOXML persistence): the xl/slicers/slicer1.xml + xl/slicerCaches/ parts
/// (incl. workbook rels, contentTypes, and x14/x15 extensions) are heavy, so
/// slicers currently exist only within the session (App state); the hidden
/// entries and output-area data they drive are persisted through the existing
/// pivot refresh write-back path — Excel shows the same filtered result on open,
/// just without the slicer control itself.

export interface SlicerMember {
  /// fieldItems index (the member's entry position in the pivot cache, stable
  /// across layout growth).
  readonly member: number
  readonly label: string
}

export interface SlicerUiState {
  readonly id: string
  readonly sheetId: string
  /// Bound pivot part path (fileMeta.pivotTables[].path).
  readonly pivotPath: string
  /// Cache-field index and display name of the sliced field.
  readonly field: number
  readonly fieldName: string
  readonly members: readonly SlicerMember[]
  /// Selected members (fieldItems indices); all selected = unfiltered.
  readonly selected: readonly number[]
}

/// Field picker when creating a slicer: lists the current pivot's dimension
/// fields (row/column/report filter).
export function SlicerFieldPicker({
  fields,
  onPick,
  onClose,
}: {
  readonly fields: readonly { readonly field: number; readonly name: string }[]
  /// Returns an error message; null = created successfully (the caller closes
  /// the dialog).
  readonly onPick: (field: number) => string | null
  readonly onClose: () => void
}): React.JSX.Element {
  const { t } = useI18n()
  return (
    <div className="dialog-backdrop" onClick={onClose}>
      <div
        className="format-cells-dialog slicer-picker"
        role="dialog"
        aria-label={t('dlgSlicerInsertTitle')}
        onClick={(event) => event.stopPropagation()}
      >
        <header>{t('dlgSlicerInsertTitle')}</header>
        <p className="dialog-note">{t('dlgSlicerPickNote')}</p>
        <div className="slicer-picker-fields">
          {fields.map((entry) => (
            <button
              key={entry.field}
              className="slicer-picker-field"
              onClick={() => {
                if (onPick(entry.field) === null) onClose()
              }}
            >
              {entry.name}
            </button>
          ))}
        </div>
        <div className="dialog-actions">
          <button className="secondary" onClick={onClose}>
            {t('dlgCancel')}
          </button>
        </div>
      </div>
    </div>
  )
}

/// Slicer overlay on the worksheet: one panel per slicer, with multi-select
/// toggle buttons for members.
export function SlicerPanels({
  slicers,
  onToggle,
  onSelectAll,
  onRemove,
}: {
  readonly slicers: readonly SlicerUiState[]
  readonly onToggle: (slicerId: string, member: number) => void
  readonly onSelectAll: (slicerId: string) => void
  readonly onRemove: (slicerId: string) => void
}): React.JSX.Element | null {
  const { t } = useI18n()
  if (slicers.length === 0) return null
  return (
    <div className="slicer-stack">
      {slicers.map((slicer) => {
        const selected = new Set(slicer.selected)
        const filtered = selected.size < slicer.members.length
        return (
          <section
            key={slicer.id}
            className="slicer-panel"
            aria-label={t('dlgSlicerAria', { name: slicer.fieldName })}
          >
            <header>
              <span className="slicer-title" data-tip={slicer.pivotPath}>
                {slicer.fieldName}
              </span>
              <button
                className="slicer-clear"
                data-tip={t('dlgSlicerClear')}
                aria-label={t('dlgSlicerClear')}
                disabled={!filtered}
                onClick={() => onSelectAll(slicer.id)}
              >
                ⧉
              </button>
              <button
                className="slicer-close"
                data-tip={t('dlgSlicerRemove')}
                aria-label={t('dlgSlicerRemove')}
                onClick={() => onRemove(slicer.id)}
              >
                ×
              </button>
            </header>
            <div className="slicer-members">
              {slicer.members.map((entry) => (
                <button
                  key={entry.member}
                  className={
                    selected.has(entry.member) ? 'slicer-member selected' : 'slicer-member'
                  }
                  onClick={() => onToggle(slicer.id, entry.member)}
                >
                  {entry.label}
                </button>
              ))}
            </div>
          </section>
        )
      })}
    </div>
  )
}
