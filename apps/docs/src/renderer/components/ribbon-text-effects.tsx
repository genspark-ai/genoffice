import { TEXT_FX_PRESETS, textFxDecls } from '../editor/text-fx'
import { useI18n } from '../i18n/locale'
import { IconCaret, IconTextEffects } from './icons'
import { useRibbonCommand, useRunCommand } from './ribbon-command'
import { styleOf } from './TextEffectsDialog'

/** Home > Text Effects: the ready-made looks, and the way into the dialog */
export function RbTextEffects({
  open,
  onToggle,
  onClose,
}: {
  open: boolean
  onToggle: () => void
  onClose: () => void
}) {
  const { t } = useI18n()
  const command = useRibbonCommand('format-text-effects')
  const runCommand = useRunCommand()
  const tip = t(command.control.tip)
  const pick = (id: 'text-effects-preset' | 'format-text-effects', payload?: string) => {
    runCommand(id, payload)
    onClose()
  }
  return (
    <div className="rb-split-wrap">
      <button
        className="rb-icon"
        disabled={command.disabled}
        data-tip={tip}
        aria-label={tip}
        // command-exempt: opens this control's own gallery
        onClick={onToggle}
      >
        <IconTextEffects />
        <span className="rb-caret-inline">
          <IconCaret />
        </span>
      </button>
      {open && (
        <div data-rb-panel="" className="spacing-menu text-fx-gallery">
          {TEXT_FX_PRESETS.map((item, index) => (
            <button
              key={item.id}
              className="text-fx-tile"
              data-preset={item.id}
              aria-label={`${t('ribbonColPresets')} ${index + 1}`}
              onClick={() => pick('text-effects-preset', item.id)}
            >
              {/* a document colour, not chrome */}
              <span style={{ ...styleOf(textFxDecls(item.fx)), color: `#${item.color}` }}>A</span>
            </button>
          ))}
          <button className="text-fx-more" onClick={() => pick('format-text-effects')}>
            {t('ribbonTextEffects')}…
          </button>
        </div>
      )}
    </div>
  )
}
