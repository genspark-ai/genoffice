import type { FeatureRecord } from '@genoffice/agent-core'
import {
  SlotHost as SharedSlotHost,
  SlotPanel as SharedSlotPanel,
  type SlotActions,
} from '@genoffice/ui'
import { useFeatures } from './features-state'
import { useI18n } from '../locale'
// the appearance rules are shell chrome, not shared rendering: they re-point the
// shell's own variables, so they stay with the shell
import './appearance.css'

/**
 * The shell's binding of the shared slot renderer.
 *
 * The rendering, the element set and the styles live in @genoffice/ui, so every
 * app draws a slot the same way; what belongs to the shell is where the registry
 * comes from (the store below) and what its bridge can do (the actions Home
 * builds). An editor binds the same component to its own bridge.
 */

export type { SlotActions }

/** Every slot registered for one place in this shell, oldest first. */
export function SlotHost({
  place,
  actions,
}: {
  place: Parameters<typeof SharedSlotHost>[0]['place']
  actions: SlotActions
}): React.JSX.Element | null {
  const features = useFeatures()
  return <SharedSlotHost place={place} features={features} actions={actions} />
}

/** One feature's own home.panel, shown full width in the content area. */
export function SlotPanel({
  featureId,
  actions,
  onClose,
}: {
  featureId: string
  actions: SlotActions
  onClose: () => void
}): React.JSX.Element {
  const { t } = useI18n()
  const features: FeatureRecord[] = useFeatures()
  return (
    <SharedSlotPanel
      featureId={featureId}
      features={features}
      actions={actions}
      labels={{ close: t('closeTab') }}
      onClose={onClose}
    />
  )
}
