import { OverlayLayer as SharedOverlayLayer } from '@genoffice/ui'
import { useFeatures } from './features-state'

/** The shell's binding of the shared overlay layer: the store in, the layer out. */
export function OverlayLayer(): React.JSX.Element | null {
  const features = useFeatures()
  return <SharedOverlayLayer features={features} />
}
