import { NO_ROTATION, type Model3dRotation } from '@genoffice/pptx-engine/model3d'
import type { StringKey } from '../i18n/locale'

/** am3d:rot units */
const DEG = 60000

/** PowerPoint's 3D Model Views: the side of the model that faces the camera. */
export const MODEL3D_VIEWS: ReadonlyArray<{ label: StringKey; rot: Model3dRotation }> = [
  { label: 'ribbonModel3dFront', rot: NO_ROTATION },
  { label: 'ribbonModel3dBack', rot: { ax: 0, ay: 180 * DEG, az: 0 } },
  { label: 'ribbonModel3dLeft', rot: { ax: 0, ay: 90 * DEG, az: 0 } },
  { label: 'ribbonModel3dRight', rot: { ax: 0, ay: -90 * DEG, az: 0 } },
  { label: 'ribbonModel3dTop', rot: { ax: 90 * DEG, ay: 0, az: 0 } },
  { label: 'ribbonModel3dBottom', rot: { ax: -90 * DEG, ay: 0, az: 0 } },
  { label: 'ribbonModel3dAboveFrontLeft', rot: { ax: 30 * DEG, ay: 45 * DEG, az: 0 } },
]
