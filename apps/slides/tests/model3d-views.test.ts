/** PowerPoint's 3D Model Views: each turns the named side of the model towards the camera. */
import { describe, expect, it } from 'vitest'
import { Euler, Vector3 } from 'three'
import { MODEL3D_VIEWS } from '../src/renderer/model3d/views'
import { renderingSize, MAX_RENDERING_PX } from '../src/renderer/model3d/scene'

const toward = (label: string) => {
  const { rot } = MODEL3D_VIEWS.find((v) => v.label === label)!
  const rad = (units: number) => ((units / 60000) * Math.PI) / 180
  return new Euler(rad(rot.ax), rad(rot.ay), rad(rot.az), 'XYZ')
}
/** The model-space direction that ends up facing the camera (+z) once turned. */
const facing = (label: string, side: Vector3) =>
  side
    .clone()
    .applyEuler(toward(label))
    .dot(new Vector3(0, 0, 1))

describe('3D Model Views', () => {
  it.each([
    ['ribbonModel3dFront', new Vector3(0, 0, 1)],
    ['ribbonModel3dBack', new Vector3(0, 0, -1)],
    ['ribbonModel3dLeft', new Vector3(-1, 0, 0)],
    ['ribbonModel3dRight', new Vector3(1, 0, 0)],
    ['ribbonModel3dTop', new Vector3(0, 1, 0)],
    ['ribbonModel3dBottom', new Vector3(0, -1, 0)],
  ])('%s faces the camera', (label, side) => {
    expect(facing(label, side)).toBeCloseTo(1)
  })

  it('stores renderings at twice the frame, capped on the longer side', () => {
    expect(renderingSize(300, 200)).toEqual({ width: 600, height: 400 })
    expect(renderingSize(2000, 1000)).toEqual({
      width: MAX_RENDERING_PX,
      height: MAX_RENDERING_PX / 2,
    })
  })
})
