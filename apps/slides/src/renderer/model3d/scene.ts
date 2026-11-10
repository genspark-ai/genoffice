/**
 * A glTF binary drawn as PowerPoint frames a 3D model (pptx-engine/model3d.ts): scaled to one
 * meter across and centred, turned by am3d:rot, seen through a 45° perspective camera that
 * spans the frame's longer side, under PowerPoint's default lights. three.js loads on first use.
 */
import type { Group, PerspectiveCamera, Scene, WebGLRenderer } from 'three'
import {
  MODEL3D_EMU_PER_METER,
  MODEL3D_FOV,
  MODEL3D_LIGHTS,
  MODEL3D_RATIO_ONE,
  model3dFraming,
  type Model3dRotation,
} from '@genoffice/pptx-engine/model3d'

/**
 * am3d light levels are photometric (lux, candela) for Office's renderer; these scales map them
 * to three.js light intensities, set by eye on textured and untextured models.
 */
const AMBIENT_SCALE = 2
const POINT_SCALE = 0.25
/** am3d:rot units */
const UNITS_PER_DEGREE = 60000
/** A rendering stored in the file is at most this many pixels on its longer side. */
export const MAX_RENDERING_PX = 1600

export interface ModelView {
  canvas: HTMLCanvasElement
  /** Width / height of the model seen from the front, unturned */
  frontAspect: number
  /** Draw at `width` × `height` CSS pixels, turned by `rot`. */
  draw(width: number, height: number, rot: Model3dRotation): void
  /** A PNG of the model at `rot`, `width` × `height` pixels. */
  png(width: number, height: number, rot: Model3dRotation): Promise<Uint8Array>
  dispose(): void
}

const toRadians = (units: number) => ((units / UNITS_PER_DEGREE) * Math.PI) / 180

export async function loadModel(
  bytes: Uint8Array,
  canvas = document.createElement('canvas'),
): Promise<ModelView> {
  const THREE = await import('three')
  const { GLTFLoader } = await import('three/examples/jsm/loaders/GLTFLoader.js')
  const buffer = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength)
  const gltf = await new GLTFLoader().parseAsync(buffer as ArrayBuffer, '')

  const box = new THREE.Box3().setFromObject(gltf.scene)
  const center = box.getCenter(new THREE.Vector3())
  const size = box.getSize(new THREE.Vector3())
  const framing = model3dFraming(
    box.isEmpty() ? null : { center: center.toArray(), size: size.toArray() },
  )
  const maxExtent = Math.max(size.x, size.y, size.z) || 1

  const scene: Scene = new THREE.Scene()
  const turned: Group = new THREE.Group()
  const model = gltf.scene
  model.scale.setScalar(1 / maxExtent)
  model.position.set(
    ...(framing.preTrans.map((d) => d / MODEL3D_EMU_PER_METER) as [number, number, number]),
  )
  turned.add(model)
  scene.add(turned)

  const linear = (rgb: readonly number[]) =>
    new THREE.Color().setRGB(
      rgb[0]! / 100000,
      rgb[1]! / 100000,
      rgb[2]! / 100000,
      THREE.LinearSRGBColorSpace,
    )
  const { ambient, points } = MODEL3D_LIGHTS
  scene.add(
    new THREE.AmbientLight(
      linear(ambient.rgb),
      (ambient.illuminance / MODEL3D_RATIO_ONE) * AMBIENT_SCALE,
    ),
  )
  for (const light of points) {
    const directional = new THREE.DirectionalLight(
      linear(light.rgb),
      (light.intensity / MODEL3D_RATIO_ONE) * POINT_SCALE,
    )
    directional.position.set(
      ...(light.pos.map((p) => p / MODEL3D_EMU_PER_METER) as [number, number, number]),
    )
    scene.add(directional)
  }

  const camera: PerspectiveCamera = new THREE.PerspectiveCamera()
  camera.position.set(0, 0, framing.cameraZ / MODEL3D_EMU_PER_METER)
  camera.lookAt(0, 0, 0)
  const renderer: WebGLRenderer = new THREE.WebGLRenderer({
    canvas,
    alpha: true,
    antialias: true,
    preserveDrawingBuffer: true,
  })

  const draw = (width: number, height: number, rot: Model3dRotation, pixelRatio: number) => {
    turned.rotation.set(toRadians(rot.ax), toRadians(rot.ay), toRadians(rot.az), 'XYZ')
    // the field of view spans the longer side, as am3d:objViewport's size is the frame's longer side
    const fov = MODEL3D_FOV / UNITS_PER_DEGREE
    camera.aspect = width / height
    camera.fov =
      width >= height
        ? (2 * Math.atan(Math.tan((fov * Math.PI) / 360) * (height / width)) * 180) / Math.PI
        : fov
    camera.near = 0.01
    camera.far = 100
    camera.updateProjectionMatrix()
    renderer.setPixelRatio(pixelRatio)
    renderer.setSize(width, height, false)
    renderer.render(scene, camera)
  }

  return {
    canvas,
    frontAspect: size.y > 0 ? size.x / size.y : 1,
    draw: (width, height, rot) => draw(width, height, rot, window.devicePixelRatio || 1),
    async png(width, height, rot) {
      draw(width, height, rot, 1)
      const blob = await new Promise<Blob | null>((resolve) => canvas.toBlob(resolve, 'image/png'))
      if (!blob) throw new Error('3D model rendering failed')
      return new Uint8Array(await blob.arrayBuffer())
    },
    dispose() {
      renderer.dispose()
      scene.traverse((node) => {
        const mesh = node as { geometry?: { dispose(): void }; material?: unknown }
        mesh.geometry?.dispose()
        const materials = Array.isArray(mesh.material) ? mesh.material : [mesh.material]
        for (const material of materials)
          (material as { dispose?(): void } | undefined)?.dispose?.()
      })
    },
  }
}

/** The size of a stored rendering for a frame of `width` × `height` slide pixels. */
export function renderingSize(width: number, height: number): { width: number; height: number } {
  const scale = Math.min(2, MAX_RENDERING_PX / Math.max(width, height, 1))
  return {
    width: Math.max(1, Math.round(width * scale)),
    height: Math.max(1, Math.round(height * scale)),
  }
}
