/**
 * The bounding box of a glTF binary (.glb) scene, from its JSON chunk alone: each mesh's
 * POSITION accessor min/max, carried through the node hierarchy of the default scene.
 * PowerPoint frames a 3D model from this box (am3d:meterPerModelUnit, preTrans, camera).
 */
type Vec3 = [number, number, number]
type Mat4 = number[]

export interface GlbBounds {
  center: Vec3
  size: Vec3
}

interface GltfNode {
  mesh?: number
  children?: number[]
  matrix?: number[]
  translation?: number[]
  rotation?: number[]
  scale?: number[]
}

interface GltfJson {
  scene?: number
  scenes?: Array<{ nodes?: number[] }>
  nodes?: GltfNode[]
  meshes?: Array<{ primitives?: Array<{ attributes?: { POSITION?: number } }> }>
  accessors?: Array<{ min?: number[]; max?: number[] }>
}

const GLB_MAGIC = 0x46546c67
const JSON_CHUNK = 0x4e4f534a
const IDENTITY: Mat4 = [1, 0, 0, 0, 0, 1, 0, 0, 0, 0, 1, 0, 0, 0, 0, 1]

/** The JSON chunk of a .glb, or null when the bytes are not glTF binary. */
export function glbJson(bytes: Uint8Array): GltfJson | null {
  if (bytes.byteLength < 20) return null
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  if (view.getUint32(0, true) !== GLB_MAGIC || view.getUint32(16, true) !== JSON_CHUNK) return null
  const length = view.getUint32(12, true)
  if (20 + length > bytes.byteLength) return null
  try {
    return JSON.parse(new TextDecoder().decode(bytes.subarray(20, 20 + length))) as GltfJson
  } catch {
    return null
  }
}

function localMatrix(node: GltfNode): Mat4 {
  if (node.matrix?.length === 16) return node.matrix
  const [tx, ty, tz] = node.translation ?? [0, 0, 0]
  const [qx, qy, qz, qw] = node.rotation ?? [0, 0, 0, 1]
  const [sx, sy, sz] = node.scale ?? [1, 1, 1]
  const x2 = qx! + qx!
  const y2 = qy! + qy!
  const z2 = qz! + qz!
  const xx = qx! * x2
  const xy = qx! * y2
  const xz = qx! * z2
  const yy = qy! * y2
  const yz = qy! * z2
  const zz = qz! * z2
  const wx = qw! * x2
  const wy = qw! * y2
  const wz = qw! * z2
  return [
    (1 - yy - zz) * sx!,
    (xy + wz) * sx!,
    (xz - wy) * sx!,
    0,
    (xy - wz) * sy!,
    (1 - xx - zz) * sy!,
    (yz + wx) * sy!,
    0,
    (xz + wy) * sz!,
    (yz - wx) * sz!,
    (1 - xx - yy) * sz!,
    0,
    tx!,
    ty!,
    tz!,
    1,
  ]
}

/** a × b, both column-major. */
function multiply(a: Mat4, b: Mat4): Mat4 {
  const out = new Array<number>(16)
  for (let col = 0; col < 4; col++) {
    for (let row = 0; row < 4; row++) {
      out[col * 4 + row] =
        a[row]! * b[col * 4]! +
        a[4 + row]! * b[col * 4 + 1]! +
        a[8 + row]! * b[col * 4 + 2]! +
        a[12 + row]! * b[col * 4 + 3]!
    }
  }
  return out
}

export function glbBounds(bytes: Uint8Array): GlbBounds | null {
  const gltf = glbJson(bytes)
  if (!gltf) return null
  const meshBox = (index: number): [Vec3, Vec3] | null => {
    let found: [Vec3, Vec3] | null = null
    for (const primitive of gltf.meshes?.[index]?.primitives ?? []) {
      const accessor = gltf.accessors?.[primitive.attributes?.POSITION ?? -1]
      if (accessor?.min?.length !== 3 || accessor.max?.length !== 3) continue
      const box: [Vec3, Vec3] = found ?? [
        [Infinity, Infinity, Infinity],
        [-Infinity, -Infinity, -Infinity],
      ]
      const [lo, hi] = box
      for (let i = 0; i < 3; i++) {
        lo[i] = Math.min(lo[i]!, accessor.min[i]!)
        hi[i] = Math.max(hi[i]!, accessor.max[i]!)
      }
      found = [lo, hi]
    }
    return found
  }

  const min: Vec3 = [Infinity, Infinity, Infinity]
  const max: Vec3 = [-Infinity, -Infinity, -Infinity]
  const visited = new Set<number>()
  const visit = (index: number, parent: Mat4) => {
    const node = gltf.nodes?.[index]
    if (!node || visited.has(index)) return
    visited.add(index)
    const world = multiply(parent, localMatrix(node))
    const box = node.mesh !== undefined ? meshBox(node.mesh) : null
    if (box) {
      for (const x of [box[0][0], box[1][0]]) {
        for (const y of [box[0][1], box[1][1]]) {
          for (const z of [box[0][2], box[1][2]]) {
            const p: Vec3 = [
              world[0]! * x + world[4]! * y + world[8]! * z + world[12]!,
              world[1]! * x + world[5]! * y + world[9]! * z + world[13]!,
              world[2]! * x + world[6]! * y + world[10]! * z + world[14]!,
            ]
            for (let i = 0; i < 3; i++) {
              min[i] = Math.min(min[i]!, p[i]!)
              max[i] = Math.max(max[i]!, p[i]!)
            }
          }
        }
      }
    }
    for (const child of node.children ?? []) visit(child, world)
  }
  const scene = gltf.scenes?.[gltf.scene ?? 0]
  for (const index of scene?.nodes ?? []) visit(index, IDENTITY)
  if (!Number.isFinite(min[0])) return null
  return {
    center: [0, 1, 2].map((i) => (min[i]! + max[i]!) / 2) as Vec3,
    size: [0, 1, 2].map((i) => max[i]! - min[i]!) as Vec3,
  }
}
