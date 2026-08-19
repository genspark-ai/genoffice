/**
 * Shape extrusion for <a:scene3d> + <a:sp3d>: builds the projected, shaded face polygons
 * PowerPoint draws for a 3D shape (front cap, back cap, side walls).
 *
 * Camera preset angles and light-rig definitions are the MS-OI29500 values (measured
 * against MS Office; same data set LibreOffice uses, see LO tdf#70039).
 * Angles rx/ry/rz are degrees; MS Office rotates first around y (lon), then x (lat),
 * then z (rev), in a left-handed frame with x right, y down, z toward the viewer.
 */

export interface Scene3DProps {
  cameraPreset: string
  cameraRot?: { lat: number; lon: number; rev: number }
  lightRig?: string
  lightDir?: string
  lightRot?: { lat: number; lon: number; rev: number }
  extrusionEmu?: number
  zEmu?: number
  extrusionColor?: string
  material?: string
}

export interface ExtrusionFaceRender {
  /** Projected face as an SVG path (local px relative to the shape box top-left; caps carry hole subpaths) */
  path: string
  /** Pre-shaded solid color ('#RRGGBB'), 'transparent' for wireframe edges */
  color: string
  /** Front cap: the adapter may substitute the shape's own (gradient/image) fill */
  front?: boolean
  stroke?: string
  strokeWidthPx?: number
}

export interface ExtrusionRender {
  /** Painter order: farthest first */
  faces: ExtrusionFaceRender[]
  /** legacyWireframe material: faces carry no fill, only edges */
  wireframe?: boolean
}

interface CameraPreset {
  parallel: boolean
  rx: number // deg, rotation around x (latitude)
  ry: number // deg, rotation around y (longitude)
  rz: number // deg, rotation around z (revolution)
  ox: number // projection origin, fraction of width from center
  oy: number
  skewAmount: number // % of depth, parallel cameras only
  skewAngle: number // deg
  vx: number // eye offset, 1/100 mm, perspective only
  vy: number
  vz: number // eye distance, 1/100 mm, perspective only
}

const D = 1 / 60000
const cam = (
  parallel: boolean,
  rx: number,
  ry: number,
  rz: number,
  ox = 0,
  oy = 0,
  skewAmount = 0,
  skewAngle = 0,
  vx = 0,
  vy = 0,
  vz = 0,
): CameraPreset => ({
  parallel,
  rx: rx * D,
  ry: ry * D,
  rz: rz * D,
  ox,
  oy,
  skewAmount,
  skewAngle,
  vx,
  vy,
  vz,
})

/** ST_PresetCameraType → rotation/projection values ([MS-OI29500], angles in 1/60000°). */
const CAMERA_PRESETS: Record<string, CameraPreset> = {
  isometricBottomDown: cam(true, 2124000, 18882000, 17988000),
  isometricBottomUp: cam(true, 2124000, 2718000, 3612000),
  isometricLeftDown: cam(true, 2100000, 2700000, 0),
  isometricLeftUp: cam(true, 19500000, 2700000, 0),
  isometricOffAxis1Left: cam(true, 1080000, 3840000, 0),
  isometricOffAxis1Right: cam(true, 1080000, 20040000, 0),
  isometricOffAxis1Top: cam(true, 18078000, 18390000, 3456000),
  isometricOffAxis2Left: cam(true, 1080000, 1560000, 0),
  isometricOffAxis2Right: cam(true, 1080000, 17760000, 0),
  isometricOffAxis2Top: cam(true, 18078000, 3210000, 18144000),
  isometricOffAxis3Bottom: cam(true, 3522000, 18390000, 18144000),
  isometricOffAxis3Left: cam(true, 20520000, 3840000, 0),
  isometricOffAxis3Right: cam(true, 20520000, 20040000, 0),
  isometricOffAxis4Bottom: cam(true, 3522000, 3210000, 3456000),
  isometricOffAxis4Left: cam(true, 20520000, 1560000, 0),
  isometricOffAxis4Right: cam(true, 20520000, 17760000, 0),
  isometricRightDown: cam(true, 19500000, 18900000, 0),
  isometricRightUp: cam(true, 2100000, 18900000, 0),
  isometricTopDown: cam(true, 19476000, 2718000, 17988000),
  isometricTopUp: cam(true, 19476000, 18882000, 3612000),
  legacyObliqueBottom: cam(true, 0, 0, 0, 0, 0.5, 50, 90),
  legacyObliqueBottomLeft: cam(true, 0, 0, 0, -0.5, 0.5, 50, 45),
  legacyObliqueBottomRight: cam(true, 0, 0, 0, 0.5, 0.5, 50, 135),
  legacyObliqueFront: cam(true, 0, 0, 0),
  legacyObliqueLeft: cam(true, 0, 0, 0, -0.5, 0, 50, -360),
  legacyObliqueRight: cam(true, 0, 0, 0, 0.5, 0, 50, 180),
  legacyObliqueTop: cam(true, 0, 0, 0, 0, -0.5, 50, -90),
  legacyObliqueTopLeft: cam(true, 0, 0, 0, -0.5, -0.5, 50, -45),
  legacyObliqueTopRight: cam(true, 0, 0, 0, 0.5, -0.5, 50, -135),
  legacyPerspectiveBottom: cam(false, 0, 0, 0, 0, 0.5, 50, 90, 0, 3472, 25000),
  legacyPerspectiveBottomLeft: cam(false, 0, 0, 0, -0.5, 0.5, 50, 45, -3472, 3472, 25000),
  legacyPerspectiveBottomRight: cam(false, 0, 0, 0, 0.5, 0.5, 50, 135, 3472, 3472, 25000),
  legacyPerspectiveFront: cam(false, 0, 0, 0, 0, 0, 0, 0, 0, 0, 25000),
  legacyPerspectiveLeft: cam(false, 0, 0, 0, -0.5, 0, 50, -360, -3472, 0, 25000),
  legacyPerspectiveRight: cam(false, 0, 0, 0, 0.5, 0, 50, 180, 3472, 0, 25000),
  legacyPerspectiveTop: cam(false, 0, 0, 0, 0, -0.5, 50, -90, 0, -3472, 25000),
  legacyPerspectiveTopLeft: cam(false, 0, 0, 0, -0.5, -0.5, 50, -45, -3472, -3472, 25000),
  legacyPerspectiveTopRight: cam(false, 0, 0, 0, 0.5, -0.5, 50, -135, 3472, -3472, 25000),
  obliqueBottom: cam(true, 0, 0, 0, 0, 0.5, 30, 90),
  obliqueBottomLeft: cam(true, 0, 0, 0, -0.5, 0.5, 30, 45),
  obliqueBottomRight: cam(true, 0, 0, 0, 0.5, 0.5, 30, 135),
  obliqueLeft: cam(true, 0, 0, 0, -0.5, 0, 30, -360),
  obliqueRight: cam(true, 0, 0, 0, 0.5, 0, 30, 180),
  obliqueTop: cam(true, 0, 0, 0, 0, -0.5, 30, -90),
  obliqueTopLeft: cam(true, 0, 0, 0, -0.5, -0.5, 30, -45),
  obliqueTopRight: cam(true, 0, 0, 0, 0.5, -0.5, 30, -135),
  orthographicFront: cam(true, 0, 0, 0),
  perspectiveAbove: cam(false, 20400000, 0, 0, 0, 0, 0, 0, 0, 0, 38451),
  perspectiveAboveLeftFacing: cam(false, 2358000, 858000, 20466000, 0, 0, 0, 0, 0, 0, 38451),
  perspectiveAboveRightFacing: cam(false, 2358000, 20742000, 1134000, 0, 0, 0, 0, 0, 0, 38451),
  perspectiveBelow: cam(false, 1200000, 0, 0, 0, 0, 0, 0, 0, 0, 38451),
  perspectiveContrastingLeftFacing: cam(false, 624000, 2634000, 21384000, 0, 0, 0, 0, 0, 0, 38451),
  perspectiveContrastingRightFacing: cam(false, 624000, 18966000, 216000, 0, 0, 0, 0, 0, 0, 38451),
  perspectiveFront: cam(false, 0, 0, 0, 0, 0, 0, 0, 0, 0, 38451),
  perspectiveHeroicExtremeLeftFacing: cam(
    false,
    486000,
    2070000,
    21426000,
    0,
    0,
    0,
    0,
    0,
    0,
    18981,
  ),
  perspectiveHeroicExtremeRightFacing: cam(
    false,
    486000,
    19530000,
    174000,
    0,
    0,
    0,
    0,
    0,
    0,
    18981,
  ),
  perspectiveHeroicLeftFacing: cam(false, 20940000, 858000, 156000, 0, 0, 0, 0, 0, 0, 38451),
  perspectiveHeroicRightFacing: cam(false, 20940000, 20742000, 21444000, 0, 0, 0, 0, 0, 0, 38451),
  perspectiveLeft: cam(false, 0, 1200000, 0, 0, 0, 0, 0, 0, 0, 38451),
  perspectiveRelaxed: cam(false, 18576000, 0, 0, 0, 0, 0, 0, 0, 0, 38451),
  perspectiveRelaxedModerately: cam(false, 19488000, 0, 0, 0, 0, 0, 0, 0, 0, 38451),
  perspectiveRight: cam(false, 0, 20400000, 0, 0, 0, 0, 0, 0, 0, 38451),
}

interface RigLight {
  r: number
  g: number
  b: number
  x: number
  y: number
  z: number
  diffuse: boolean
}

interface LightRig {
  ambient: number
  lights: RigLight[]
}

const L = (
  r: number,
  g: number,
  b: number,
  x: number,
  y: number,
  z: number,
  diffuse = true,
): RigLight => ({ r, g, b, x, y, z, diffuse })

/** ST_LightRigType → ambient + individual lights ([MS-OI29500]; direction = travel direction of the light). */
const LIGHT_RIGS: Record<string, LightRig> = {
  balanced: {
    ambient: 0.13,
    lights: [
      L(1.05, 1.05, 1.05, 0.5263, -0.4092, -0.7453),
      L(1, 1, 1, -0.9386, 0.3426, -0.041),
      L(0.5, 0.5, 0.5, 0.0934, 0.763, 0.6396),
    ],
  },
  brightRoom: {
    ambient: 1.5,
    lights: [
      L(1, 1, 1, 0, -1, 0),
      L(1, 1, 1, 0.8227, -0.1882, -0.5364, false),
      L(-0.5, -0.5, -0.5, 0, 0, -1),
      L(0.5, 0.5, 0.5, 0, 1, 0),
    ],
  },
  chilly: {
    ambient: 0.11,
    lights: [
      L(0.31, 0.32, 0.32, 0.6574, -0.7316, -0.1806),
      L(0.45, 0.45, 0.45, -0.3539, -0.1505, -0.9231),
      L(1.03, 1.02, 1.15, 0.672, -0.6185, -0.4073),
      L(0.41, 0.45, 0.48, -0.5781, 0.7976, 0.1722),
    ],
  },
  contrasting: { ambient: 1, lights: [L(1, 1, 1, 0, -1, 0, false), L(1, 1, 1, 0, 1, 0, false)] },
  flat: {
    ambient: 1,
    lights: [
      L(0.821, 0.821, 0.821, -0.9546, -0.1619, -0.2502, false),
      L(2.072, 2.54, 2.91, 0.0009, 0.8605, 0.5095, false),
      L(3.843, 3.843, 3.843, 0.6574, -0.7316, -0.1806, false),
    ],
  },
  flood: {
    ambient: 0.13,
    lights: [
      L(1.1, 1.1, 1.1, 0.5685, -0.7651, -0.3022),
      L(1.1, 1.1, 1.1, -0.2366, -0.9595, -0.1531),
      L(0.55, 0.55, 0.55, -0.8982, 0.1386, -0.4171),
    ],
  },
  freezing: {
    ambient: 0,
    lights: [
      L(0.53, 0.567, 0.661, 0.6574, -0.7316, -0.1806),
      L(0.37, 0.461, 0.461, -0.2781, -0.4509, -0.8482),
      L(0.649, 0.638, 0.904, 0.672, -0.6185, -0.4073),
      L(0.971, 1.19, 1.363, -0.1825, 0.968, 0.1722),
    ],
  },
  glow: { ambient: 1, lights: [L(1, 1, 1, 0, -1, 0), L(0.7, 0.7, 0.7, 0, 1, 0)] },
  harsh: {
    ambient: 0.28,
    lights: [
      L(0.88, 0.88, 0.88, 0.6689, -0.6755, -0.3104),
      L(0.88, 0.88, 0.88, -0.592, -0.7371, -0.326),
    ],
  },
  legacyFlat1: {
    ambient: 0.305,
    lights: [L(0.58, 0.58, 0.58, 0, 0, -0.2), L(0.29, 0.29, 0.29, 0, 0, -0.2)],
  },
  legacyFlat2: {
    ambient: 0.305,
    lights: [L(0.58, 0.58, 0.58, -1, -1, -0.2), L(0.29, 0.29, 0.29, 0, 1, -0.2)],
  },
  legacyFlat3: {
    ambient: 0.305,
    lights: [L(0.58, 0.58, 0.58, 0, -1, -0.2), L(0.29, 0.29, 0.29, 0, 1, -0.2)],
  },
  legacyFlat4: {
    ambient: 0.305,
    lights: [L(0.58, 0.58, 0.58, 1, -1, -0.2), L(0.29, 0.29, 0.29, 0, 1, -0.2)],
  },
  legacyHarsh1: {
    ambient: 0.061,
    lights: [L(0.793, 0.793, 0.793, 0, 0, -0.2), L(0.214, 0.214, 0.214, 0, 0, -0.2)],
  },
  legacyHarsh2: {
    ambient: 0.061,
    lights: [L(0.793, 0.793, 0.793, -1, -1, -0.2), L(0.214, 0.214, 0.214, 0, 1, -0.2)],
  },
  legacyHarsh3: {
    ambient: 0.061,
    lights: [L(0.793, 0.793, 0.793, 0, -1, -0.2), L(0.214, 0.214, 0.214, 0, 1, -0.2)],
  },
  legacyHarsh4: {
    ambient: 0.061,
    lights: [L(0.793, 0.793, 0.793, 1, -1, -0.2), L(0.214, 0.214, 0.214, 0, 1, -0.2)],
  },
  legacyNormal1: {
    ambient: 0.153,
    lights: [L(0.671, 0.671, 0.671, 0, 0, -0.2), L(0.183, 0.183, 0.183, 0, 0, -0.2)],
  },
  legacyNormal2: {
    ambient: 0.153,
    lights: [L(0.671, 0.671, 0.671, -1, -1, -0.2), L(0.183, 0.183, 0.183, 0, 1, -0.2)],
  },
  legacyNormal3: {
    ambient: 0.153,
    lights: [L(0.671, 0.671, 0.671, 0, -1, -0.2), L(0.183, 0.183, 0.183, 0, 1, -0.2)],
  },
  legacyNormal4: {
    ambient: 0.153,
    lights: [L(0.671, 0.671, 0.671, 1, -1, -0.2), L(0.183, 0.183, 0.183, 0, 1, -0.2)],
  },
  morning: {
    ambient: 0,
    lights: [
      L(0.669, 0.648, 0.596, 0.6574, -0.7316, -0.1806),
      L(0.459, 0.454, 0.385, -0.2781, -0.4509, -0.8482),
      L(0.9, 0.86, 0.83, 0.672, -0.6185, -0.4073),
      L(0.911, 0.846, 0.728, -0.1825, 0.968, 0.1722),
    ],
  },
  soft: { ambient: 0.3, lights: [L(0.8, 0.8, 0.8, -0.6897, 0.2484, -0.6802)] },
  sunrise: {
    ambient: 0,
    lights: [
      L(0.667, 0.63, 0.527, 0.6574, -0.7316, -0.1806),
      L(0.459, 0.459, 0.371, -0.2781, -0.4509, -0.8482),
      L(0.826, 0.712, 0.638, 0.672, -0.6185, -0.4073),
      L(1.511, 1.319, 0.994, -0.1825, 0.968, 0.1722),
    ],
  },
  sunset: {
    ambient: 0,
    lights: [
      L(0.672, 0.169, 0.169, 0.6574, -0.7316, -0.1806),
      L(0.459, 0.448, 0.327, 0.0922, -0.3551, -0.9303),
      L(0.775, 0.612, 0.502, 0.672, -0.6185, -0.4073),
      L(0.761, 0.69, 0.397, -0.424, 0.8891, 0.1722),
    ],
  },
  threePt: {
    ambient: 0,
    lights: [
      L(1.141, 1.141, 1.141, -0.6515, -0.2693, -0.7093),
      L(0.5, 0.5, 0.5, 0.8482, 0.2469, -0.4686),
      L(1, 1, 1, 0.5634, -0.2812, 0.7769),
    ],
  },
  twoPt: {
    ambient: 0.25,
    lights: [
      L(0.84, 0.84, 0.84, 0.5266, -0.4089, -0.7454),
      L(0.3, 0.3, 0.3, -0.8983, 0.2365, -0.3704),
    ],
  },
}

/** lightRig dir attribute → clockwise rig rotation around z (degrees); 't' is the preset's own orientation. */
const RIG_DIR_DEG: Record<string, number> = {
  t: 0,
  tr: 45,
  r: 90,
  br: 135,
  b: 180,
  bl: 225,
  l: 270,
  tl: 315,
}

type Vec3 = [number, number, number]
type Mat3 = [number, number, number, number, number, number, number, number, number]

const IDENT: Mat3 = [1, 0, 0, 0, 1, 0, 0, 0, 1]

function matMul(a: Mat3, b: Mat3): Mat3 {
  const o = new Array(9).fill(0) as unknown as Mat3
  for (let r = 0; r < 3; r++)
    for (let c = 0; c < 3; c++)
      o[r * 3 + c] = a[r * 3]! * b[c]! + a[r * 3 + 1]! * b[3 + c]! + a[r * 3 + 2]! * b[6 + c]!
  return o
}

function apply(m: Mat3, v: Vec3): Vec3 {
  return [
    m[0] * v[0] + m[1] * v[1] + m[2] * v[2],
    m[3] * v[0] + m[4] * v[1] + m[5] * v[2],
    m[6] * v[0] + m[7] * v[1] + m[8] * v[2],
  ]
}

const rad = (deg: number) => (deg * Math.PI) / 180

// Left-handed frame: x right, y down, z toward the viewer (matches the MS-OI29500 tables).
const rotX = (a: number): Mat3 => [
  1,
  0,
  0,
  0,
  Math.cos(a),
  Math.sin(a),
  0,
  -Math.sin(a),
  Math.cos(a),
]
const rotY = (a: number): Mat3 => [
  Math.cos(a),
  0,
  -Math.sin(a),
  0,
  1,
  0,
  Math.sin(a),
  0,
  Math.cos(a),
]
const rotZ = (a: number): Mat3 => [
  Math.cos(a),
  Math.sin(a),
  0,
  -Math.sin(a),
  Math.cos(a),
  0,
  0,
  0,
  1,
]

/** MS Office camera rotation: first lon (y), then lat (x), then rev (z). */
function cameraMatrix(latDeg: number, lonDeg: number, revDeg: number): Mat3 {
  return matMul(rotZ(rad(revDeg)), matMul(rotX(rad(latDeg)), rotY(rad(lonDeg))))
}

// EMU→px at the render viewport's 96dpi logical scale is handled by the caller; the
// perspective eye distance is defined in 1/100 mm and converted here (96 px / 25.4 mm).
const PX_PER_MM100 = 96 / 2540

function parseRgb(color: string): [number, number, number] {
  const h = color.replace(/^#/, '')
  if (/^[0-9a-fA-F]{6}([0-9a-fA-F]{2})?$/.test(h)) {
    return [parseInt(h.slice(0, 2), 16), parseInt(h.slice(2, 4), 16), parseInt(h.slice(4, 6), 16)]
  }
  const m = /rgba?\((\d+)[,\s]+(\d+)[,\s]+(\d+)/.exec(color)
  if (m) return [Number(m[1]), Number(m[2]), Number(m[3])]
  return [128, 128, 128]
}

const toHex = (n: number) =>
  Math.max(0, Math.min(255, Math.round(n)))
    .toString(16)
    .padStart(2, '0')

/** prstMaterial shading params: Blinn-Phong specular strength/exponent, diffuse/ambient scale,
 *  rim darkening (dkEdge). Calibrated against Scene3d_material_highlight PowerPoint renders. */
interface MaterialParams {
  spec: number
  shin: number
  diffuse: number
  ambient: number
  rim?: number
}
const MATERIALS: Record<string, MaterialParams> = {
  warmMatte: { spec: 0.5, shin: 3, diffuse: 1, ambient: 1 },
  matte: { spec: 0, shin: 1, diffuse: 1, ambient: 1 },
  flat: { spec: 0, shin: 1, diffuse: 1, ambient: 1 },
  legacyMatte: { spec: 0, shin: 1, diffuse: 1, ambient: 1 },
  plastic: { spec: 0.6, shin: 6, diffuse: 1, ambient: 1 },
  legacyPlastic: { spec: 0.8, shin: 4, diffuse: 1, ambient: 1 },
  metal: { spec: 0.9, shin: 4, diffuse: 1.05, ambient: 1 },
  legacyMetal: { spec: 0.9, shin: 5, diffuse: 0.55, ambient: 0.6 },
  softmetal: { spec: 0.6, shin: 4, diffuse: 1, ambient: 1 },
  dkEdge: { spec: 0.8, shin: 5, diffuse: 1, ambient: 1, rim: 0.45 },
  softEdge: { spec: 0.4, shin: 4, diffuse: 1, ambient: 1 },
  clear: { spec: 0.8, shin: 5, diffuse: 1, ambient: 1 },
  powder: { spec: 0.3, shin: 3, diffuse: 1, ambient: 1 },
  translucentPowder: { spec: 0.3, shin: 3, diffuse: 1, ambient: 1 },
}
const DEFAULT_MATERIAL = MATERIALS.warmMatte!
const LAMBERT: MaterialParams = { spec: 0, shin: 1, diffuse: 1, ambient: 1 }

/** Blinn-Phong shading of a base color by the rig's ambient + diffuse lights plus a
 *  material specular term (view direction = +z toward the viewer). */
function shade(
  base: [number, number, number],
  normal: Vec3,
  rig: LightRig,
  lightXf: Mat3,
  mat: MaterialParams = LAMBERT,
  // Specular highlights depend on the eye, so they always use the view-space normal
  // even when diffuse shading is camera-invariant (shape-space normal)
  viewNormal: Vec3 = normal,
): string {
  let fr = rig.ambient * mat.ambient
  let fg = fr
  let fb = fr
  let spec = 0
  for (const l of rig.lights) {
    // MS-OI29500 light directions are in MSO's rig frame; map into the screen frame
    // (x right, y down, z toward viewer): (x,y,z) → (-y, x, z). Verified against the
    // harsh rig's lit/dark walls in PowerPoint renders.
    const dir = apply(lightXf, [-l.y, l.x, l.z])
    const len = Math.hypot(dir[0], dir[1], dir[2]) || 1
    // The table stores the light's travel direction; illumination uses the vector toward the light.
    const lx = -dir[0] / len
    const ly = -dir[1] / len
    const lz = -dir[2] / len
    const d = Math.max(0, normal[0] * lx + normal[1] * ly + normal[2] * lz)
    if (l.diffuse) {
      fr += d * l.r * mat.diffuse
      fg += d * l.g * mat.diffuse
      fb += d * l.b * mat.diffuse
    }
    if (mat.spec > 0) {
      // Blinn-Phong half vector against the straight-on viewer
      const hl = Math.hypot(lx, ly, lz + 1) || 1
      const nh = Math.max(
        0,
        (viewNormal[0] * lx + viewNormal[1] * ly + viewNormal[2] * (lz + 1)) / hl,
      )
      spec += Math.pow(nh, mat.shin) * mat.spec * Math.max(l.r, l.g, l.b)
    }
  }
  if (mat.rim) {
    // dkEdge: faces turning away from the viewer darken toward the silhouette
    const facing = Math.min(1, Math.abs(viewNormal[2]))
    const k = mat.rim + (1 - mat.rim) * facing
    fr *= k
    fg *= k
    fb *= k
  }
  const s = spec * 255
  return `#${toHex(base[0] * fr + s)}${toHex(base[1] * fg + s)}${toHex(base[2] * fb + s)}`
}

/**
 * Flatten an absolute M/L/C/Q/Z SVG path (the only commands preset/custGeom paths emit)
 * into closed polygon rings; curves are sampled uniformly.
 */
export function flattenSvgPath(d: string, curveSegs = 10): number[][] {
  const toks = d.trim().split(/[\s,]+/)
  const rings: number[][] = []
  let ring: number[] = []
  let x = 0
  let y = 0
  let i = 0
  const num = () => Number(toks[i++])
  const closeRing = () => {
    if (ring.length >= 6) rings.push(ring)
    ring = []
  }
  while (i < toks.length) {
    const t = toks[i++]
    switch (t) {
      case 'M':
        closeRing()
        x = num()
        y = num()
        ring.push(x, y)
        break
      case 'L':
        x = num()
        y = num()
        ring.push(x, y)
        break
      case 'C': {
        const c1x = num()
        const c1y = num()
        const c2x = num()
        const c2y = num()
        const ex = num()
        const ey = num()
        for (let k = 1; k <= curveSegs; k++) {
          const u = k / curveSegs
          const v = 1 - u
          ring.push(
            v * v * v * x + 3 * v * v * u * c1x + 3 * v * u * u * c2x + u * u * u * ex,
            v * v * v * y + 3 * v * v * u * c1y + 3 * v * u * u * c2y + u * u * u * ey,
          )
        }
        x = ex
        y = ey
        break
      }
      case 'Q': {
        const cx1 = num()
        const cy1 = num()
        const ex = num()
        const ey = num()
        for (let k = 1; k <= curveSegs; k++) {
          const u = k / curveSegs
          const v = 1 - u
          ring.push(
            v * v * x + 2 * v * u * cx1 + u * u * ex,
            v * v * y + 2 * v * u * cy1 + u * u * ey,
          )
        }
        x = ex
        y = ey
        break
      }
      case 'Z':
      case 'z':
        closeRing()
        break
      default:
        // Bare coordinate pair after M/L (SVG shorthand) — treat as L.
        if (t !== undefined && Number.isFinite(Number(t))) {
          x = Number(t)
          y = num()
          ring.push(x, y)
        }
        break
    }
  }
  closeRing()
  // Drop consecutive duplicate points (flattened curve endpoints repeat corners).
  return rings.map((r) => {
    const o: number[] = []
    for (let k = 0; k < r.length; k += 2) {
      const n = o.length
      if (n >= 2 && Math.abs(o[n - 2]! - r[k]!) < 0.01 && Math.abs(o[n - 1]! - r[k + 1]!) < 0.01)
        continue
      o.push(r[k]!, r[k + 1]!)
    }
    while (
      o.length >= 4 &&
      Math.abs(o[0]! - o[o.length - 2]!) < 0.01 &&
      Math.abs(o[1]! - o[o.length - 1]!) < 0.01
    )
      o.splice(o.length - 2, 2)
    return o
  })
}

/** Sampled ellipse ring (local px). */
export function ellipseRing(w: number, h: number, segs = 48): number[] {
  const out: number[] = []
  for (let k = 0; k < segs; k++) {
    const a = (k / segs) * 2 * Math.PI
    out.push((w / 2) * (1 + Math.cos(a)), (h / 2) * (1 + Math.sin(a)))
  }
  return out
}

/** Rounded-rect ring (local px), r clamped to half the short side. */
export function roundRectRing(w: number, h: number, r: number, segsPerCorner = 6): number[] {
  const rr = Math.max(0, Math.min(r, w / 2, h / 2))
  if (rr < 0.5) return [0, 0, w, 0, w, h, 0, h]
  const out: number[] = []
  const corner = (cx: number, cy: number, a0: number) => {
    for (let k = 0; k <= segsPerCorner; k++) {
      const a = a0 + (k / segsPerCorner) * (Math.PI / 2)
      out.push(cx + rr * Math.cos(a), cy + rr * Math.sin(a))
    }
  }
  corner(rr, rr, Math.PI) // top-left, sweeping to top edge
  corner(w - rr, rr, -Math.PI / 2)
  corner(w - rr, h - rr, 0)
  corner(rr, h - rr, Math.PI / 2)
  return out
}

/** Even-odd point-in-rings test (local px). */
function insideRings(rings: number[][], x: number, y: number): boolean {
  let inside = false
  for (const r of rings) {
    for (let i = 0, j = r.length - 2; i < r.length; j = i, i += 2) {
      const xi = r[i]!
      const yi = r[i + 1]!
      const xj = r[j]!
      const yj = r[j + 1]!
      if (yi > y !== yj > y && x < ((xj - xi) * (y - yi)) / (yj - yi) + xi) inside = !inside
    }
  }
  return inside
}

interface Face3 {
  verts: Vec3[]
  normal: Vec3
  kind: 'front' | 'back' | 'side'
}

export interface BuildExtrusionInput {
  /** Flattened closed outline rings (local px; holes as additional rings, even-odd) */
  rings: number[][]
  w: number
  h: number
  depthPx: number
  /** sp3d z: whole-shape shift toward the viewer (px) */
  zPx: number
  scene: Scene3DProps
  /** Resolved solid front color (callers reduce gradients to a mid color) */
  frontColor: string
  sideColor: string
  strokeColor?: string
  strokeWidthPx?: number
  /** Front cap keeps the shape's own non-solid fill (adapter substitutes it) */
  frontUsesFill?: boolean
}

/**
 * Camera with no extrusion and only an in-plane rotation (rev around z): the shape is
 * simply rotated in 2D — callers should add this to the node rotation instead of meshing.
 */
export function inPlaneRotationDeg(scene: Scene3DProps): number | null {
  const preset = CAMERA_PRESETS[scene.cameraPreset]
  if (!preset) return null
  if ((scene.extrusionEmu ?? 0) > 0) return null
  // A z shift changes the projection (perspective scale, oblique skew offset) — mesh it.
  if ((scene.zEmu ?? 0) !== 0) return null
  const lat = scene.cameraRot ? scene.cameraRot.lat * D : preset.rx
  const lon = scene.cameraRot ? scene.cameraRot.lon * D : preset.ry
  const rev = scene.cameraRot ? scene.cameraRot.rev * D : preset.rz
  const near = (a: number) =>
    Math.abs(((a % 360) + 360) % 360) < 0.5 || Math.abs((((a % 360) + 360) % 360) - 360) < 0.5
  if (!near(lat) || !near(lon)) return null
  // MS Office camera revolution turns the scene clockwise-positive in screen terms.
  return rev === 0 ? 0 : -rev
}

export function buildExtrusion(input: BuildExtrusionInput): ExtrusionRender | null {
  const { rings, w, h, depthPx, scene } = input
  const preset = CAMERA_PRESETS[scene.cameraPreset]
  // depth 0 is legal: a flat shape tilted by the camera projects to just its front cap
  if (!preset || !rings.length || depthPx < 0) return null

  const lat = scene.cameraRot ? scene.cameraRot.lat * D : preset.rx
  const lon = scene.cameraRot ? scene.cameraRot.lon * D : preset.ry
  const rev = scene.cameraRot ? scene.cameraRot.rev * D : preset.rz
  const m = cameraMatrix(lat, lon, rev)

  const cx = w / 2
  const cy = h / 2
  const zOff = input.zPx

  // 3D mesh: front cap at z=zOff, back cap at z=zOff-depth, one wall per outline edge.
  const faces: Face3[] = []
  const frontRings: Vec3[][] = []
  const backRings: Vec3[][] = []
  for (const ring of rings) {
    const fr: Vec3[] = []
    const br: Vec3[] = []
    for (let i = 0; i < ring.length; i += 2) {
      fr.push([ring[i]! - cx, ring[i + 1]! - cy, zOff])
      br.push([ring[i]! - cx, ring[i + 1]! - cy, zOff - depthPx])
    }
    frontRings.push(fr)
    backRings.push(br)
    if (depthPx <= 0) continue
    for (let i = 0; i < fr.length; i++) {
      const j = (i + 1) % fr.length
      const ex = fr[j]![0] - fr[i]![0]
      const ey = fr[j]![1] - fr[i]![1]
      const elen = Math.hypot(ex, ey)
      if (elen < 1e-6) continue
      // Outward 2D wall normal: probe both perpendiculars against the even-odd interior.
      let nx = ey / elen
      let ny = -ex / elen
      const mx = (fr[i]![0] + fr[j]![0]) / 2 + cx
      const my = (fr[i]![1] + fr[j]![1]) / 2 + cy
      if (insideRings(rings, mx + nx * 0.75, my + ny * 0.75)) {
        nx = -nx
        ny = -ny
      }
      faces.push({ verts: [fr[i]!, fr[j]!, br[j]!, br[i]!], normal: [nx, ny, 0], kind: 'side' })
    }
  }

  // Rotate the mesh (normals rotate with it; the camera matrix is orthonormal). The
  // pre-rotation normal is kept for shading: with modern cameras the light rig follows
  // the camera, so illumination is camera-invariant (evaluated in shape space).
  const shapeNormals = faces.map((f) => f.normal)
  for (const f of faces) {
    f.verts = f.verts.map((v) => apply(m, v))
    f.normal = apply(m, f.normal)
  }
  const frontR = frontRings.map((r) => r.map((v) => apply(m, v)))
  const backR = backRings.map((r) => r.map((v) => apply(m, v)))
  const frontNormal = apply(m, [0, 0, 1])
  const backNormal = apply(m, [0, 0, -1])

  // Projection to screen px (still centered at the shape center).
  const eyeDist = preset.vz * PX_PER_MM100
  const eyeX = preset.vx * PX_PER_MM100
  const eyeY = preset.vy * PX_PER_MM100
  const skew = preset.skewAmount / 100
  const skewA = rad(preset.skewAngle)
  const origX = preset.ox * w
  const origY = preset.oy * h
  const project = (v: Vec3): [number, number] => {
    if (preset.parallel) {
      // Oblique parallel projection: depth leaks into xy along the skew angle.
      return [v[0] - v[2] * skew * Math.cos(skewA) + cx, v[1] - v[2] * skew * Math.sin(skewA) + cy]
    }
    const s = eyeDist / Math.max(eyeDist - (v[2] - 0), 1)
    return [
      origX + eyeX + (v[0] - origX - eyeX) * s + cx,
      origY + eyeY + (v[1] - origY - eyeY) * s + cy,
    ]
  }

  const rig = LIGHT_RIGS[scene.lightRig ?? ''] ?? LIGHT_RIGS.threePt!
  let lightXf: Mat3 = IDENT
  if (scene.lightRot) {
    // Measured on Scene3d_material_highlight (twoPt rev=90°): PowerPoint keeps the rig's
    // table directions — a rev spin does not move the lights. lat/lon still reposition
    // the rig (inverse, like the dir attribute's z-steps).
    const c = cameraMatrix(scene.lightRot.lat * D, scene.lightRot.lon * D, 0)
    lightXf = [c[0], c[3], c[6], c[1], c[4], c[7], c[2], c[5], c[8]]
  } else {
    const dirDeg = RIG_DIR_DEG[scene.lightDir ?? 't'] ?? 0
    if (dirDeg) lightXf = rotZ(rad(dirDeg))
  }

  const frontBase = parseRgb(input.frontColor)
  const sideBase = parseRgb(scene.extrusionColor ?? input.sideColor)
  const wireframe = scene.material === 'legacyWireframe'
  // Material shading only for an explicit prstMaterial: the default (absent) keeps the
  // plain Lambert look every calibrated deck was measured with
  const mat = scene.material ? (MATERIALS[scene.material] ?? DEFAULT_MATERIAL) : LAMBERT
  // Caps keep the authored look — PowerPoint puts the material highlight on the walls only
  const capMat = { ...mat, spec: 0 }

  const out: ExtrusionFaceRender[] = []
  const rnd = (n: number) => Math.round(n * 100) / 100
  const ringPath = (verts: Vec3[]): string => {
    const pts = verts.map(project)
    return `M ${pts.map((p) => `${rnd(p[0])} ${rnd(p[1])}`).join(' L ')} Z`
  }
  const capPath = (ringsOf: Vec3[][]): string => ringsOf.map(ringPath).join(' ')
  const meanZ = (verts: Vec3[]) => verts.reduce((s, v) => s + v[2], 0) / verts.length

  if (wireframe) {
    const sw = Math.max(input.strokeWidthPx ?? 1, 0.75)
    const stroke = input.strokeColor ?? '#000000'
    out.push({ path: capPath(backR), color: 'transparent', stroke, strokeWidthPx: sw })
    // Connecting edges only at real corners (flattened curves would spray lines everywhere).
    for (let ri = 0; ri < frontR.length; ri++) {
      const fr = frontR[ri]!
      const br = backR[ri]!
      for (let i = 0; i < fr.length; i++) {
        const prev = fr[(i - 1 + fr.length) % fr.length]!
        const next = fr[(i + 1) % fr.length]!
        const a1 = Math.atan2(fr[i]![1] - prev[1], fr[i]![0] - prev[0])
        const a2 = Math.atan2(next[1] - fr[i]![1], next[0] - fr[i]![0])
        let da = Math.abs(a1 - a2)
        if (da > Math.PI) da = 2 * Math.PI - da
        if (da < rad(20)) continue
        const p1 = project(fr[i]!)
        const p2 = project(br[i]!)
        out.push({
          path: `M ${rnd(p1[0])} ${rnd(p1[1])} L ${rnd(p2[0])} ${rnd(p2[1])}`,
          color: 'transparent',
          stroke,
          strokeWidthPx: sw,
        })
      }
    }
    out.push({ path: capPath(frontR), color: 'transparent', stroke, strokeWidthPx: sw })
    return { faces: out, wireframe: true }
  }

  // Painter order: back cap, walls far→near, front cap. Walls facing away from the
  // viewer are culled (interior walls seen through holes survive: they face the viewer).
  const viewDirAt = (verts: Vec3[]): Vec3 => {
    // Skewed parallel projection: the projector direction leans along the skew, so
    // side walls of an unrotated oblique box stay visible.
    if (preset.parallel) return [skew * Math.cos(skewA), skew * Math.sin(skewA), 1]
    const cxv = verts.reduce((s, v) => s + v[0], 0) / verts.length
    const cyv = verts.reduce((s, v) => s + v[1], 0) / verts.length
    const czv = verts.reduce((s, v) => s + v[2], 0) / verts.length
    return [eyeX + origX - cxv, eyeY + origY - cyv, eyeDist - czv]
  }
  const facing = (normal: Vec3, verts: Vec3[]) => {
    const vd = viewDirAt(verts)
    return normal[0] * vd[0] + normal[1] * vd[1] + normal[2] * vd[2] > 0
  }

  // Skewed parallel cameras (oblique/legacy) rotate the shape under fixed lights (shade
  // with view-space normals); rotating cameras (isometric/perspective/orthographic) move
  // the camera with the rig attached (shade with shape-space normals, so the authored
  // appearance survives every preset and the back cap goes dark).
  const legacy = preset.skewAmount > 0 || scene.cameraPreset.startsWith('legacy')
  if (backR[0] && facing(backNormal, backR[0])) {
    out.push({
      path: capPath(backR),
      color: legacy
        ? shade(sideBase, backNormal, rig, lightXf, capMat)
        : shade(frontBase, [0, 0, -1], rig, lightXf, capMat, backNormal),
    })
  }
  const wallIdx = faces
    .map((f, i) => i)
    .filter((i) => facing(faces[i]!.normal, faces[i]!.verts))
    .sort((a, b) => meanZ(faces[a]!.verts) - meanZ(faces[b]!.verts))
  for (const i of wallIdx) {
    const f = faces[i]!
    out.push({
      path: ringPath(f.verts),
      color: shade(sideBase, legacy ? f.normal : shapeNormals[i]!, rig, lightXf, mat, f.normal),
    })
  }
  if (frontR[0] && facing(frontNormal, frontR[0])) {
    out.push({
      path: capPath(frontR),
      color: input.frontUsesFill
        ? input.frontColor
        : shade(frontBase, legacy ? frontNormal : [0, 0, 1], rig, lightXf, capMat, frontNormal),
      ...(input.frontUsesFill ? { front: true } : {}),
      ...(input.strokeColor
        ? { stroke: input.strokeColor, strokeWidthPx: input.strokeWidthPx ?? 1 }
        : {}),
    })
  }
  return { faces: out }
}
