import { PNG } from 'pngjs'
import { readFileSync } from 'node:fs'

const [a, b] = process.argv.slice(2)
const pa = PNG.sync.read(readFileSync(a))
const pb = PNG.sync.read(readFileSync(b))
if (pa.width !== pb.width || pa.height !== pb.height) {
  console.log(`size differs: ${pa.width}x${pa.height} vs ${pb.width}x${pb.height}`)
  process.exit(0)
}
let differing = 0
let minX = Infinity, minY = Infinity, maxX = -1, maxY = -1
const rows = new Map()
for (let y = 0; y < pa.height; y++) {
  for (let x = 0; x < pa.width; x++) {
    const i = (y * pa.width + x) * 4
    if (
      pa.data[i] !== pb.data[i] ||
      pa.data[i + 1] !== pb.data[i + 1] ||
      pa.data[i + 2] !== pb.data[i + 2] ||
      pa.data[i + 3] !== pb.data[i + 3]
    ) {
      differing++
      rows.set(y, (rows.get(y) ?? 0) + 1)
      if (x < minX) minX = x
      if (x > maxX) maxX = x
      if (y < minY) minY = y
      if (y > maxY) maxY = y
    }
  }
}
const total = pa.width * pa.height
console.log(`differing pixels: ${differing} / ${total} (${((differing / total) * 100).toFixed(3)}%)`)
if (differing) {
  console.log(`bounding box: x ${minX}..${maxX}  y ${minY}..${maxY}  (image ${pa.width}x${pa.height})`)
  const worst = [...rows.entries()].sort((p, q) => q[1] - p[1]).slice(0, 6)
  console.log('worst rows:', worst.map(([y, n]) => `${y}:${n}`).join(' '))
}
