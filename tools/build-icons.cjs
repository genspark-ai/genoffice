/**
 * Regenerate the app icon set from apps/shell/build/logo.svg.
 *
 *   node_modules/.bin/electron tools/build-icons.mjs
 *
 * Rasterizes the SVG master through an offscreen Chromium window (no native
 * deps) and packs:
 *   apps/shell/build/icon.png      1024 (Linux + electron-builder source)
 *   apps/shell/build/icon-mac.png  1024
 *   apps/shell/build/icon.ico      16/24/32/48/64/128/256 PNG-in-ICO
 *   apps/shell/build/icon.icns     128/256/512/1024 PNG chunks
 */
const { app, BrowserWindow } = require('electron')
const { readFileSync, writeFileSync } = require('node:fs')
const { join } = require('node:path')

const ROOT = join(__dirname, '..')
const BUILD = join(ROOT, 'apps', 'shell', 'build')
const SVG = readFileSync(join(BUILD, 'logo.svg'), 'utf8')

const ICO_SIZES = [16, 24, 32, 48, 64, 128, 256]
const ICNS = [
  ['ic07', 128],
  ['ic08', 256],
  ['ic09', 512],
  ['ic10', 1024],
]

const SVG_B64 = Buffer.from(SVG).toString('base64')

/** one offscreen window rasterizes every size: canvas resize is enough */
async function rasterizeAll(sizes) {
  const win = new BrowserWindow({
    width: 1024,
    height: 1024,
    show: false,
    webPreferences: { offscreen: true },
  })
  const html = `<!DOCTYPE html><html><body><canvas id="c"></canvas></body></html>`
  await win.loadURL(`data:text/html;base64,${Buffer.from(html).toString('base64')}`)
  const out = new Map()
  for (const size of sizes) {
    const dataUrl = await win.webContents.executeJavaScript(`(async () => {
      const img = new Image()
      img.src = 'data:image/svg+xml;base64,${SVG_B64}'
      await img.decode()
      const c = document.getElementById('c')
      c.width = ${size}; c.height = ${size}
      const ctx = c.getContext('2d')
      ctx.clearRect(0, 0, ${size}, ${size})
      ctx.drawImage(img, 0, 0, ${size}, ${size})
      return c.toDataURL('image/png')
    })()`)
    out.set(size, Buffer.from(dataUrl.split(',')[1], 'base64'))
  }
  win.destroy()
  return out
}

function packIco(pngs) {
  const header = Buffer.alloc(6)
  header.writeUInt16LE(0, 0)
  header.writeUInt16LE(1, 2)
  header.writeUInt16LE(pngs.length, 4)
  let offset = 6 + 16 * pngs.length
  const entries = []
  for (const { size, data } of pngs) {
    const e = Buffer.alloc(16)
    e.writeUInt8(size >= 256 ? 0 : size, 0)
    e.writeUInt8(size >= 256 ? 0 : size, 1)
    e.writeUInt8(0, 2)
    e.writeUInt8(0, 3)
    e.writeUInt16LE(1, 4)
    e.writeUInt16LE(32, 6)
    e.writeUInt32LE(data.length, 8)
    e.writeUInt32LE(offset, 12)
    entries.push(e)
    offset += data.length
  }
  return Buffer.concat([header, ...entries, ...pngs.map((p) => p.data)])
}

function packIcns(pngs) {
  const chunks = pngs.map(({ type, data }) => {
    const head = Buffer.alloc(8)
    head.write(type, 0, 'ascii')
    head.writeUInt32BE(data.length + 8, 4)
    return Buffer.concat([head, data])
  })
  const body = Buffer.concat(chunks)
  const head = Buffer.alloc(8)
  head.write('icns', 0, 'ascii')
  head.writeUInt32BE(body.length + 8, 4)
  return Buffer.concat([head, body])
}

app.whenReady().then(async () => {
  const sizes = [...new Set([1024, ...ICO_SIZES, ...ICNS.map(([, s]) => s)])]
  const pngs = await rasterizeAll(sizes)

  writeFileSync(join(BUILD, 'icon.png'), pngs.get(1024))
  writeFileSync(join(BUILD, 'icon-mac.png'), pngs.get(1024))

  const icoPngs = ICO_SIZES.map((size) => ({ size, data: pngs.get(size) }))
  writeFileSync(join(BUILD, 'icon.ico'), packIco(icoPngs))

  const icnsPngs = ICNS.map(([type, size]) => ({ type, data: pngs.get(size) }))
  writeFileSync(join(BUILD, 'icon.icns'), packIcns(icnsPngs))

  console.log('icons regenerated: icon.png, icon-mac.png, icon.ico, icon.icns')
  app.quit()
})
